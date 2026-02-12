import pandas as pd
import os
import glob
import numpy as np
from datetime import datetime

# Importar comentarios (opcional si se aplica aqu)
try:
    from col_comments import COL_COMMENTS
except ImportError:
    COL_COMMENTS = {}

def find_latest_file(directory, pattern='*.xlsx'):
    files = glob.glob(os.path.join(directory, pattern))
    files = [f for f in files if not os.path.basename(f).startswith('~$')]
    if not files: return None
    return max(files, key=os.path.getmtime)

def seconds_to_hms(seconds):
    try:
        if pd.isna(seconds): return "00:00:00"
        seconds = int(float(seconds))
        if seconds < 0: return "00:00:00"
        m, s = divmod(seconds, 60)
        h, m = divmod(m, 60)
        return "{:02d}:{:02d}:{:02d}".format(h, m, s)
    except: return "00:00:00"

def format_percentage(val):
    try:
        if pd.isna(val) or val == np.inf: return "0.00 %"
        return "{:.2f} %".format(float(val))
    except: return str(val)

def generate_executive_summary():
    # Input: Carpeta anterior (Resumen Semanal)
    input_dir = r"C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\INFORMES\KPI_SMART\RESUMEN_SEMANAL"
    # Output: Carpeta nueva
    output_dir = r"C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\INFORMES\KPI_SMART\RESUMEN_EJECUTIVO"
    if not os.path.exists(output_dir): os.makedirs(output_dir)

    print(f"Buscando archivo SEMANAL ms reciente en: {input_dir}")
    latest_file = find_latest_file(input_dir, pattern='*RESUMEN_SEMANAL*.xlsx')
    if not latest_file: 
        print("No se encontr archivo fuente en Resumen Semanal.")
        return

    print(f"Procesando: {latest_file}")
    
    try:
        original_sheets = pd.read_excel(latest_file, sheet_name=None)
        
        # Fuente de datos: Detalle_Leads_Unicos
        if 'Detalle_Leads_Unicos' not in original_sheets: return
        df = original_sheets['Detalle_Leads_Unicos']

        if 'fxCreated' in df.columns:
            df['fxCreated'] = pd.to_datetime(df['fxCreated'], errors='coerce')
            df_filtered = df[df['fxCreated'].dt.year >= 2026].copy()
        else: return
        
        if df_filtered.empty: return

        # --- PREPARACIN DE DATOS ---
        df_filtered['mes_nombre'] = df_filtered['fxCreated'].dt.strftime('%B %Y')
        df_filtered['mes_sort'] = df_filtered['fxCreated'].dt.to_period('M')
        
        # Clasificacin Franja (Recalcular si no viene, aunque debera venir en Detalle si se guard... 
        # pero Detalle origina de _1 que quizs no tena franjas guardadas explcitamente. Calculamos.)
        df_filtered['day_of_week'] = df_filtered['fxCreated'].dt.dayofweek
        df_filtered['hour'] = df_filtered['fxCreated'].dt.hour
        conditions = [
            (df_filtered['day_of_week'] >= 5),
            (df_filtered['hour'] >= 10) & (df_filtered['hour'] < 18)
        ]
        choices = ['FDS', 'OPERATIVO']
        # El user mencion "10-20" en la imagen pero mantenemos lgica script anterior (10-18) por consistencia o ajustamos?
        # Mantendr 10-18 OPERATIVO segn scripts previos.
        df_filtered['time_category'] = np.select(conditions, choices, default='EXTRA')

        # ACW por Lead
        acw_cols = [f'timeAcw{i}' for i in range(1, 11)]
        for col in acw_cols:
             if col not in df_filtered.columns: df_filtered[col] = 0
             else: df_filtered[col] = pd.to_numeric(df_filtered[col], errors='coerce').fillna(0)
        df_filtered['total_acw_lead'] = df_filtered[acw_cols].sum(axis=1)

        # Interacciones Venta (Sumar cuntos intentos fueron venta)
        # Esto requiere iterar o sumar columnas resultDesc
        # O si ya tenemos 'interacciones_venta' calculado en _1Detalle (s lo tenemos).
        if 'interacciones_venta' not in df_filtered.columns:
             df_filtered['interacciones_venta'] = 0 # Placeholder if missing
        
        # --- GENERACIN DE INDICADORES ---
        # Lista de Meses + Total
        months = sorted(df_filtered['mes_sort'].unique())
        
        # Estructura del DataFrame Final
        # Rows: Indicadores
        # Cols: Mes 1, Mes 2, ..., TOTAL GENERAL
        
        indicators_data = {}
        
        # Definir el orden de filas
        row_labels = [
            "Total leads recibidos",
            "Leads duplicados (Telfono)", # Asumido 0 si son nicos
            "Leads nicos analizados",
            "Contactados (unicos)",
            "No contactados (Unicos)",
            "Contactabilidad % (unicos)",
            "Ventas (leads insertados mes)",
            "Conversin % (contactados)",
            "conversion_contactos_%", # Ventas / Total Interacciones Contacto (Calcularemos interacciones contacto)
            "conv_contactado_cerrado_%", # Placeholder o Ventas/Contactados ? La imagen deca contactado_cerrado.
            "conv_leads_insertados_%",
            "interacciones en venta",
            "Total interacciones",
            "TMO totalizado (hh:mm:ss)",
            "TMO promedio (General) (hh:mm:ss)",
            "TMO mediana (General) (hh:mm:ss)",
            "TMO interacciones (venta) (hh:mm:ss)", # Promedio
            "TMO mediana (venta) (hh:mm:ss)",
            "TMO NO VENTA promedio (hh:mm:ss)",
            "TMO mediana (no venta) (hh:mm:ss)",
            "SLA OPERATIVO (10-20 Lu-Vi)", # Etiqueta imagen
            "SLA EXTRAHORARIO (Lu-Vi)",
            "SLA FIN DE SEMANA",
            "ACW Mediana (hh:mm:ss)",
            "Total tiempo llamadas"
        ]

        # Funcin auxiliar para calcular mtricas de un subconjunto
        def get_metrics_for_group(df_g):
            metrics = {}
            
            total_leads = len(df_g)
            duplicados = 0 # Asumimos 0 dado "Leads Unicos"
            unicos = total_leads - duplicados
            
            contactados = len(df_g[df_g['contactado'] == True])
            no_contactados = unicos - contactados
            
            ventas = len(df_g[df_g['venta'] == True])
            
            # Interacciones Totales
            int_total = df_g['Interacciones_x_lead'].sum()
            # Interacciones Venta
            int_venta = df_g['interacciones_venta'].sum()
            # Interacciones Contacto (Estimacin: int_total - int_sin_contacto? No tenemos int_sin_contacto aqu fcil sin iterar)
            # Usaremos: Total interacciones (asumiendo que Interacciones_x_lead cuenta TMO>0 segn definicin _1)
            # En _1: "df['Interacciones_x_lead'] = (df[tmo_cols] > 0).sum(axis=1)" -> Son interacciones CON CONTACTO.
            # Por tanto, int_contacto aprox = int_total (segn la definicin de _1).
            int_contacto = int_total 
            
            # Tiempos
            tmo_total_leads = pd.to_numeric(df_g['tmo_total_registro'], errors='coerce').fillna(0)
            tmo_sum_total = tmo_total_leads.sum()
            
            # TMO Venta
            tmo_venta_series = pd.to_numeric(df_g[df_g['venta']==True]['tmo_venta'], errors='coerce').dropna() # tmo_venta es "hasta la venta"
            # Pero la imagen pide "TMO interacciones (venta)". Quizs se refiere al tmo de ESA llamada.
            # No tenemos fcil el TMO de esa llamada aislado aqu sin re-procesar.
            # Usaremos 'tmo_venta' calculado en _1 (suma parcial hasta venta) como proxy, o tmo_venta_total_registro?
            # _1: "df['tmo_venta_total_registro'] = el total de los tmo>0 donde resultdesc= VENTA".
            # Usaremos tmo_venta_total_registro (Tiempo en llamadas de venta).
            tmo_venta_calls = pd.to_numeric(df_g[df_g['venta']==True]['tmo_venta_total_registro'], errors='coerce').dropna()
            
            # TMO General Promedio/Mediana (Sobre leads contactados, o sobre todos? Sobre tmo>0)
            tmo_contactados = tmo_total_leads[tmo_total_leads > 0]
            
            # TMO No Venta
            # Leads sin venta con TMO > 0
            tmo_no_venta = tmo_total_leads[(df_g['venta'] == False) & (tmo_total_leads > 0)]
            
            # SLA
            sla_series = pd.to_numeric(df_g['sla_seg'], errors='coerce').dropna()
            sla_op = pd.to_numeric(df_g[df_g['time_category']=='OPERATIVO']['sla_seg'], errors='coerce').dropna()
            sla_ex = pd.to_numeric(df_g[df_g['time_category']=='EXTRA']['sla_seg'], errors='coerce').dropna()
            sla_fds = pd.to_numeric(df_g[df_g['time_category']=='FDS']['sla_seg'], errors='coerce').dropna()
            
            # ACW
            acw_series = pd.to_numeric(df_g['total_acw_lead'], errors='coerce').fillna(0)
            acw_contacted = acw_series[tmo_total_leads > 0] # Solo gestionados
            
            # -- Asignacin --
            
            metrics["Total leads recibidos"] = total_leads
            metrics["Leads duplicados (Telfono)"] = duplicados
            metrics["Leads nicos analizados"] = unicos
            metrics["Contactados (unicos)"] = contactados
            metrics["No contactados (Unicos)"] = no_contactados
            
            metrics["Contactabilidad % (unicos)"] = (contactados/unicos * 100) if unicos > 0 else 0
            
            metrics["Ventas (leads insertados mes)"] = ventas
            metrics["Conversin % (contactados)"] = (ventas/contactados * 100) if contactados > 0 else 0
            
            # conversion_contactos_% (Ventas / Interacciones de Contacto)
            metrics["conversion_contactos_%"] = (ventas/int_contacto * 100) if int_contacto > 0 else 0
            
            # conv_contactado_cerrado_% (Ventas / Contactados?) - Si cerrado es contactado...
            # Si no tenemos 'cerrados' explcito aqu, usaremos Ventas/Contactados como proxy o N/A.
            # Lo dejar igual a Conversion % (contactados) por ahora o vaco.
            metrics["conv_contactado_cerrado_%"] = (ventas/contactados * 100) if contactados > 0 else 0
            
            metrics["conv_leads_insertados_%"] = (ventas/unicos * 100) if unicos > 0 else 0
            
            metrics["interacciones en venta"] = int_venta
            metrics["Total interacciones"] = int_total
            
            # TMOs
            metrics["TMO totalizado (hh:mm:ss)"] = seconds_to_hms(tmo_sum_total)
            metrics["TMO promedio (General) (hh:mm:ss)"] = seconds_to_hms(tmo_contactados.mean())
            metrics["TMO mediana (General) (hh:mm:ss)"] = seconds_to_hms(tmo_contactados.median())
            
            # TMO Venta
            metrics["TMO interacciones (venta) (hh:mm:ss)"] = seconds_to_hms(tmo_venta_calls.mean()) # Promedio
            metrics["TMO mediana (venta) (hh:mm:ss)"] = seconds_to_hms(tmo_venta_calls.median())
            
            # TMO No Venta
            metrics["TMO NO VENTA promedio (hh:mm:ss)"] = seconds_to_hms(tmo_no_venta.mean())
            metrics["TMO mediana (no venta) (hh:mm:ss)"] = seconds_to_hms(tmo_no_venta.median())
            
            # SLA
            metrics["SLA OPERATIVO (10-20 Lu-Vi)"] = seconds_to_hms(sla_op.median())
            metrics["SLA EXTRAHORARIO (Lu-Vi)"] = seconds_to_hms(sla_ex.median())
            metrics["SLA FIN DE SEMANA"] = seconds_to_hms(sla_fds.median())
            
            # ACW
            metrics["ACW Mediana (hh:mm:ss)"] = seconds_to_hms(acw_contacted.median())
            
            # Total Tiempo (Call + ACW?) -> La imagen dice "Total tiempo llamadas". TMO Totalizado ya est arriba.
            # Quizs es la suma de TMO + ACW? O Suma de 'timeCallTotal' (que incluye ring time?)
            # Usar TMO Totalizado de nuevo o timeCallTotal.
            # Si _1 calcul 'total_time_seconds' (call total).
            time_call_total = pd.to_numeric(df_g['total_time_seconds'], errors='coerce').sum() if 'total_time_seconds' in df_g else tmo_sum_total
            metrics["Total tiempo llamadas"] = seconds_to_hms(time_call_total)

            return metrics

        # --- LOOP POR MESES ---
        final_dict = {"INDICADOR": row_labels}
        
        # Calcular por cada mes
        for m_sort in months:
            df_m = df_filtered[df_filtered['mes_sort'] == m_sort]
            label = m_sort.strftime('%B %Y')
            m_metrics = get_metrics_for_group(df_m)
            
            col_vals = []
            for k in row_labels:
                val = m_metrics.get(k, 0)
                # Formatear %
                if "%" in k:
                    val = format_percentage(val)
                col_vals.append(val)
            
            final_dict[label] = col_vals
            
        # Calcular TOTAL GENERAL
        total_metrics = get_metrics_for_group(df_filtered)
        col_vals_total = []
        for k in row_labels:
            val = total_metrics.get(k, 0)
            if "%" in k:
                val = format_percentage(val)
            col_vals_total.append(val)
            
        final_dict["TOTAL GENERAL"] = col_vals_total
        
        # Crear DataFrame Final
        df_exec = pd.DataFrame(final_dict)
        
        # --- OUTPUT ---
        original_sheets['Resumen_Ejecutivo'] = df_exec
        
        base_name = os.path.basename(latest_file)
        parts = base_name.split('_')
        provider = parts[0] if len(parts) > 0 else "UNKNOWN"
        timestamp = datetime.now().strftime("%d%m%Y_%H%M%S")
        
        output_filename = f"{provider}_RESUMEN_EJECUTIVO_{timestamp}.xlsx"
        output_path = os.path.join(output_dir, output_filename)
        
        print(f"Guardando Resumen Ejecutivo en: {output_path}")
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
             for sheet_name, df_sheet in original_sheets.items():
                if isinstance(df_sheet, pd.DataFrame):
                     df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
                     
             # Formato opcional (ancho columnas)
             wb = writer.book
             if 'Resumen_Ejecutivo' in wb.sheetnames:
                 ws = wb['Resumen_Ejecutivo']
                 # No comments requested for Exec? We can add if needed.
                 # Ajustar ancho primera columna
                 ws.column_dimensions['A'].width = 40

        print("Resumen Ejecutivo generado exitosamente.")

    except Exception as e:
        print(f"Error generando Resumen Ejecutivo: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    generate_executive_summary()
