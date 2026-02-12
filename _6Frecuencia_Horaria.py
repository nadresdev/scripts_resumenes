import pandas as pd
import os
import glob
import numpy as np
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.comments import Comment

# Importar diccionario de comentarios (si existe, o definir aqui)
# Usaremos un diccionario base similar al diario
FRECUENCIA_COMMENTS = {
    "hora_franja": "Franja horaria basada en fxCreated",
    "leads_insertados": "Total leads recibidos en la franja",
    "contactados": "Total leads con al menos 1 TMO > 0",
    "ventas": "Total leads con venta=SI",
    "interacciones_total": "Suma de interacciones (contacto + intentos)",
    "tiempo_total_llamadas_hms": "Suma de tiempo de llamada total",
    "contactabilidad_%": "Contactados / Leads * 100",
    "conversion_%": "Ventas / Leads * 100",
    "mediana_tmo_x_periodo_hms": "Mediana TMO General (TMO>0)",
    "mediana_tmo_venta_x_periodo_hms": "Mediana TMO Ventas",
    "mediana_sla_hms": "Mediana SLA Global",
    "sla_operativo_mediana_hms": "Mediana SLA (10-18h L-V)",
    "sla_extra_mediana_hms": "Mediana SLA (Extra Horario)",
    "sla_fds_mediana_hms": "Mediana SLA (FDS)",
    "timeAcw_mediana_dia_hms": "Mediana ACW por lead"
}

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
    
def format_float_2dec(val):
    try:
        if pd.isna(val): return 0.00
        return round(float(val), 2)
    except: return val

def calculate_metrics(df_grouper, group_col):
    if df_grouper.empty:
         cols = [group_col, 'leads_insertados', 'contactados', 'ventas', 'interacciones_total', 
                 'tiempo_total_llamadas_hms', 'contactabilidad_%', 'conversion_%',
                 'mediana_tmo_x_periodo_seg', 'mediana_tmo_x_periodo_hms', 
                 'mediana_tmo_venta_x_periodo_seg', 'mediana_tmo_venta_x_periodo_hms',
                 'mediana_sla_seg', 'mediana_sla_hms',
                 'sla_operativo_mediana_hms', 'sla_extra_mediana_hms', 'sla_fds_mediana_hms',
                 'timeAcw_mediana_dia_hms']
         return pd.DataFrame(columns=cols)

    grouped = df_grouper.groupby(group_col).size().reset_index(name='leads_insertados')
    
    # Asegurar columnas booleanas/numericas
    if 'tmo_total_registro' not in df_grouper.columns: df_grouper['tmo_total_registro'] = 0
    else: df_grouper['tmo_total_registro'] = pd.to_numeric(df_grouper['tmo_total_registro'], errors='coerce').fillna(0)
    
    contactados = df_grouper[df_grouper['tmo_total_registro'] > 0].groupby(group_col).size().reset_index(name='contactados')
    
    if 'venta' in df_grouper.columns:
        ventas = df_grouper[df_grouper['venta'] == True].groupby(group_col).size().reset_index(name='ventas')
    else:
        ventas = pd.DataFrame(columns=[group_col, 'ventas'])
        
    if 'Interacciones_x_lead' in df_grouper.columns:
        interacciones = df_grouper.groupby(group_col)['Interacciones_x_lead'].sum().reset_index(name='interacciones_total')
    else:
        interacciones = pd.DataFrame(columns=[group_col, 'interacciones_total'])
    
    # Tiempo Total Llamadas (Sum tmo_contact or timeTotalCall column)
    # Preferimos sumar tmos si timeCallTotal no es confiable, pero usaremos tmo_total_registro sumado
    tiempo_total = df_grouper.groupby(group_col)['tmo_total_registro'].sum().reset_index(name='tiempo_total_llamadas_seg')
    
    # --- Medianas ---
    tmo_mediana = df_grouper[df_grouper['tmo_total_registro'] > 0].groupby(group_col)['tmo_total_registro'].median().reset_index(name='mediana_tmo_x_periodo_seg')
    
    if 'tmo_venta' in df_grouper.columns:
        df_grouper['tmo_venta'] = pd.to_numeric(df_grouper['tmo_venta'], errors='coerce').fillna(0)
        tmo_venta_mediana = df_grouper[df_grouper['venta'] == True].groupby(group_col)['tmo_venta'].median().reset_index(name='mediana_tmo_venta_x_periodo_seg')
    else:
         tmo_venta_mediana = pd.DataFrame(columns=[group_col, 'mediana_tmo_venta_x_periodo_seg'])
    
    if 'sla_seg' in df_grouper.columns:
         df_grouper['sla_seg'] = pd.to_numeric(df_grouper['sla_seg'], errors='coerce')
         sla_mediana = df_grouper.groupby(group_col)['sla_seg'].median().reset_index(name='mediana_sla_seg')
    else:
         sla_mediana = pd.DataFrame(columns=[group_col, 'mediana_sla_seg'])

    # ACW
    acw_cols = [f'timeAcw{i}' for i in range(1, 11)]
    for col in acw_cols:
         if col not in df_grouper.columns: df_grouper[col] = 0
         else: df_grouper[col] = pd.to_numeric(df_grouper[col], errors='coerce').fillna(0)
    df_grouper['total_acw_lead'] = df_grouper[acw_cols].sum(axis=1)
    # Filter ACW for managed leads only? Similar to Daily logic:
    acw_mediana = df_grouper[df_grouper['tmo_total_registro'] > 0].groupby(group_col)['total_acw_lead'].median().reset_index(name='timeAcw_mediana_dia_seg')

    # SLA Franjas
    # Necesitamos time_category
    day_of_week = df_grouper['fxCreated'].dt.dayofweek
    hour = df_grouper['fxCreated'].dt.hour
    conditions = [
        (day_of_week >= 5), 
        (hour >= 10) & (hour < 18)
    ]
    choices = ['FDS', 'OPERATIVO']
    df_grouper['time_category'] = np.select(conditions, choices, default='EXTRA')

    df_op = df_grouper[df_grouper['time_category'] == 'OPERATIVO']
    sla_op = df_op.groupby(group_col)['sla_seg'].median().reset_index(name='sla_operativo_mediana_seg') if not df_op.empty else pd.DataFrame(columns=[group_col, 'sla_operativo_mediana_seg'])
    
    df_ex = df_grouper[df_grouper['time_category'] == 'EXTRA']
    sla_ex = df_ex.groupby(group_col)['sla_seg'].median().reset_index(name='sla_extra_mediana_seg') if not df_ex.empty else pd.DataFrame(columns=[group_col, 'sla_extra_mediana_seg'])
    
    df_fds = df_grouper[df_grouper['time_category'] == 'FDS']
    sla_fds = df_fds.groupby(group_col)['sla_seg'].median().reset_index(name='sla_fds_mediana_seg') if not df_fds.empty else pd.DataFrame(columns=[group_col, 'sla_fds_mediana_seg'])
    
    # Merge Final
    summary = grouped
    summary = summary.merge(contactados, on=group_col, how='left').fillna({'contactados': 0})
    summary = summary.merge(ventas, on=group_col, how='left').fillna({'ventas': 0})
    summary = summary.merge(interacciones, on=group_col, how='left').fillna({'interacciones_total': 0})
    summary = summary.merge(tiempo_total, on=group_col, how='left').fillna({'tiempo_total_llamadas_seg': 0})
    summary = summary.merge(tmo_mediana, on=group_col, how='left')
    summary = summary.merge(tmo_venta_mediana, on=group_col, how='left')
    summary = summary.merge(sla_mediana, on=group_col, how='left')
    summary = summary.merge(acw_mediana, on=group_col, how='left')
    summary = summary.merge(sla_op, on=group_col, how='left')
    summary = summary.merge(sla_ex, on=group_col, how='left')
    summary = summary.merge(sla_fds, on=group_col, how='left')
    
    # Calculos Finales
    summary['contactabilidad_%'] = (summary['contactados'] / summary['leads_insertados']) * 100
    summary['conversion_%'] = (summary['ventas'] / summary['leads_insertados']) * 100
    
    summary['contactabilidad_%'] = summary['contactabilidad_%'].apply(format_percentage)
    summary['conversion_%'] = summary['conversion_%'].apply(format_percentage)
    
    float_cols = ['mediana_tmo_x_periodo_seg', 'mediana_tmo_venta_x_periodo_seg', 'mediana_sla_seg']
    for col in float_cols:
         if col in summary.columns: summary[col] = summary[col].apply(format_float_2dec)
         
    summary['tiempo_total_llamadas_hms'] = summary['tiempo_total_llamadas_seg'].apply(seconds_to_hms)
    summary['mediana_tmo_x_periodo_hms'] = summary['mediana_tmo_x_periodo_seg'].apply(seconds_to_hms)
    summary['mediana_tmo_venta_x_periodo_hms'] = summary['mediana_tmo_venta_x_periodo_seg'].apply(seconds_to_hms)
    summary['mediana_sla_hms'] = summary['mediana_sla_seg'].apply(seconds_to_hms)
    summary['timeAcw_mediana_dia_hms'] = summary['timeAcw_mediana_dia_seg'].apply(seconds_to_hms)
    
    summary['sla_operativo_mediana_hms'] = summary['sla_operativo_mediana_seg'].apply(seconds_to_hms)
    summary['sla_extra_mediana_hms'] = summary['sla_extra_mediana_seg'].apply(seconds_to_hms)
    summary['sla_fds_mediana_hms'] = summary['sla_fds_mediana_seg'].apply(seconds_to_hms)
    
    hms_cols = ['tiempo_total_llamadas_hms', 'mediana_tmo_x_periodo_hms', 'mediana_tmo_venta_x_periodo_hms', 
                'mediana_sla_hms', 'sla_operativo_mediana_hms', 'sla_extra_mediana_hms', 'sla_fds_mediana_hms',
                'timeAcw_mediana_dia_hms']
    for col in hms_cols:
         summary[col] = summary[col].fillna("00:00:00")
         
    final_cols = [
        group_col, 'leads_insertados', 'contactados', 'ventas', 
        'interacciones_total', 'tiempo_total_llamadas_hms', 
        'contactabilidad_%', 'conversion_%',
        'mediana_tmo_x_periodo_hms',
        'mediana_tmo_venta_x_periodo_hms',
        'mediana_sla_hms', 'sla_operativo_mediana_hms', 'sla_extra_mediana_hms', 'sla_fds_mediana_hms',
        'timeAcw_mediana_dia_hms'
    ]
    # Filter columns that exist
    final_cols = [c for c in final_cols if c in summary.columns]
    return summary[final_cols]

def generate_frecuencia_report():
    input_dir = r"C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\INFORMES\KPI_SMART\RESUMEN_EJECUTIVO"
    
    print(f"Buscando archivo EJECUTIVO ms reciente en: {input_dir}")
    latest_file = find_latest_file(input_dir, pattern='*_RESUMEN_EJECUTIVO_*.xlsx')
    
    if not latest_file:
        print("No se encontr archivo Resumen Ejecutivo.")
        return

    print(f"Procesando: {latest_file}")
    
    try:
        # Cargar con pandas
        original_sheets = pd.read_excel(latest_file, sheet_name=None)
        
        if 'Detalle_Leads_Unicos' not in original_sheets:
            print("No se encontr Detalle_Leads_Unicos en el archivo.")
            return
            
        df = original_sheets['Detalle_Leads_Unicos']
        
        if 'fxCreated' in df.columns:
            df['fxCreated'] = pd.to_datetime(df['fxCreated'], errors='coerce')
            df_filtered = df[df['fxCreated'].dt.year >= 2026].copy()
        else:
            print("No fxCreated column.")
            return
            
        if df_filtered.empty:
            print("No data 2026+.")
            return
            
        # Crear columna Franja Horaria
        df_filtered['hour_int'] = df_filtered['fxCreated'].apply(lambda x: f"{x.hour:02d}-{x.hour+1:02d}" if pd.notnull(x) else "Unknown")
        df_filtered['mes_sort'] = df_filtered['fxCreated'].dt.to_period('M')

        # --- Iterar Por Mes ---
        unique_months = sorted(df_filtered['mes_sort'].unique(), reverse=True)
        final_rows_list = []
        
        # Obtener columnas base de structure
        dummy_df = calculate_metrics(df_filtered.head(1), 'hour_int')
        dummy_df.rename(columns={'hour_int': 'hora_franja'}, inplace=True)
        
        # DEFINIR ORDEN EXPLICITO COLUMNAS
        # hora_franja PRIMERO
        cols = dummy_df.columns.tolist()
        if 'hora_franja' in cols:
            cols.remove('hora_franja')
            cols.insert(0, 'hora_franja')
        base_columns = cols
        
        for mes in unique_months:
            pass # Solo para el loop
            
            # Data Mes
            df_month = df_filtered[df_filtered['mes_sort'] == mes].copy()
            
            if df_month.empty: continue
            
            mes_label = mes.strftime('%B %Y').upper()
            
            # Separator Row
            separator_row = {col: '' for col in base_columns}
            separator_row['hora_franja'] = f"MES: {mes_label}" 
            final_rows_list.append(pd.DataFrame([separator_row]))
            
            # Group By Hour
            hourly_stats = calculate_metrics(df_month, group_col='hour_int')
            if not hourly_stats.empty:
                 hourly_stats.rename(columns={'hour_int': 'hora_franja'}, inplace=True)
                 hourly_stats = hourly_stats.sort_values('hora_franja')
                 hourly_stats = hourly_stats.reindex(columns=base_columns, fill_value=0)
                 final_rows_list.append(hourly_stats)
            
            # Total Mes Row
            df_month['dummy_group'] = 'TOTAL MES'
            month_total = calculate_metrics(df_month, group_col='dummy_group')
            month_total.rename(columns={'dummy_group': 'hora_franja'}, inplace=True)
            month_total['hora_franja'] = "TOTAL MES"
            month_total = month_total.reindex(columns=base_columns, fill_value=0)
            final_rows_list.append(month_total)
            
            # Empty Row separator
            empty_row = {col: '' for col in base_columns}
            final_rows_list.append(pd.DataFrame([empty_row]))
            
        final_df = pd.concat(final_rows_list, ignore_index=True)
        
        # PREPARAR NUEVO ARCHIVO
        # Generar nombre nuevo timestamps
        base_name = os.path.basename(latest_file)
        # Intentar preservar parte del nombre original o simplemente generar uno nuevo
        # Asumimos formato PROVIDER_RESUMEN_EJECUTIVO_FECHA_HORA.xlsx
        parts = base_name.split('_RESUMEN_EJECUTIVO_')
        provider = parts[0] if len(parts) > 0 else "UNKNOWN"
        
        timestamp = datetime.now().strftime("%d%m%Y_%H%M%S")
        
        # NUEVA RUTA DE SALIDA: KPI_SMART/FRECUENCIA
        # El input estaba en RESUMEN_EJECUTIVO, subimos un nivel y entramos a FRECUENCIA
        base_dir = os.path.dirname(input_dir) # KPI_SMART
        output_dir = os.path.join(base_dir, "FRECUENCIA")
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            
        # NOMBRE DE ARCHIVO: PROVIDER_FRECUENCIA_TIMESTAMP.xlsx (Patron solicitado)
        # User dijo: "guardarlo en FRECUENCIA CON EL PATRON DE NOMBRE CORRESPONDIENTE"
        # Asumo que se refiere a mantener el patrÓn del archivo origen pero en la carpeta nueva?
        # O quiza PROVIDER_FRECUENCIA_...
        # Voy a usar PROVIDER_FRECUENCIA_... para diferenciarlo claramente del Ejecutivo base.
        output_filename = f"{provider}_FRECUENCIA_{timestamp}.xlsx"
        output_path = os.path.join(output_dir, output_filename)
        
        print(f"Guardando NUEVO archivo con Frecuencia en: {output_path}")

        # Guardar todas las hojas + Frecuencia
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
             # Copiar hojas originales
             for sheet_name, df_sheet in original_sheets.items():
                 # Evitar duplicar Frecuencia si ya exista (la estamos regenerando)
                 if sheet_name != 'Frecuencia':
                      df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
             
             # Agregar Frecuencia agregada
             final_df.to_excel(writer, sheet_name='Frecuencia', index=False)
             
             # Estilos Basicos Frecuencia (Ancho)
             wb = writer.book
             ws = wb['Frecuencia']
             for col_idx, col_name in enumerate(final_df.columns, 1):
                 cell = ws.cell(row=1, column=col_idx)
                 if col_name in FRECUENCIA_COMMENTS:
                     cell.comment = Comment(FRECUENCIA_COMMENTS[col_name], "System")
                 ws.column_dimensions[cell.column_letter].width = 20

        print("Archivo generado exitosamente en carpeta FRECUENCIA.")

    except Exception as e:
        print(f"Error generando reporte Frecuencia: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    generate_frecuencia_report()
