import pandas as pd
import os
import glob
import re
from datetime import datetime, timedelta
import numpy as np

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

def format_float_2dec(val):
    try:
        if pd.isna(val): return 0.00
        return round(float(val), 2)
    except: return val

def generate_weekly_summary():
    input_dir = r"C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\INFORMES\KPI_SMART\DETALLE_LEADS_UNICOS"
    output_dir = r"C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\INFORMES\KPI_SMART\RESUMEN_SEMANAL"
    if not os.path.exists(output_dir): os.makedirs(output_dir)

    print(f"Buscando archivo DETALLE ms reciente en: {input_dir}")
    latest_file = find_latest_file(input_dir, pattern='*DETALLE_LEADS_UNICOS*.xlsx')
    if not latest_file: return

    print(f"Procesando: {latest_file}")
    
    try:
        original_sheets = pd.read_excel(latest_file, sheet_name=None)
        
        if 'Detalle_Leads_Unicos' not in original_sheets: return
        df = original_sheets['Detalle_Leads_Unicos']

        if 'fxCreated' in df.columns:
            df['fxCreated'] = pd.to_datetime(df['fxCreated'], errors='coerce')
            df['mes'] = df['fxCreated'].dt.to_period('M') 
            df_filtered = df[df['fxCreated'].dt.year >= 2026].copy()
        else: return
        
        if df_filtered.empty: return

        # --- PREPARACIN SEMANAL ---
        # Agrupar por Semana (Lunes a Domingo)
        # dt.to_period('W-SUN') significa semana terminando en Domingo (empieza Lunes)
        # Esto nos dar un Period object que representa la semana.
        df_filtered['week_group'] = df_filtered['fxCreated'].dt.to_period('W-SUN')
        
        # Clasificacin Horaria (para SLA Franjas - Igual que diario)
        df_filtered['day_of_week'] = df_filtered['fxCreated'].dt.dayofweek
        df_filtered['hour'] = df_filtered['fxCreated'].dt.hour
        conditions = [
            (df_filtered['day_of_week'] >= 5),
            (df_filtered['hour'] >= 10) & (df_filtered['hour'] < 18)
        ]
        choices = ['FDS', 'OPERATIVO']
        df_filtered['time_category'] = np.select(conditions, choices, default='EXTRA')

        # --- CLCULOS GRUPALES (SEMANALES) ---
        # Usamos la misma lgica de clculo mtrico que para el diario
        summary_weekly = calculate_metrics(df_filtered.copy(), group_col='week_group')
        
        # --- Formato Etiqueta Semana ---
        # week_group es Period. .start_time y .end_time nos dan rango.
        # Format: YYYY-MM-DD/YYYY-MM-DD
        summary_weekly['fecha_inicio'] = summary_weekly['week_group'].apply(lambda x: x.start_time.date())
        summary_weekly['fecha_fin'] = summary_weekly['week_group'].apply(lambda x: x.end_time.date())
        summary_weekly['semana_label'] = summary_weekly.apply(
            lambda x: f"{x['fecha_inicio']}/{x['fecha_fin']}", axis=1
        )
        
        # Ordenar Descendente por la semana
        summary_weekly = summary_weekly.sort_values('week_group', ascending=False)
        
        # Asignar Mes para agrupacin visual (basado en fecha inicio de semana)
        summary_weekly['mes_sort'] = summary_weekly['week_group'].apply(lambda x: x.start_time.strftime('%Y-%m'))
        summary_weekly['mes_label'] = summary_weekly['week_group'].apply(lambda x: x.start_time.strftime('%B %Y').upper())

        # --- ESTRUCTURA FINAL HBRIDA (Con separadores mensuales) ---
        final_summary_list = []
        unique_months = summary_weekly['mes_sort'].unique()
        #unique_months = sorted(unique_months, reverse=True) # Ya ordenado por week_group desc
        
        # Definir columnas finales
        # Renombrar 'semana_label' a 'semana' (o 'fecha' como en diario, pero aqu es rango)
        final_cols_base = [
            'leads_insertados', 'contactados', 'ventas', 
            'interacciones_total', 'tiempo_total_llamadas_hms', 
            'contactabilidad_%', 'conversion_%',
            'mediana_tmo_x_periodo_seg', 'mediana_tmo_x_periodo_hms',
            'mediana_tmo_venta_x_periodo_seg', 'mediana_tmo_venta_x_periodo_hms',
            'mediana_sla_seg', 'mediana_sla_hms',
            'sla_operativo_mediana_hms', 'sla_extra_mediana_hms', 'sla_fds_mediana_hms',
            'timeAcw_mediana_dia_hms' # Renombrado a periodo luego
        ]
        
        # Generar lista final
        processed_months = set()
        
        # Iterar sobre el DF ordenado para ir inyectando cabeceras cuando cambie el mes
        current_month = None
        
        # Crear un DF vaco con la estructura final para append
        cols_display = ['semana'] + final_cols_base
        
        # Como iteramos de ms reciente a ms antiguo (desc), la primera fila es el mes ms reciente.
        # Agrupamos por mes para sacar totales mensuales tambin? 
        # En el ejemplo solo se ven filas de semana bajo el mes.
        # Pero en el Diario s haba totales mensuales. Asumimos lo mismo aqu.
        
        for mes_key in unique_months:
             # Datos de ese mes (semanas que inician en ese mes)
             df_month_weeks = summary_weekly[summary_weekly['mes_sort'] == mes_key]
             if df_month_weeks.empty: continue
             
             mes_nombre = df_month_weeks.iloc[0]['mes_label']
             
             # Cabecera Mes
             sep_row = {c: '' for c in cols_display}
             sep_row['semana'] = f"MES: {mes_nombre}"
             final_summary_list.append(pd.DataFrame([sep_row]))
             
             # Filas Semanales
             weeks_data = df_month_weeks.copy()
             weeks_data = weeks_data.rename(columns={'semana_label': 'semana'})
             # Filtrar solo columnas display
             weeks_data = weeks_data[cols_display]
             final_summary_list.append(weeks_data)
             
             # Total Mensual (Agregado de las semanas? O recalculado del source?)
             # Lo ideal es recalcular del source para que sea exacto (suma de leads ok, pero medianas no son suma de medianas).
             # Filtramos source data por el mes
             mes_period = pd.Period(mes_key, freq='M')
             df_month_source = df_filtered[df_filtered['fxCreated'].dt.to_period('M') == mes_period]
             
             if not df_month_source.empty:
                 df_month_source['dummy_group'] = 'Total Mes'
                 month_stats = calculate_metrics(df_month_source.copy(), group_col='dummy_group')
                 month_stats['semana'] = "TOTAL MES"
                 month_stats = month_stats[['semana'] + final_cols_base]
                 final_summary_list.append(month_stats)
                 
             # Espacio
             empty_row = {c: '' for c in cols_display}
             final_summary_list.append(pd.DataFrame([empty_row]))
             
        final_df = pd.concat(final_summary_list, ignore_index=True)

        # --- GUARDAR ---
        original_sheets['Resumen_Semanal'] = final_df
        
        base_name = os.path.basename(latest_file)
        parts = base_name.split('_')
        provider = parts[0] if len(parts) > 0 else "UNKNOWN"
        timestamp = datetime.now().strftime("%d%m%Y_%H%M%S")
        
        output_filename = f"{provider}_RESUMEN_SEMANAL_{timestamp}.xlsx"
        output_path = os.path.join(output_dir, output_filename)
        
        print(f"Guardando Resumen Semanal en: {output_path}")
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
             for sheet_name, df_sheet in original_sheets.items():
                if isinstance(df_sheet, pd.DataFrame):
                     df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
                     
             wb = writer.book
             if 'Resumen_Semanal' in wb.sheetnames:
                 ws = wb['Resumen_Semanal']
                 # Reusar comentarios de Resumen_Diario (mismas metricas, solo cambia periodo)
                 apply_comments(ws, final_df.columns, "Resumen_Diario") # Usamos config de Diario

        print("Resumen Semanal generado exitosamente.")

    except Exception as e:
        print(f"Error generando Resumen Semanal: {e}")
        import traceback
        traceback.print_exc()

def calculate_metrics(df_grouper, group_col):
    if df_grouper.empty:
         cols = [
            group_col,
            'leads_insertados', 'contactados', 'ventas', 
            'interacciones_total', 'tiempo_total_llamadas_hms', 
            'contactabilidad_%', 'conversion_%',
            'mediana_tmo_x_periodo_seg', 'mediana_tmo_x_periodo_hms',
            'mediana_tmo_venta_x_periodo_seg', 'mediana_tmo_venta_x_periodo_hms',
            'mediana_sla_seg', 'mediana_sla_hms',
            'sla_operativo_mediana_hms', 'sla_extra_mediana_hms', 'sla_fds_mediana_hms',
            'timeAcw_mediana_dia_hms'
         ]
         return pd.DataFrame(columns=cols)

    grouped = df_grouper.groupby(group_col).size().reset_index(name='leads_insertados')
    contactados = df_grouper[df_grouper['contactado'] == True].groupby(group_col).size().reset_index(name='contactados')
    ventas = df_grouper[df_grouper['venta'] == True].groupby(group_col).size().reset_index(name='ventas')
    interacciones = df_grouper.groupby(group_col)['Interacciones_x_lead'].sum().reset_index(name='interacciones_total')
    
    time_cols = [f'timeCall{i}' for i in range(1, 11)]
    for col in time_cols:
         if col not in df_grouper.columns: df_grouper[col] = 0
         else: df_grouper[col] = pd.to_numeric(df_grouper[col], errors='coerce').fillna(0)
    
    df_grouper['total_seconds_call'] = df_grouper[time_cols].sum(axis=1)
    tiempo_total = df_grouper.groupby(group_col)['total_seconds_call'].sum().reset_index(name='tiempo_total_llamadas_seg')
    
    # Medianas
    df_grouper['tmo_total_registro'] = pd.to_numeric(df_grouper['tmo_total_registro'], errors='coerce').fillna(0)
    tmo_mediana = df_grouper[df_grouper['tmo_total_registro'] > 0].groupby(group_col)['tmo_total_registro'].median().reset_index(name='mediana_tmo_x_periodo_seg')
    
    df_grouper['tmo_venta'] = pd.to_numeric(df_grouper['tmo_venta'], errors='coerce').fillna(0)
    tmo_venta_mediana = df_grouper[df_grouper['venta'] == True].groupby(group_col)['tmo_venta'].median().reset_index(name='mediana_tmo_venta_x_periodo_seg')
    
    df_grouper['sla_seg'] = pd.to_numeric(df_grouper['sla_seg'], errors='coerce')
    sla_mediana = df_grouper.groupby(group_col)['sla_seg'].median().reset_index(name='mediana_sla_seg')
    
    # ACW Mediana Lead
    acw_cols = [f'timeAcw{i}' for i in range(1, 11)]
    for col in acw_cols:
         if col not in df_grouper.columns: df_grouper[col] = 0
         else: df_grouper[col] = pd.to_numeric(df_grouper[col], errors='coerce').fillna(0)
    df_grouper['total_acw_lead'] = df_grouper[acw_cols].sum(axis=1)
    acw_mediana = df_grouper[df_grouper['tmo_total_registro'] > 0].groupby(group_col)['total_acw_lead'].median().reset_index(name='timeAcw_mediana_dia_seg')
    
    # SLA Franjas
    sla_op = df_grouper[df_grouper['time_category'] == 'OPERATIVO'].groupby(group_col)['sla_seg'].median().reset_index(name='sla_operativo_mediana_seg')
    sla_ex = df_grouper[df_grouper['time_category'] == 'EXTRA'].groupby(group_col)['sla_seg'].median().reset_index(name='sla_extra_mediana_seg')
    sla_fds = df_grouper[df_grouper['time_category'] == 'FDS'].groupby(group_col)['sla_seg'].median().reset_index(name='sla_fds_mediana_seg')

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
    
    # Calcs
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
    
    hms_cols = [c for c in summary.columns if 'hms' in c]
    for c in hms_cols: summary[c] = summary[c].fillna("00:00:00")
    
    final_cols = [
        group_col, 'leads_insertados', 'contactados', 'ventas', 
        'interacciones_total', 'tiempo_total_llamadas_hms', 
        'contactabilidad_%', 'conversion_%',
        'mediana_tmo_x_periodo_seg', 'mediana_tmo_x_periodo_hms',
        'mediana_tmo_venta_x_periodo_seg', 'mediana_tmo_venta_x_periodo_hms',
        'mediana_sla_seg', 'mediana_sla_hms',
        'sla_operativo_mediana_hms', 'sla_extra_mediana_hms', 'sla_fds_mediana_hms',
        'timeAcw_mediana_dia_hms'
    ]
    
    # Asegurar existencia
    for c in final_cols: 
        if c not in summary.columns: summary[c] = 0 if 'hms' not in c else "00:00:00"
        
    return summary[final_cols]

def apply_comments(worksheet, columns, sheet_config_name):
    from openpyxl.comments import Comment
    if sheet_config_name not in COL_COMMENTS:
         # Fallback
         for k in COL_COMMENTS.keys():
             if k in sheet_config_name:
                 sheet_config_name = k
                 break
    if sheet_config_name not in COL_COMMENTS: return

    comments_dict = COL_COMMENTS[sheet_config_name]
    
    for col_idx, col_name in enumerate(columns, 1):
        cell = worksheet.cell(row=1, column=col_idx)
        comment_text = None
        if col_name in comments_dict:
            comment_text = comments_dict[col_name]
        else:
             clean_name = col_name.replace('periodo', 'dia').replace('total', 'dia')
             if clean_name in comments_dict:
                 comment_text = comments_dict[clean_name]

        if comment_text:
            cell.comment = Comment(comment_text, "System")

if __name__ == "__main__":
    generate_weekly_summary()
