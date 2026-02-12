import pandas as pd
import os
import glob
import re
from datetime import datetime
import numpy as np

# Importar comentarios
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

def generate_agent_summary():
    input_dir = r"C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\INFORMES\KPI_SMART\DETALLE_LEADS_UNICOS"
    output_dir = r"C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\INFORMES\KPI_SMART\RESUMEN_AGENTES"
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
            df['mes'] = df['fxCreated'].dt.to_period('M') # Para agrupar
            df_filtered = df[df['fxCreated'].dt.year >= 2026].copy()
        else: return
        
        if df_filtered.empty: return

        # --- PREPARACIN ---
        interactions = []
        df_filtered['day_of_week'] = df_filtered['fxCreated'].dt.dayofweek
        df_filtered['hour'] = df_filtered['fxCreated'].dt.hour
        conditions = [
            (df_filtered['day_of_week'] >= 5),
            (df_filtered['hour'] >= 10) & (df_filtered['hour'] < 18)
        ]
        choices = ['FDS', 'OPERATIVO']
        df_filtered['time_category'] = np.select(conditions, choices, default='EXTRA')

        for idx, row in df_filtered.iterrows():
            lead_sla = row.get('sla_seg', 0)
            lead_time_cat = row.get('time_category', 'EXTRA')
            lead_last_agent = str(row.get('lastOcmAgent', '')).strip()
            lead_mes = row.get('mes') # Period object

            for i in range(1, 11):
                agent = str(row.get(f'callAgent{i}', '')).strip()
                if not agent or agent.lower() in ['nan', '0', '']: continue
                
                tmo = pd.to_numeric(row.get(f'tmo{i}', 0), errors='coerce') or 0
                acw = pd.to_numeric(row.get(f'timeAcw{i}', 0), errors='coerce') or 0
                res_desc = str(row.get(f'resultDesc{i}', '')).upper()
                
                is_contact = tmo > 0
                is_sale = 'VENTA' in res_desc or 'POLIZA' in res_desc
                sla_val = lead_sla if i == 10 else np.nan
                sla_cat = lead_time_cat if i == 10 else None
                is_closer = (agent == lead_last_agent)
                
                interactions.append({
                    'agente': agent,
                    'mes': lead_mes, # aadido para agrupar
                    'tmo': tmo,
                    'acw': acw,
                    'is_contact': is_contact,
                    'is_sale': is_sale,
                    'sla': sla_val,
                    'sla_cat': sla_cat,
                    'lead_id': row.get('_id', idx), 
                    'is_closer': is_closer
                })
                
        df_int = pd.DataFrame(interactions)
        if df_int.empty: return

        # --- GENERACIN REPORTES POR MES ---
        
        # Obtener lista de meses nicos ordenados desc
        unique_months = df_int['mes'].dropna().unique()
        unique_months = sorted(unique_months, reverse=True)
        
        final_summary_list = []
        
        # Calcular columnas requeridas (dummy vaco para usar sus nombres luego)
        # Hacemos una pasada rpida global para obtener la estructura de columnas final
        # O definimos manualmente las columnas
        final_cols_order = [
            'agente', 'leads_cerrados', 'int_contacto', 'int_sin_contacto', 
            'interacciones_ventas', 'interacciones_total', 'ventas', 
            'conversion_contactos_%', 'conv_contactado_cerrado_%', 
            'tmo_total_hms', 'tmo_total_mediana_hms', 'tmo_ventas_hms', 'tmo_ventas_mediana_hms',
            'tiempo_total_llamadas_hms', 
            'sla_hms_medio', 'sla_mediana_hms', 
            'sla_operativo_mediana_hms', 'sla_extra_mediana_hms', 'sla_fds_mediana_hms', 
            'time_acw_hms', 'acw_mediana_hms', 
            'tmo_no_venta_hms', 'tmo_no_venta_mediana_hms'
        ]

        # --- BUCLE MESES ---
        for mes in unique_months:
            # Filtrar datos del mes
            df_mes = df_int[df_int['mes'] == mes]
            
            # Encabezado Mes
            mes_label = mes.strftime('%B %Y').upper()
            separator_row = {col: '' for col in final_cols_order}
            separator_row['agente'] = f"MES: {mes_label}"
            final_summary_list.append(pd.DataFrame([separator_row]))
            
            # Calcular Mtricas Agrupadas por Agente para este Mes
            if not df_mes.empty:
                agent_metrics = calculate_agent_metrics(df_mes, final_cols_order)
                final_summary_list.append(agent_metrics)
            
            # Espacio
            empty_row = {col: '' for col in final_cols_order}
            final_summary_list.append(pd.DataFrame([empty_row]))
            
        # --- TOTAL GLOBAL (Opcional, si se desea al final) ---
        # Si se desea un total global de todos los meses juntos:
        # separator_row['agente'] = "RESUMEN TOTAL"
        # final_summary_list.append(pd.DataFrame([separator_row]))
        # global_metrics = calculate_agent_metrics(df_int, final_cols_order)
        # final_summary_list.append(global_metrics)
        
        final_df = pd.concat(final_summary_list, ignore_index=True)
        
        # --- OUTPUT ---
        original_sheets['Resumen_Agentes'] = final_df
        
        base_name = os.path.basename(latest_file)
        parts = base_name.split('_')
        provider = parts[0] if len(parts) > 0 else "UNKNOWN"
        timestamp = datetime.now().strftime("%d%m%Y_%H%M%S")
        
        output_filename = f"{provider}_RESUMEN_AGENTES_{timestamp}.xlsx"
        output_path = os.path.join(output_dir, output_filename)
        
        print(f"Guardando Resumen Agentes en: {output_path}")
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
             for sheet_name, df_sheet in original_sheets.items():
                if isinstance(df_sheet, pd.DataFrame):
                     df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
                     
             wb = writer.book
             if 'Resumen_Agentes' in wb.sheetnames:
                 ws = wb['Resumen_Agentes']
                 apply_comments(ws, final_df.columns, "Resumen_Agentes")

        print("Resumen de Agentes generado exitosamente.")

    except Exception as e:
        print(f"Error generando Resumen Agentes: {e}")
        import traceback
        traceback.print_exc()

def calculate_agent_metrics(df_int, final_cols_order):
    closers = df_int[df_int['is_closer'] == True].groupby('agente')['lead_id'].nunique().reset_index(name='leads_cerrados')
    contacts = df_int[df_int['is_contact'] == True].groupby('agente').size().reset_index(name='int_contacto')
    no_contacts = df_int[df_int['is_contact'] == False].groupby('agente').size().reset_index(name='int_sin_contacto')
    sales_int = df_int[df_int['is_sale'] == True].groupby('agente').size().reset_index(name='interacciones_ventas')
    ventas = df_int[df_int['is_sale'] == True].groupby('agente').size().reset_index(name='ventas')
    total_int = df_int.groupby('agente').size().reset_index(name='interacciones_total')
    
    tmo_sum = df_int.groupby('agente')['tmo'].sum().reset_index(name='tmo_total_seg')
    tmo_med = df_int[df_int['tmo'] > 0].groupby('agente')['tmo'].median().reset_index(name='tmo_total_mediana_seg')
    
    tmo_v_sum = df_int[df_int['is_sale'] == True].groupby('agente')['tmo'].sum().reset_index(name='tmo_ventas_seg')
    tmo_v_med = df_int[df_int['is_sale'] == True].groupby('agente')['tmo'].median().reset_index(name='tmo_ventas_mediana_seg')
    
    tmo_nv_sum = df_int[(df_int['is_sale'] == False) & (df_int['tmo'] > 0)].groupby('agente')['tmo'].sum().reset_index(name='tmo_no_venta_seg')
    tmo_nv_med = df_int[(df_int['is_sale'] == False) & (df_int['tmo'] > 0)].groupby('agente')['tmo'].median().reset_index(name='tmo_no_venta_mediana_seg')
    
    acw_sum = df_int.groupby('agente')['acw'].sum().reset_index(name='time_acw_seg')
    acw_med = df_int.groupby('agente')['acw'].median().reset_index(name='acw_mediana_seg')
    
    sla_mean = df_int.groupby('agente')['sla'].mean().reset_index(name='sla_medio_seg')
    sla_med = df_int.groupby('agente')['sla'].median().reset_index(name='sla_mediana_seg')
    
    sla_op = df_int[df_int['sla_cat'] == 'OPERATIVO'].groupby('agente')['sla'].median().reset_index(name='sla_operativo_mediana_seg')
    sla_ex = df_int[df_int['sla_cat'] == 'EXTRA'].groupby('agente')['sla'].median().reset_index(name='sla_extra_mediana_seg')
    sla_fds = df_int[df_int['sla_cat'] == 'FDS'].groupby('agente')['sla'].median().reset_index(name='sla_fds_mediana_seg')

    summary = total_int
    dfs = [closers, contacts, no_contacts, sales_int, ventas, tmo_sum, tmo_med, tmo_v_sum, tmo_v_med, 
           tmo_nv_sum, tmo_nv_med, acw_sum, acw_med, sla_mean, sla_med, sla_op, sla_ex, sla_fds]
    for d in dfs: summary = summary.merge(d, on='agente', how='left')
    
    # Fill 0
    numeric_cols_zero = ['leads_cerrados', 'int_contacto', 'int_sin_contacto', 'interacciones_ventas', 'ventas', 
                         'tmo_total_seg', 'tmo_ventas_seg', 'tmo_no_venta_seg', 'time_acw_seg']
    for c in numeric_cols_zero: 
        if c in summary.columns: summary[c] = summary[c].fillna(0)
        
    summary['conversion_contactos_%'] = (summary['ventas'] / summary['int_contacto']) * 100
    summary['conv_contactado_cerrado_%'] = (summary['ventas'] / summary['leads_cerrados']) * 100
    summary['tiempo_total_llamadas_hms'] = summary['tmo_total_seg'].apply(seconds_to_hms)
    
    summary['tmo_total_hms'] = summary['tmo_total_seg'].apply(seconds_to_hms)
    summary['tmo_total_mediana_hms'] = summary['tmo_total_mediana_seg'].apply(seconds_to_hms)
    summary['tmo_ventas_hms'] = summary['tmo_ventas_seg'].apply(seconds_to_hms)
    summary['tmo_ventas_mediana_hms'] = summary['tmo_ventas_mediana_seg'].apply(seconds_to_hms)
    summary['tmo_no_venta_hms'] = summary['tmo_no_venta_seg'].apply(seconds_to_hms)
    summary['tmo_no_venta_mediana_hms'] = summary['tmo_no_venta_mediana_seg'].apply(seconds_to_hms)
    summary['time_acw_hms'] = summary['time_acw_seg'].apply(seconds_to_hms)
    summary['acw_mediana_hms'] = summary['acw_mediana_seg'].apply(seconds_to_hms)
    summary['sla_hms_medio'] = summary['sla_medio_seg'].apply(seconds_to_hms)
    summary['sla_mediana_hms'] = summary['sla_mediana_seg'].apply(seconds_to_hms)
    summary['sla_operativo_mediana_hms'] = summary['sla_operativo_mediana_seg'].apply(seconds_to_hms)
    summary['sla_extra_mediana_hms'] = summary['sla_extra_mediana_seg'].apply(seconds_to_hms)
    summary['sla_fds_mediana_hms'] = summary['sla_fds_mediana_seg'].apply(seconds_to_hms)
    
    summary['conversion_contactos_%'] = summary['conversion_contactos_%'].apply(format_percentage)
    summary['conv_contactado_cerrado_%'] = summary['conv_contactado_cerrado_%'].apply(format_percentage)
    
    hms_cols = [c for c in summary.columns if 'hms' in c]
    for c in hms_cols: summary[c] = summary[c].fillna("00:00:00")
    
    for c in final_cols_order:
         if c not in summary.columns: summary[c] = 0 if 'hms' not in c else "00:00:00"
         
    return summary[final_cols_order].sort_values('ventas', ascending=False)

def apply_comments(worksheet, columns, sheet_config_name):
    from openpyxl.comments import Comment
    if sheet_config_name not in COL_COMMENTS: return

    comments_dict = COL_COMMENTS[sheet_config_name]
    
    for col_idx, col_name in enumerate(columns, 1):
        cell = worksheet.cell(row=1, column=col_idx)
        comment_text = None
        if col_name in comments_dict:
            comment_text = comments_dict[col_name]
        
        if comment_text:
            cell.comment = Comment(comment_text, "System")

if __name__ == "__main__":
    generate_agent_summary()
