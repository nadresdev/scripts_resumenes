import pandas as pd
import os
import glob
import re
from datetime import datetime
import numpy as np

# Definir las columnas estructuradas para Detalle_Leads_Unicos
# Se elimina sla_minutos segn requerimiento
# Definir las columnas estructuradas para Detalle_Leads_Unicos
# Se elimina sla_minutos segn requerimiento
target_columns_detalle = [
     '_id', 'fullname', 'phone', 'provider', 'fxCreated', 'fxFirstcall', 'sla', 'status', 'lastOcmCoding', 'lastOcmAgent', 
    'fxNextcall', 'calidad', 'timeCallTotal', 'resultDesc1', 'timeCall1', 'fecha1', 'tmo1', 'timeAcw1', 'callAgent1', 
    'resultDesc2', 'timeCall2', 'fecha2', 'tmo2', 'timeAcw2', 'callAgent2', 'resultDesc3', 'timeCall3', 'fecha3', 
    'tmo3', 'timeAcw3', 'callAgent3', 'resultDesc4', 'timeCall4', 'fecha4', 'tmo4', 'timeAcw4', 'callAgent4', 
    'resultDesc5', 'timeCall5', 'fecha5', 'tmo5', 'timeAcw5', 'callAgent5', 'resultDesc6', 'timeCall6', 'fecha6', 
    'tmo6', 'timeAcw6', 'callAgent6', 'resultDesc7', 'timeCall7', 'fecha7', 'tmo7', 'timeAcw7', 'callAgent7', 
    'resultDesc8', 'timeCall8', 'fecha8', 'tmo8', 'timeAcw8', 'callAgent8', 'resultDesc9', 'timeCall9', 'fecha9', 
    'tmo9', 'timeAcw9', 'callAgent9', 'resultDesc10', 'timeCall10', 'fecha10', 'tmo10', 'timeAcw10', 'callAgent10', 
    'tiempo_total_llamadas_hms', 'tmo_total_registro', 'tmo_total_registro_hms', 'Interacciones_x_lead', 'tmo_venta', 
    'tmo_venta_hms', 'tmo_venta_total_registro', 'interacciones_venta', 'sla_hms', 'sla_seg', 
    'contactado', 'venta'
]


def find_latest_file(directory, pattern='*.xlsx'):
    files = glob.glob(os.path.join(directory, pattern))
    files = [f for f in files if not os.path.basename(f).startswith('~$')]
    
    if not files:
        return None
    latest_file = max(files, key=os.path.getmtime)
    return latest_file

def extract_provider(df):
    if 'provider' not in df.columns:
        return "UNKNOWN"
    
    providers = df['provider'].dropna().unique()
    
    # Tomar el primero no nulo/vacio
    if len(providers) == 0: return "UNKNOWN"
    
    raw_prov = str(providers[0]).strip().upper()
    
    if 'CAPTA' in raw_prov: return 'CAPTA'
    elif 'PLAYFILM' in raw_prov: return 'PLAYFILM'
    elif 'STARTEND' in raw_prov: return 'STARTEND'
    
    # Limpiar caracteres raros si no coincide
    clean_prov = "".join([c for c in raw_prov if c.isalnum() or c in (' ', '_', '-')]).strip()
    return clean_prov if clean_prov else "UNKNOWN"

def seconds_to_hms(seconds):
    try:
        if pd.isna(seconds): return "00:00:00"
        seconds = int(float(seconds))
        if seconds < 0: return "00:00:00"
        m, s = divmod(seconds, 60)
        h, m = divmod(m, 60)
        return "{:02d}:{:02d}:{:02d}".format(h, m, s)
    except (ValueError, TypeError):
        return "00:00:00"

def calculate_sla_hms(seconds):
     # Similar a seconds_to_hms pero para manejo individual si es necesario
     return seconds_to_hms(seconds)

def process_leads_detalle():
    input_dir = r"C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\INFORMES\KPI_SMART\LEADS_UNICOS"
    output_dir = r"C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\INFORMES\KPI_SMART\DETALLE_LEADS_UNICOS"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    print(f"Buscando archivo ms reciente en: {input_dir}")
    latest_file = find_latest_file(input_dir)
    
    if not latest_file:
        print("No se encontraron archivos XLSX vlidos.")
        return

    print(f"Procesando archivo: {latest_file}")
    
    try:
        original_sheets = pd.read_excel(latest_file, sheet_name=None)
        if 'Leads_Unicos' not in original_sheets:
            print("No se encontr la hoja 'Leads_Unicos'")
            return
            
        df = original_sheets['Leads_Unicos']
        provider = extract_provider(df)
        print(f"Proveedor detectado: {provider}")
        
        # --- Preprocesamiento de Columnas ---
        # Asegurar columnas numricas
        tmo_cols = [f'tmo{i}' for i in range(1, 11)]
        result_cols = [f'resultDesc{i}' for i in range(1, 11)] # Asumiendo que resultDesc estn presentes
        time_call_cols = [f'timeCall{i}' for i in range(1, 11)]

        for col in tmo_cols + time_call_cols:
            if col not in df.columns:
                df[col] = 0
            else:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # Asegurar columnas de fecha
        date_cols = ['fxCreated', 'fxFirstcall']
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
                # Eliminar timezone
                try:
                    if pd.api.types.is_datetime64_any_dtype(df[col]):
                         df[col] = df[col].dt.tz_localize(None)
                except Exception:
                    pass

        # === CLCULOS ===
        
        # 1. tiempo_total_llamadas_hms (Suma timeCall1-10 a HMS)
        df['total_time_seconds'] = df[time_call_cols].sum(axis=1)
        df['tiempo_total_llamadas_hms'] = df['total_time_seconds'].apply(seconds_to_hms)

        # 2. tmo_total_registro (Sumatoria de tmo > 0)
        # Como llenamos NaNs con 0, la suma directa funciona.
        df['tmo_total_registro'] = df[tmo_cols].sum(axis=1)
        
        # 3. tmo_total_registro_hms
        df['tmo_total_registro_hms'] = df['tmo_total_registro'].apply(seconds_to_hms)
        
        # 4. Interacciones_x_lead (count tmo > 0)
        # Cuenta cuantas columnas de tmo tienen valor > 0 para cada fila
        df['Interacciones_x_lead'] = (df[tmo_cols] > 0).sum(axis=1)
        
        # 5. sla_seg (fxFirstcall - fxCreated en segundos)
        # Asegurar que ambos son datetime antes de restar
        df['sla_seg'] = (df['fxFirstcall'] - df['fxCreated']).dt.total_seconds()
        
        # 6. sla_hms
        df['sla_hms'] = df['sla_seg'].apply(seconds_to_hms)
        
        # 9. contactado (Al menos un tmo > 0) - True/False
        # Ya calculamos la suma total de tmos, si es > 0, es contactado
        df['contactado'] = df['tmo_total_registro'] > 0
        
        # 10. venta (lastOcmCoding = VENTA / POLIZA) - True/False
        # Normalizar lastOcmCoding a maysculas para comparacin y quitar espacios
        if 'lastOcmCoding' in df.columns:
            df['venta'] = df['lastOcmCoding'].astype(str).str.upper().str.strip().isin(['VENTA', 'POLIZA', 'VENTA / POLIZA'])
        else:
             df['venta'] = False

        # --- Clculos Complejos de Venta (Iteracin por filas) ---
        # tmo_venta
        # tmo_venta_total_registro
        # interacciones_venta
        
        tmo_venta_list = []
        tmo_venta_total_list = []
        interacciones_venta_list = []
        
        # Optimización: Extraer arrays numpy para velocidad
        res_cols = [f'resultDesc{i}' for i in range(1, 11)]
        tmo_cols_idx = [f'tmo{i}' for i in range(1, 11)]
        
        # Llenar nulos y asegurar tipos antes de extraer
        if all(c in df.columns for c in res_cols) and all(c in df.columns for c in tmo_cols_idx):
             # Convertir a clean numpy arrays
             # resultDesc -> Upper string
             # Importante: Convertir a dtype='U' para que np.char.find funcione correctamente
             results_arr = df[res_cols].fillna('').astype(str).apply(lambda x: x.str.upper().str.strip()).to_numpy(dtype='U')
             tmos_arr = df[tmo_cols_idx].fillna(0).values
             
             # Iterar arrays (mucho mas rapido que iterrows)
             for i in range(len(df)):
                 row_res = results_arr[i] # Array de 10 elementos
                 row_tmo = tmos_arr[i]    # Array de 10 elementos -> indices 0..9 corresponden a 1..10
                 
                 # 1. tmo_venta_total_registro y interacciones_venta
                 # Sumar tmo>0 donde res sea venta
                 # Indices 0..9 -> tmo1..10
                 # Detectar indices de ventas
                 is_venta = np.char.find(row_res, 'VENTA') >= 0
                 is_poliza = np.char.find(row_res, 'POLIZA') >= 0
                 sale_mask = is_venta | is_poliza
                 
                 # Filtrar tmo > 0
                 tmo_pos = row_tmo > 0
                 
                 # Interseccion: TMO>0 y ES VENTA
                 mask_venta_valid = sale_mask & tmo_pos
                 
                 t_venta_total = np.sum(row_tmo[mask_venta_valid])
                 i_venta = np.sum(mask_venta_valid)
                 
                 tmo_venta_total_list.append(t_venta_total)
                 interacciones_venta_list.append(i_venta)
                 
                 # 2. tmo_venta (La logica compleja: desde 10 hacia venta)
                 # En numpy, indices son 0..9 (tmo1..tmo10)
                 # Queremos ir de 9 hacia 0.
                 
                 val_tmo_venta = 0
                 
                 # Buscar ultimo indice de venta (mas a la derecha / mas alto tmoN)
                 # np.where devuelve indices donde mask es true
                 sale_indices = np.where(sale_mask)[0]
                 
                 if len(sale_indices) > 0:
                      # El ultimo indice de venta (mas a la derecha / mas alto tmoN)
                      last_sale_idx = sale_indices[-1] # Ej. 4 (tmo5)
                      
                      # Sumar TMO de items > 0 y NO VENTA, desde fin (9) hasta last_sale_idx (exclusive? o inclusive segun logica anterior?)
                      # Logica anterior: range(10, sale_idx - 1, -1) -> incluia sale_idx, pero filtraba "if not venta".
                      # Si last_sale_idx es venta, no se suma.
                      # Asi que sumamos en rango [last_sale_idx, 9] todo lo que NO sea venta y TMO>0
                      
                      # Segmento de interes: row_tmo[last_sale_idx : 10]
                      # Segmento de mask ventas: sale_mask[last_sale_idx : 10]
                      
                      # Queremos sumar donde NO es venta y TMO>0
                      # slice notation: array[start:end]
                      
                      relevant_tmos = row_tmo[last_sale_idx:]
                      relevant_sales = sale_mask[last_sale_idx:]
                      
                      # Condicion: TMO > 0 y NOT VENTA
                      # Mask local
                      valid_tmo_mask = relevant_tmos > 0
                      not_sale_mask = ~relevant_sales
                      
                      final_mask = valid_tmo_mask & not_sale_mask
                      
                      val_tmo_venta = np.sum(relevant_tmos[final_mask])
                 
                 tmo_venta_list.append(val_tmo_venta)
                 
        else:
             # Fallback si faltan columnas (todo a 0)
             print("Warning: Faltan columnas tmo/resultDesc para calculo optimizado. Usando ceros.")
             tmo_venta_list = [0] * len(df)
             tmo_venta_total_list = [0] * len(df)
             interacciones_venta_list = [0] * len(df)
                 
        df['tmo_venta'] = tmo_venta_list
        df['tmo_venta_hms'] = df['tmo_venta'].apply(seconds_to_hms)
        df['tmo_venta_total_registro'] = tmo_venta_total_list
        df['interacciones_venta'] = interacciones_venta_list
        

        # --- Limpieza Final de Nulos ---
        print("Aplicando limpieza de nulos (Numricos -> 0, Tiempo -> 00:00:00)...")
        
        # Columnas HMS (formato HH:MM:SS)
        hms_columns = [
            'tiempo_total_llamadas_hms', 'tmo_total_registro_hms', 'tmo_venta_hms', 'sla_hms'
        ]
        
        for col in hms_columns:
            if col in df.columns:
                df[col] = df[col].fillna("00:00:00")
                # Asegurar que strings vacos tambin sean 00:00:00
                df[col] = df[col].replace('', "00:00:00")

        # Columnas Numricas (incluyendo las generadas y las originales del target)
        # Identificamos columnas numricas en el target (excluyendo _id, fechas, textos)
        # Lista aproximada de columnas que deberan ser numricas en el output
        numeric_target_cols = [
            'phone', 'sla', 'timeCallTotal', 
            'tmo_total_registro', 'Interacciones_x_lead', 'tmo_venta', 
            'tmo_venta_total_registro', 'interacciones_venta', 'sla_seg'
        ]
        
        # Agregar tmo1..10, timeCall1..10, timeAcw1..10
        for i in range(1, 11):
            numeric_target_cols.extend([f'timeCall{i}', f'tmo{i}', f'timeAcw{i}'])

        for col in numeric_target_cols:
             if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # Reordenar columnas finales
        df_detalle = df.reindex(columns=target_columns_detalle)
        
        # Aplicar limpieza de nuevo tras reindex para asegurar que columnas nuevas (si las hay) tengan 0 y no NaN
        # (Reindex introduce NaN si la columna no exista en df)
        
        # Repetir limpieza para las columnas en df_detalle
        for col in hms_columns:
             if col in df_detalle.columns:
                 df_detalle[col] = df_detalle[col].fillna("00:00:00")
        
        for col in numeric_target_cols:
             if col in df_detalle.columns:
                 df_detalle[col] = df_detalle[col].fillna(0)

        # Generar Nombre Salida
        base_name = os.path.basename(latest_file)
        # Buscar patron YYYYMMDD_HHMMSS al final del nombre (mas seguro)
        match = re.search(r'(\d{8}_\d{6})', base_name)
        # Siempre generar nuevo timestamp para evitar sobreescrituras si se corre rapido
        timestamp = datetime.now().strftime("%d%m%Y_%H%M%S")
        
        output_filename = f"{provider}_DETALLE_LEADS_UNICOS_{timestamp}.xlsx"
        output_path = os.path.join(output_dir, output_filename)
        
        print(f"Guardando en: {output_path}")
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Guardar hojas originales
            for sheet_name, sheet_df in original_sheets.items():
                if isinstance(sheet_df, pd.DataFrame):
                    sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Guardar hoja detalle (si ya exista, se sobrescribe con la nueva versin, pero aqu es un nuevo archivo)
            # Si 'Detalle_Leads_Unicos' estaba en original, aqu lo actualizamos.
            df_detalle.to_excel(writer, sheet_name='Detalle_Leads_Unicos', index=False)
            
        print("Proceso completado exitosamente.")
        
    except Exception as e:
        print(f"Error durante el procesamiento: {e}")
        # Imprimir traza completa para debug si es necesario
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    process_leads_detalle()
