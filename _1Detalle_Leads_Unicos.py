import pandas as pd
import os
import glob
import re
from datetime import datetime
import numpy as np

# Definir las columnas estructuradas para Detalle_Leads_Unicos
# Se elimina sla_minutos segn requerimiento
target_columns_detalle = [
     '_id', 'fullname', 'phone', 'fxCreated', 'fxFirstcall', 'sla', 'status', 'lastOcmCoding', 'lastOcmAgent', 
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
        return "Unknown"
    
    unique_providers = df['provider'].dropna().astype(str).unique()
    
    for prov in unique_providers:
        prov_upper = prov.upper()
        if 'CAPTA' in prov_upper:
            return 'CAPTA'
        elif 'PLAYFILM' in prov_upper:
            return 'PLAYFILM'
        elif 'STARTEND' in prov_upper:
            return 'STARTEND'
            
    return unique_providers[0] if len(unique_providers) > 0 else "Unknown"

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
        
        for index, row in df.iterrows():
            current_tmo_venta_sum = 0 # Para tmo_venta (acumulado hasta la venta)
            current_tmo_venta_total = 0 # Para tmo_venta_total_registro (solo tmo de ventas)
            current_interacciones_venta = 0 # Para interacciones_venta
            
            # Recorrer intentos del 10 al 1 (descendente, como se solicita para tmo_venta?)
            # La instruccin dice: "se suman todos los tmo>0 donde su resultdesc no haya sido venta desde tmo10 descendiendo hasta el tmo correspondiente a la venta."
            # Esto sugiere que buscamos la PRIMERA venta desde el final? o la ltima en orden temporal (que sera tmo10 si fuese cronolgico?)
            # Asumiremos orden cronolgico estndar 1->10 (1 es primero).
            # Si dice "descendiendo desde tmo10", quizs se refiere a acumular HACIA ATRS desde la venta encontrada?
            # Releamos: "suman todos los tmo>0 donde su resultdesc no haya sido venta desde tmo10 descendiendo hasta el tmo correspondiente a la venta."
            # Interpretacin: Encontrar el ndice de la venta (ej. intento 5). Sumar TMOs desde el 10 bajando hacia el 5? o del 5 subiendo?
            # Usualmente se quiere saber cunto tiempo se invirti HASTA conseguir la venta. Eso sera suma de TMO 1 a TMO N (donde N es venta).
            # Pero la instruccin es especfica. "desde tmo10 descendiendo...".
            # CASO A: Venta en intento 3. Sumar TMOs de 3, 4, 5... 10? (Tiempo post-venta?) -> Raro.
            # CASO B: Es un acumulado inverso?
            # CASO C: Simplemente sumar TODO lo que NO SEA venta, HASTA llegar a la venta?
            
            # Vamos a simplificar con una lgica comn de "Costo de Venta":
            # 2. tmo_venta: Suma de TMOs de interacciones que NO son venta, ANTES de la venta? 
            # Re-leyendo tu imagen adjunta (si pudiera verla, pero me baso en texto):
            # texto anterior: "en las columnas resultDesc[] VENTA / POLIZA se suman todos los tmo>0 donde su resultdesc no haya sido venta desde tmo10 descendiendo hasta el tmo correspondiente a la venta."
            # Esto suena a que vamos del 10 al 1. Si encuentro venta en el 8, sumo tmo de 10 y 9 (si no son venta)?
            # Voy a implementar la iteracin 10 -> 1.
            
            found_sale_index = -1
            temp_tmo_acum = 0
            
            # Buscando la venta (asumiendo que puede haber mltiples, tomamos... la primera que encontremos desde el final? o la ltima cronolgica?)
            # Si vamos de 10 a 1, la primera que encontremos ser la ltima cronolgicamente.
            
            # Estructura auxiliar para iterar 10 a 1
            indices_desc = range(10, 0, -1)
            
            # Primero identifiquemos dnde ocurri la venta (si la hubo en los parciales)
            # A veces lastOcmCoding dice venta pero no los parciales, o viceversa.
            # Usaremos resultDescN para identificar el intento de venta.
            
            # Clculo de: tmo_venta_total_registro y interacciones_venta
            # "el total de los tmo>0 donde resultdesc= VENTA / POLIZA" (Esto es fcil, suma directa condicional)
            # "cada tmo>0 donde resultdesc= VENTA / POLIZA cuenta como una interaccin venta"
            
            t_venta_total = 0
            i_venta = 0
            
            for i in range(1, 11):
                rd = str(row.get(f'resultDesc{i}', '')).upper().strip()
                t = row.get(f'tmo{i}', 0)
                if t > 0 and ('VENTA' in rd or 'POLIZA' in rd):
                    t_venta_total += t
                    i_venta += 1
            
            tmo_venta_total_list.append(t_venta_total)
            interacciones_venta_list.append(i_venta)
            
            # Clculo de: tmo_venta (la lgica compleja)
            # "suman todos los tmo>0 donde su resultdesc no haya sido venta desde tmo10 descendiendo hasta el tmo correspondiente a la venta."
            # Si Venta est en intento 5. Sumamos 10, 9, 8, 7, 6...? 
            # Si no hay venta, es 0? Asumiremos s.
            
            val_tmo_venta = 0
            # Solo calcular si hubo alguna venta detectada en los parciales o global?
            # La instruccin implica buscar "el tmo correspondiente a la venta". As que debe haber una venta en resultDesc.
            
            # Buscar el ndice de la venta ms alto (ltimo intento que fue venta)
            sale_idx = -1
            for i in range(10, 0, -1):
                rd = str(row.get(f'resultDesc{i}', '')).upper().strip()
                if 'VENTA' in rd or 'POLIZA' in rd:
                    sale_idx = i
                    break # Encontramos la ltima venta (desde 10 hacia abajo es la primera que vemos)
            
            if sale_idx != -1:
                # Si hubo venta en sale_idx (ej. 5).
                # Sumar tmo>0 donde NO sea venta, desde 10 hasta 5? O desde 10 hasta 6? "hasta el tmo correspondiente a la venta".
                # Interpretacin literal: Suma Loop i=10 to i=sale_idx. IF resultDesc[i] NO ES VENTA -> Suma.
                
                for k in range(10, sale_idx - 1, -1): # Range incluye sale_idx
                    rd_k = str(row.get(f'resultDesc{k}', '')).upper().strip()
                    t_k = row.get(f'tmo{k}', 0)
                    
                    if t_k > 0:
                         # Si resultdesc NO es venta, suma?
                         # "donde su resultdesc no haya sido venta"
                         # Pero la instruccin dice "hasta el tmo correspondiente a la venta".
                         # Si k == sale_idx, es venta. Se suma?
                         # La lgica "Costo de Adquisicin" suele incluir todo el tiempo gastado.
                         # Pero la frase "donde ... no haya sido venta" excluye las ventas de la suma?
                         
                         if not ('VENTA' in rd_k or 'POLIZA' in rd_k):
                             val_tmo_venta += t_k
                             
                # Nota: Si en el intento de venta (sale_idx) tambin hay TMO, pero la condicin dice "donde no haya sido venta", entonces el tiempo propio de la venta NO se suma a esta mtrica `tmo_venta`.
                # Solo se suman los tiempos "desperdiciados" o de no-cierre posteriores (o anteriores si el orden fuera distinto)?
                # Al ir de 10 a Venta, estamos sumando los intentos POSTERIORES a la venta (si 10 es lo ltimo) + los intentos fallidos intercalados?
                # Si la venta fue el ltimo paso (ej 10), el loop va de 10 a 10. Si 10 es venta, no suma. Resultado 0.
                # Si venta fue en 1. Loop 10 a 1. Suma 10,9,8...2 (si no ventas). Venta en 1 no suma.
                pass
            
            tmo_venta_list.append(val_tmo_venta)

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

        base_name = os.path.basename(latest_file)
        match = re.search(r'LEADS_UNICOS_(\d{8}_\d{6})', base_name)
        if match:
            timestamp = match.group(1)
        else:
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
