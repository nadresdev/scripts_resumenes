import pandas as pd
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime

# Definir las columnas objetivo de la hoja Leads_Unicos
target_columns = [
    '_id', 'fullname', 'phone', 'provider', 'fxCreated', 'fxFirstcall', 'sla', 'status', 
    'lastOcmCoding', 'lastOcmAgent', 'fxNextcall', 'calidad', 'timeCallTotal', 'resultDesc1', 
    'timeCall1', 'fecha1', 'tmo1', 'timeAcw1', 'callAgent1', 'resultDesc2', 'timeCall2', 
    'fecha2', 'tmo2', 'timeAcw2', 'callAgent2', 'resultDesc3', 'timeCall3', 'fecha3', 'tmo3', 
    'timeAcw3', 'callAgent3', 'resultDesc4', 'timeCall4', 'fecha4', 'tmo4', 'timeAcw4', 
    'callAgent4', 'resultDesc5', 'timeCall5', 'fecha5', 'tmo5', 'timeAcw5', 'callAgent5', 
    'resultDesc6', 'timeCall6', 'fecha6', 'tmo6', 'timeAcw6', 'callAgent6', 'resultDesc7', 
    'timeCall7', 'fecha7', 'tmo7', 'timeAcw7', 'callAgent7', 'resultDesc8', 'timeCall8', 
    'fecha8', 'tmo8', 'timeAcw8', 'callAgent8', 'resultDesc9', 'timeCall9', 'fecha9', 'tmo9', 
    'timeAcw9', 'callAgent9', 'resultDesc10', 'timeCall10', 'fecha10', 'tmo10', 'timeAcw10', 
    'callAgent10', 'contactado', 'venta', 'email', 'message', 'acceptancePolicy', 'acceptance3Party', 
    'campaignLeadId', 'campaignProduct', 'fxSincro', 'OcmId', 'ocm_motor', 'sincro', 'fxLastCall', 
    'playfilmReportStatus', 'playfilmReportDate', 'playfilmReportMessage', 'playfilmReportNumberSales', 
    'description', 'fxModification', 'policyNumber', 'sla_minutos', 'sla_seg', 'sla_hms', 
    'tmo_total_registro', 'Interacciones_x_lead', 'tmo_total_registro_hms', 'tmo_venta', 
    'tmo_venta_total_registro', 'interacciones_venta', 'tmo_venta_hms', 'tiempo_total_llamadas_hms'
]

def select_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Selecciona el archivo CSV de origen",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    root.destroy()
    return file_path

def convert_file_headless(input_path, output_dir=None):
    if not input_path or not os.path.exists(input_path):
        print(f"Error: Archivo no encontrado {input_path}")
        return False
        
    # Si no se especifica output_dir, usar el mismo directorio del archivo de entrada
    if output_dir is None:
        output_dir = os.path.dirname(input_path)

    try:
        # Detectar encoding y leer
        try:
             df = pd.read_csv(input_path)
        except UnicodeDecodeError:
             try:
                 df = pd.read_csv(input_path, encoding='latin1')
             except:
                 df = pd.read_csv(input_path, encoding='cp1252')
        except Exception as e:
             print(f"Error leyendo CSV: {e}")
             return False

        # FILTRO: Excluir filas donde 'provider' contenga 'EXA'
        if 'provider' in df.columns:
            initial_count = len(df)
            df = df[~df['provider'].astype(str).str.upper().str.contains('EXA', na=False)]
            filtered_count = initial_count - len(df)
            if filtered_count > 0:
                print(f"Excluidas {filtered_count} filas con 'EXA' en provider")

        # Reindexar con columnas objetivo
        # Rellenar con nulos si faltan columnas
        df_final = df.reindex(columns=target_columns)
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            
        # Generar nombre output
        # 1. Intentar sacar provider de la columna 'provider'
        provider_name = "UNKNOWN"
        if 'provider' in df.columns:
            providers = df['provider'].dropna().unique()
            if len(providers) > 0:
                raw_prov = str(providers[0]).upper().strip()
                # Estandarizar a valores conocidos
                if 'CAPTA' in raw_prov:
                    provider_name = 'CAPTA'
                elif 'PLAYFILM' in raw_prov:
                    provider_name = 'PLAYFILM'
                elif 'STARTEND' in raw_prov:
                    provider_name = 'STARTEND'
                else:
                    # Limpiar caracteres invalidos si no coincide con los estandar
                    provider_name = "".join([c for c in raw_prov if c.isalnum() or c in (' ', '_', '-')]).strip()
        
        # Si falla, fallback al nombre del archivo original
        if provider_name == "UNKNOWN":
             provider_name = os.path.basename(input_path).split('.')[0]
             # Intentar aplicar logica de keywords al nombre de archivo tambien
             if 'CAPTA' in provider_name.upper(): provider_name = 'CAPTA'
             elif 'PLAYFILM' in provider_name.upper(): provider_name = 'PLAYFILM'
             elif 'STARTEND' in provider_name.upper(): provider_name = 'STARTEND'
             
        timestamp = datetime.now().strftime("%d%m%Y_%H%M%S")
        
        output_filename = f"{provider_name}_LEADS_UNICOS_{timestamp}.xlsx"
        output_path = os.path.join(output_dir, output_filename)

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            original_len = len(df_final)
            df_final.to_excel(writer, sheet_name='Leads_Unicos', index=False)
            
        print(f"Conversión exitosa: {output_path} (Provider detectado: {provider_name})")
        return output_path

    except Exception as e:
        print(f"Error en conversión: {e}")
        return False

def convert_csv_to_excel():
    # Modo Interactivo: Seleccionar archivo y procesar sin popups
    csv_path = select_file()
    if not csv_path: 
        print("No se seleccionó ningún archivo.")
        return
    
    # Al no pasar output_dir, se usará el mismo directorio del csv_path
    out = convert_file_headless(csv_path)
    
    if out:
        print(f"Proceso Completado. Archivo generado: {out}")
    else:
        print("Error: Falló la conversión")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        # Modo CLI
        input_file = sys.argv[1]
        # Si se pasa un segundo argumento, se usa como output_dir. 
        # Si no, se pasa None para que convert_file_headless use el directorio del input.
        out_dir = sys.argv[2] if len(sys.argv) > 2 else None
        
        # Nota: El comportamiento anterior por defecto era una ruta fija. 
        # Si se requiere mantener esa ruta fija en CLI sin argumentos, descomentar la linea siguiente:
        # if out_dir is None: out_dir = r"C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\INFORMES\KPI_SMART\LEADS_UNICOS"
        
        convert_file_headless(input_file, out_dir)
    else:
        convert_csv_to_excel()
