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
    return file_path

def convert_file_headless(input_path, output_dir=r"C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\INFORMES\KPI_SMART\LEADS_UNICOS"):
    if not input_path or not os.path.exists(input_path):
        print(f"Error: Archivo no encontrado {input_path}")
        return False

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

        # Reindexar con columnas objetivo
        # Rellenar con nulos si faltan columnas
        df_final = df.reindex(columns=target_columns)
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            
        # Generar nombre output
        provider_match = os.path.basename(input_path).split('.')[0]
        # Deberiamos intentar preservar el nombre o usar timestamp?
        # El user quiere batch. Si usamos timestamp, cada uno tendra uno distinto.
        # Pero _1 busca "latest".
        timestamp = datetime.now().strftime("%d%m%Y_%H%M%S")
        
        # Intentar mantener algo del nombre original para trazabilidad
        # Ojo: El conversor original generaba LEADS_UNICOS_{TIMESTAMP}.xlsx
        # Si corremos en batch, necesitaremos que el nombre sea "nuevo" para cada archivo.
        # Pero si _1 solo busca el "latest", y corremos secuencialmente, est bien.
        # Podriamos poner el nombre del archivo original en el output.
        output_filename = f"LEADS_UNICOS_{provider_match}_{timestamp}.xlsx"
        output_path = os.path.join(output_dir, output_filename)

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='Leads_Unicos', index=False)
            
        print(f"Conversión exitosa: {output_path}")
        return output_path

    except Exception as e:
        print(f"Error en conversión: {e}")
        return False

def convert_csv_to_excel():
    # Modo Interactivo Original
    csv_path = select_file()
    if not csv_path: return
    
    out = convert_file_headless(csv_path)
    if out:
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Proceso Completado", f"Archivo generado:\n{out}")
        root.destroy()
    else:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Error", "Fallo la conversión")
        root.destroy()

if __name__ == "__main__":
    if len(sys.argv) > 1:
        # Modo CLI
        input_file = sys.argv[1]
        out_dir = sys.argv[2] if len(sys.argv) > 2 else r"C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\INFORMES\KPI_SMART\LEADS_UNICOS"
        convert_file_headless(input_file, out_dir)
    else:
        convert_csv_to_excel()
