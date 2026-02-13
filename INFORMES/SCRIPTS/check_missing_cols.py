import pandas as pd
import os
import glob

# Definir las columnas estructuradas para Detalle_Leads_Unicos (Target)
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
    'tmo_venta_hms', 'tmo_venta_total_registro', 'interacciones_venta', 'sla_minutos', 'sla_hms', 'sla_seg', 
    'contactado', 'venta'
]

def find_latest_file(directory, pattern='*.xlsx'):
    files = glob.glob(os.path.join(directory, pattern))
    files = [f for f in files if not os.path.basename(f).startswith('~$')]
    if not files:
        return None
    return max(files, key=os.path.getmtime)

def compare_columns():
    input_dir = r"C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\INFORMES\KPI_SMART\LEADS_UNICOS"
    latest_file = find_latest_file(input_dir)
    
    if not latest_file:
        print("No se encontró archivo de origen.")
        return

    print(f"Analizando archivo: {latest_file}")
    
    try:
        df = pd.read_excel(latest_file, sheet_name='Leads_Unicos')
        source_columns = df.columns.tolist()
        
        missing_columns = [col for col in target_columns_detalle if col not in source_columns]
        
        print(f"\nTotal columnas en Destino (Detalle): {len(target_columns_detalle)}")
        print(f"Total columnas en Origen (Leads_Unicos): {len(source_columns)}")
        print(f"Columnas faltantes ({len(missing_columns)}):")
        for col in missing_columns:
            print(f"- {col}")
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    compare_columns()
