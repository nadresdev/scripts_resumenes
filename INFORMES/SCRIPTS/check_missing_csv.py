import pandas as pd
import os

# Definir las columnas estructuradas para Detalle_Leads_Unicos
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

def check_missing_from_csv():
    # Ruta al archivo CSV de origen original
    csv_path = r"C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\INFORMES\KPI_SMART\ORIGEN\11022026_ORIGEN\PLAYFILM.csv"
    
    if not os.path.exists(csv_path):
        print(f"Error: {csv_path} no encontrado.")
        return

    try:
        try:
            df_csv = pd.read_csv(csv_path)
        except UnicodeDecodeError:
            df_csv = pd.read_csv(csv_path, encoding='latin1')
            
        csv_columns = df_csv.columns.tolist()
        
        missing_in_csv = [col for col in target_columns_detalle if col not in csv_columns]
        
        print(f"Total columnas en Detalle_Leads_Unicos: {len(target_columns_detalle)}")
        print(f"Total columnas en CSV Original: {len(csv_columns)}")
        print(f"Columnas presentes en Detalle pero NO en el CSV original ({len(missing_in_csv)}):")
        for col in missing_in_csv:
            print(col)
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    check_missing_from_csv()
