import pandas as pd
import os
import glob
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.comments import Comment

# --- CONFIGURACION DE COLORES ---
# Azul Claro para Encabezados
COLOR_HEADER = "BDD7EE" 
# Coral para Totales (#FF8169)
COLOR_TOTAL = "FF8169"

# --- DICCIONARIOS DE COMENTARIOS ---
# Copiados de scripts anteriores (_5Resumen_Ejecutivo y _2Resumen_Diario)

COMMENTS_EJECUTIVO = {
    "Total leads recibidos": "todos los >=2026",
    "Leads nicos analizados": "todos los >=2026",
    "Contactados (unicos)": "todos los contactado si",
    "No contactados (Unicos)": "todos los contactado no",
    "Contactabilidad % (unicos)": "todos los contactado si/todos los >=2026 * 100",
    "Leads Cerrados": "Conteo de leads con status CERRADO",
    "Ventas (leads insertados mes)": "leads venta Si",
    "interacciones_contacto": "contar(todos los tmo>0)",
    "interacciones_sin_contacto": "contar(todos los tmo=0 o nulo)",
    "Conversin % (contactados)": "leads venta Si/todos los contactado si*100",
    "conversion sobre interacciones con contacto_%": "interacciones de venta/ interacciones tmo > 0",
    "conv_contactado_cerrado_%": "leads contactados si /leads venta y cerrado",
    "interacciones en venta": "interacciones resuldesc= VENTA /POLIZA",
    "Total interacciones": "conteo de todos los tmo>0",
    "TMO totalizado (hh:mm:ss)": "sumatoria de todos los tmo>0",
    "TMO mediana (General) (hh:mm:ss)": "mediana de  todos los tmo>0",
    "TMO interacciones (venta) (hh:mm:ss)": "sumatoria(tmo>0 con resuldesc = VENTA /POLIZA)",
    "TMO mediana (venta) (hh:mm:ss)": "MEDIANA (tmo>0 con resuldesc = VENTA /POLIZA)",
    "TMO mediana (no venta) (hh:mm:ss)": "sumatoria(tmo>0 con resuldesc != VENTA /POLIZA)", 
    "SLA OPERATIVO (10-20 Lu-Vi)": "Mediana de sla en la franja 10-20 Lu-Vi por fxCreated",
    "SLA EXTRAHORARIO (Lu-Vi)": "Mediana de sla en la franja fuera de 10-20 dias Lu-Vi por fxCreated",
    "SLA FIN DE SEMANA": "Mediana de sla en la franja no  Lu-Vi por fxCreated",
    "ACW Mediana (hh:mm:ss)": "ACW Mediana",
    "Total tiempo llamadas": "sumatoria de todo el tiempo en llamadas (timecall)"
}

COMMENTS_DIARIO_GENERIC = {
    "leads_insertados": "Total leads recibidos",
    "contactados": "Total leads con al menos 1 TMO > 0",
    "ventas": "Total leads con venta=SI",
    "interacciones_total": "Suma de interacciones (contacto + intentos)",
    "tiempo_total_llamadas_hms": "Suma de tiempo de llamada total",
    "contactabilidad_%": "Contactados / Leads * 100",
    "conversion_%": "Ventas / Leads * 100",
    "mediana_tmo_x_periodo_seg": "Mediana TMO General (seg)",
    "mediana_tmo_x_periodo_hms": "Mediana TMO General (TMO>0)",
    "mediana_tmo_venta_x_periodo_seg": "Mediana TMO Ventas (seg)",
    "mediana_tmo_venta_x_periodo_hms": "Mediana TMO Ventas",
    "mediana_sla_seg": "Mediana SLA Global (seg)",
    "mediana_sla_hms": "Mediana SLA Global",
    "sla_operativo_mediana_hms": "Mediana SLA (10-18h L-V)",
    "sla_extra_mediana_hms": "Mediana SLA (Extra Horario)",
    "sla_fds_mediana_hms": "Mediana SLA (FDS)",
    "timeAcw_mediana_dia_hms": "Mediana ACW por lead"
}

# Mapeo Hoja -> Diccionario
SHEET_COMMENTS_MAP = {
    "Resumen_Ejecutivo": COMMENTS_EJECUTIVO,
    "Resumen_Diario": COMMENTS_DIARIO_GENERIC,
    "Frecuencia": COMMENTS_DIARIO_GENERIC,
    # Agentes y otros pueden usar el generico si las columnas coinciden
}

def find_latest_file(directory, pattern='*.xlsx'):
    files = glob.glob(os.path.join(directory, pattern))
    files = [f for f in files if not os.path.basename(f).startswith('~$')]
    if not files: return None
    return max(files, key=os.path.getmtime)

def apply_styles_and_order():
    input_dir = r"C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\INFORMES\KPI_SMART\RESUMEN_EJECUTIVO"
    
    print(f"Buscando archivo para estilos en: {input_dir}")
    latest_file = find_latest_file(input_dir, pattern='*_RESUMEN_EJECUTIVO_*.xlsx')
    
    if not latest_file:
        print("No se encontr archivo Resumen Ejecutivo.")
        return

    print(f"Procesando estilos en: {latest_file}")
    
    try:
        wb = load_workbook(latest_file)
        
        # --- 1. REORDENAR HOJAS ---
        # Orden deseado: [Resumen_Ejecutivo, Resumen_Agentes, ...resto en orden original]
        
        # Identificar nombres reales (case sensitive)
        sheet_names = wb.sheetnames
        
        # Normalizar para busqueda
        exec_name = next((s for s in sheet_names if 'Resumen_Ejecutivo' in s), None)
        agents_name = next((s for s in sheet_names if 'Resumen_Agentes' in s or 'Agentes' in s), None)
        
        # Logica mover
        # Mover Agentes primero al index 0
        if agents_name:
            # En openpyxl, move_sheet mueve relativo. Calcular offset para ir al principio.
            idx = wb.sheetnames.index(agents_name)
            wb.move_sheet(wb[agents_name], offset=-idx)
            
        # Mover Ejecutivo despues al index 0 (empujando Agentes al 1)
        if exec_name:
            idx = wb.sheetnames.index(exec_name)
            wb.move_sheet(wb[exec_name], offset=-idx)
            
        print(f"Orden de hojas ajustado: {wb.sheetnames}")
        
        # --- 2. ESTILOS Y COMENTARIOS ---
        
        fill_header = PatternFill(start_color=COLOR_HEADER, end_color=COLOR_HEADER, fill_type="solid")
        fill_total = PatternFill(start_color=COLOR_TOTAL, end_color=COLOR_TOTAL, fill_type="solid")
        
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            
            # Determinar diccionario de comentarios
            comments_dict = {}
            if sheet_name in SHEET_COMMENTS_MAP:
                comments_dict = SHEET_COMMENTS_MAP[sheet_name]
            else:
                s_lower = sheet_name.lower()
                if 'diario' in s_lower or 'frecuencia' in s_lower or 'agentes' in s_lower:
                    comments_dict = COMMENTS_DIARIO_GENERIC
            
            # Iterar Filas
            max_row = ws.max_row
            max_col = ws.max_column
            
            for row_idx, row in enumerate(ws.iter_rows(), 1):
                # Check si es fila TOTAL (Columna 1 contiene 'TOTAL')
                is_total_row = False
                first_cell_val = str(row[0].value).upper() if row[0].value else ""
                
                if "TOTAL" in first_cell_val:
                    is_total_row = True
                
                for cell in row:
                    # Estilo Fila Total
                    if is_total_row:
                        cell.fill = fill_total
                        
                    # Encabezados (Fila 1)
                    if row_idx == 1:
                        cell.fill = fill_header
                        # Agregar Comentario
                        val = str(cell.value).strip() if cell.value else ""
                        if val in comments_dict:
                            # Evitar duplicar si ya tiene
                            if not cell.comment:
                                cell.comment = Comment(comments_dict[val], "System")
                                
        # Guardar
        wb.save(latest_file)
        print("Estilos aplicados exitosamente.")

    except Exception as e:
        print(f"Error aplicando estilos: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    apply_styles_and_order()
