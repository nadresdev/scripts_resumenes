import os
import shutil
import glob
import time
import subprocess
import sys

from datetime import datetime

# Rutas - Ajustar segun entorno
BASE_DIR = r"C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\INFORMES\KPI_SMART"

# Definir la carpeta de origen basada en la fecha actual DDMMYYYY
current_date_str = datetime.now().strftime("%d%m%Y")
# Buscar carpetas que coincidan con el patron DDMMYYYY_ORIGEN
search_pattern = os.path.join(BASE_DIR, "ORIGEN", f"*_ORIGEN")
candidates = glob.glob(search_pattern)

# Filtrar para encontrar la mas reciente o la de hoy
# Si existe la exacta de hoy, usamos esa.
target_origin = os.path.join(BASE_DIR, "ORIGEN", f"{current_date_str}_ORIGEN")

if os.path.exists(target_origin):
    ORIGIN_DIR = target_origin
else:
    # Si no existe la de hoy, buscamos la mas reciente por fecha en el nombre folder
    # Logica simple: Buscar la mas reciente por fecha de modificacion si no existe hoy
    if candidates:
        ORIGIN_DIR = max(candidates, key=os.path.getmtime)
        print(f"No existe carpeta de hoy ({target_origin}). Usando la más reciente encontrada: {ORIGIN_DIR}")
    else:
        # Fallback create
        ORIGIN_DIR = target_origin

INPUT_LEADS_DIR = os.path.join(BASE_DIR, "LEADS_UNICOS")
# Carpeta PROCESO dentro de la misma carpeta de origen para mantener orden por dia
PROCESSED_DIR = os.path.join(ORIGIN_DIR, "PROCESO")

# Lista de Scripts a Ejecutar en Orden
SCRIPTS_DIR = r"C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\INFORMES\SCRIPTS"
SCRIPTS_SEQUENCE = [
    "_1Detalle_Leads_Unicos.py",
    "_2Resumen_Diario.py",
    "_3Resumen_Agentes.py",
    "_4Resumen_Semanal.py",
    "_5Resumen_Ejecutivo.py",
    "_6Frecuencia_Horaria.py",
    "_7Estilos_Finales.py"
]

def ensure_dir(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)

def run_script(script_name):
    script_path = os.path.join(SCRIPTS_DIR, script_name)
    print(f"--- Ejecutando {script_name} ---")
    try:
        # Usar el mismo python executable que corre este script
        # NO capturar output para que se vea en consola principal
        subprocess.run([sys.executable, script_path], check=True)
        print(f"OK: {script_name}")
    except subprocess.CalledProcessError as e:
        print(f"ERROR: Fallo al ejecutar {script_name}")
        return False
    return True

def clean_intermediate_files():
    # Opcional: Limpiar carpetas intermedias si se desea pureza total
    # Pero los scripts buscan "latest", asi que con timestamp nuevo basta.
    pass


def run_converter(input_file, output_dir):
    script_name = "_0conversor.py"
    script_path = os.path.join(SCRIPTS_DIR, script_name)
    print(f"--- Ejecutando Conversor para: {input_file} ---")
    try:
        # Llama a _0conversor.py con argumentos: input_file output_dir
        # No capturar output para ver logs del conversor
        subprocess.run([sys.executable, script_path, input_file, output_dir], check=True)
        print(f"OK: Conversión exitosa.")
    except subprocess.CalledProcessError as e:
        print(f"ERROR: Fallo al convertir {input_file}")
        print(e.stderr)
        return False
    return True

def main():
    print("INICIANDO PROCESAMIENTO POR LOTES AUTOMATICO")
    print(f"Buscando archivos en: {ORIGIN_DIR}")
    
    ensure_dir(PROCESSED_DIR)
    ensure_dir(INPUT_LEADS_DIR)
    
    # 1. Listar CSVs en ORIGEN
    csv_files = glob.glob(os.path.join(ORIGIN_DIR, "*.csv"))
    # Tambien xlsx si aplica
    xlsx_files = glob.glob(os.path.join(ORIGIN_DIR, "*.xlsx"))
    all_files = csv_files + xlsx_files
    
    if not all_files:
        print(f"No se encontraron archivos en {ORIGIN_DIR}")
        # Intentar crear la carpeta si no existe para evitar confundir al usuario
        if not os.path.exists(ORIGIN_DIR):
             print(f"La carpeta '{ORIGIN_DIR}' no existe. Creándola...")
             os.makedirs(ORIGIN_DIR)
             print("Por favor, coloca los archivos CSV ahí y vuelve a ejecutar.")
        return

    print(f"Encontrados {len(all_files)} archivos para procesar.")

    for file_path in all_files:
        filename = os.path.basename(file_path)
        print(f"\n==================================================")
        print(f"PROCESANDO: {filename}")
        print(f"==================================================")
        
        # 2. Ejecutar Conversor: ORIGEN -> LEADS_UNICOS
        if not run_converter(file_path, INPUT_LEADS_DIR):
            print(f"ABORTANDO CICLO PARA: {filename} (Fallo conversión)")
            continue
            
        # 3. Ejecutar Secuencia Standard (_1 en adelante)
        # _1Detalle... buscará el 'latest' en LEADS_UNICOS, que acaba de ser creado por el conversor.
        success_sequence = True
        for script in SCRIPTS_SEQUENCE:
            if not run_script(script):
                success_sequence = False
                break
        
        if success_sequence:
            print(f"CICLO COMPLETADO EXITOSAMENTE PARA: {filename}")
            # 4. Mover original a PROCESSED (Opcional)
            try:
                dest_path = os.path.join(PROCESSED_DIR, filename)
                if os.path.exists(dest_path):
                    try:
                        os.remove(dest_path)
                        print(f"Sobrescribiendo archivo existente en PROCESO: {filename}")
                    except Exception as e:
                        print(f"Warning: No se pudo eliminar archivo destino: {e}")
                
                shutil.move(file_path, dest_path)
                print(f"Archivo movido a PROCESO: {filename}")
            except Exception as e:
                print(f"Warning: No se pudo mover a PROCESO: {e}")
        else:
            print(f"CICLO FALLIDO PARA: {filename}. Se detiene este archivo.")
            
        # Pausa para asegurar timestamps distintos
        time.sleep(2)

    print("\nPROCESAMIENTO POR LOTES FINALIZADO.")

if __name__ == "__main__":
    main()
