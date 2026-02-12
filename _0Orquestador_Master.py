import os
import shutil
import glob
import time
import subprocess
import sys

# Rutas - Ajustar segun entorno
BASE_DIR = r"C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\INFORMES\KPI_SMART"
ORIGIN_DIR = os.path.join(BASE_DIR, r"ORIGEN\11022026_ORIGEN")
INPUT_LEADS_DIR = os.path.join(BASE_DIR, "LEADS_UNICOS")
PROCESSED_DIR = os.path.join(BASE_DIR, "PROCESSED_BATCH")

# Lista de Scripts a Ejecutar en Orden
SCRIPTS_DIR = r"C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\INFORMES\SCRIPTS"
SCRIPTS_SEQUENCE = [
    "_0conversor.py",
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
        result = subprocess.run([sys.executable, script_path], check=True, capture_output=True, text=True)
        print(f"OK: {script_name}")
        # print(result.stdout) # Descomentar para ver output detallado
    except subprocess.CalledProcessError as e:
        print(f"ERROR: Fallo al ejecutar {script_name}")
        print(e.stderr)
        return False
    return True

def clean_intermediate_files():
    # Opcional: Limpiar carpetas intermedias si se desea pureza total
    # Pero los scripts buscan "latest", asi que con timestamp nuevo basta.
    pass

def main():
    print("INICIANDO PROCESAMIENTO POR LOTES")
    
    ensure_dir(PROCESSED_DIR)
    ensure_dir(INPUT_LEADS_DIR)
    
    # 1. Listar CSVs en ORIGEN
    csv_files = glob.glob(os.path.join(ORIGIN_DIR, "*.csv"))
    # Tambien xlsx si aplica
    xlsx_files = glob.glob(os.path.join(ORIGIN_DIR, "*.xlsx"))
    all_files = csv_files + xlsx_files
    
    if not all_files:
        print(f"No se encontraron archivos en {ORIGIN_DIR}")
        return

    print(f"Encontrados {len(all_files)} archivos para procesar.")

    for file_path in all_files:
        filename = os.path.basename(file_path)
        print(f"\n==================================================")
        print(f"PROCESANDO: {filename}")
        print(f"==================================================")
        
        # 2. CONVERSION / INGESTION
        # En lugar de copiar crudo, usamos _0conversor.py para convertir a XLSX en LEADS_UNICOS
        # Argumentos: input_file output_dir
        conversor_script = os.path.join(SCRIPTS_DIR, "_0conversor.py")
        print(f"--- Ejecutando {os.path.basename(conversor_script)} con archivo {filename} ---")
        
        try:
            # Llamamos a _0conversor con argumentos
            cmd = [sys.executable, conversor_script, file_path, INPUT_LEADS_DIR]
            subprocess.run(cmd, check=True, capture_output=True, text=True)
            print(f"OK: Conversión completada para {filename}")
        except subprocess.CalledProcessError as e:
            print(f"ERROR: Fallo conversión de {filename}")
            print(e.stderr)
            print("Saltando este archivo...")
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
            # shutil.move(file_path, os.path.join(PROCESSED_DIR, filename))
            
            # Limpieza: No borramos de LEADS_UNICOS para trazabilidad o borramos?
            # Si borramos, _1 en la siguiente iteracion fallaria si _0 falla.
            # Mejor dejar que se acumulen, _1 siempre toma el último.
            pass
        else:
            print(f"CICLO FALLIDO PARA: {filename}. Se detiene este archivo.")
            
        # Pausa para asegurar timestamps distintos
        time.sleep(2)

    print("\nPROCESAMIENTO POR LOTES FINALIZADO.")

if __name__ == "__main__":
    main()
