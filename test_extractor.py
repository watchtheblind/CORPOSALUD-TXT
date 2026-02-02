import os
from openpyxl import load_workbook
from extractor import PDFWorker
from excel_manager import ExcelBox

# 1. Rutas (Usa la carpeta donde tengas varios PDFs de la 1era quincena)
ruta_carpeta_1era = r"E:/Andrea/Desktop/CUADROS TXT/LISTADOS BANCARIOS BENITEZ/1. ORDINARIA/LISTADOS BANCARIOS/1ERA QUINCENA"
ruta_excel = r"E:/Andrea/Desktop/CUADROS TXT/BENITEZ CUADRO TXT.xlsx"

NOMINA_MAPPER = {
    "Empleado Fijo Administrativo": ["EF ADM", "EFA"],
    "Empleado Fijo Enfermeros": ["EF ENFER", "EFE"],
    "Empleado Fijo Médicos": ["EF MEDICO", "EFM"],
    "Empleado Fijo Otros Gremios": ["EF OTROS GRE", "EFOG"],
    "Obrero Fijo": ["OF OBRERO", "OF"],
    "Empleado Contratado Administrativo": ["EC ADM", "ECA"],
    "Empleado Contratado Enfermeros": ["EC ENFER", "ECE"],
    "Empleado Contratado Médicos": ["EC MEDICO", "ECM"],
    "Empleado Contratado Otros Gremios": ["EC OTROS GRE", "ECOG"],
    "Obrero Contratado": ["OC OBRERO", "OC"]
}

# 2. PREPARACIÓN DEL EXCEL
wb = load_workbook(ruta_excel)
ws = wb['CUADRE TXT BENITEZ']

# Creamos la instancia para la 1era Quincena
cuadro_1era = ExcelBox(ws, "1ERA QCNA", NOMINA_MAPPER)

print(f"--- INICIANDO ESCANEO DE CARPETA ---")

# 3. BUCLE DE PROCESAMIENTO
archivos = [f for f in os.listdir(ruta_carpeta_1era) if f.endswith('.pdf')]

for archivo_nombre in archivos:
    ruta_completa = os.path.join(ruta_carpeta_1era, archivo_nombre)
    worker = PDFWorker(ruta_completa)
    
    # --- LLAMADA TEMPORAL DE AUDITORIA ---
    # Esto creará una carpeta 'logs_debug' y llenará un txt por cada PDF
    worker.generar_log_auditoria("logs_debug")
    
    res = worker.process(NOMINA_MAPPER)
    
    if res["success"]:
        # Aquí el ExcelBox hace su magia usando la sigla detectada en el PDF
        cuadro_1era.fill(res['sigla'], res['monto'], res['trabajadores'])
    else:
        print(f"❌ Error en {archivo_nombre}: {res['error']}")

# 4. GUARDAR CAMBIOS
wb.save(ruta_excel)
print("\n✅ ¡Todos los archivos de la carpeta han sido procesados y guardados!")