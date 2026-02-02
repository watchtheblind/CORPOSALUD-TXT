from openpyxl import load_workbook 
from extractor import PDFWorker
from excel_manager import ExcelBox

# 1. Rutas y Diccionario
ruta_pdf = r"E:/Andrea/Desktop/CUADROS TXT/LISTADOS BANCARIOS BENITEZ/1. ORDINARIA/LISTADOS BANCARIOS/1ERA QUINCENA/ECACA0025.pdf"
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

# 2. PROCESO DE EXTRACCIÓN
worker = PDFWorker(ruta_pdf)
res = worker.process(NOMINA_MAPPER)

if res["success"]:
    wb = load_workbook(ruta_excel)
    ws = wb['CUADRE TXT BENITEZ']
    
    # 3. INSTANCIAS DINÁMICAS
    # Recuerda que en el ExcelBox que hicimos, ahora pasamos el mapper al __init__
    cuadro_1era = ExcelBox(ws, "1ERA QCNA", NOMINA_MAPPER)
    
    # 4. LLENADO
    # res['sigla'] ya contiene el valor que ExcelBox buscará en las filas mapeadas
    cuadro_1era.fill(res['sigla'], res['monto'], res['trabajadores'])
    
    wb.save(ruta_excel)
    print("¡Proceso completado con éxito!")
else:
    print(f"Error en extracción: {res['error']}")