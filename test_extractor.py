
from extractor import PDFWorker
from excel_manager import ExcelBox  # Usamos tu clase original

# 1. Rutas y Diccionario
ruta_pdf = r"C:/Users/luis/Documents/CUADROS TXT/LISTADOS BANCARIOS BENITEZ/1. ORDINARIA/LISTADOS BANCARIOS/1ERA QUINCENA/ECACA0025.pdf"
ruta_excel = r"C:/Users/luis/Desktop/CORPOSALUD-TXT/BENITEZ CUADRO TXT.xlsx"

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
    sigla = res['sigla']
    monto = res['monto']
    trabajadores = res['trabajadores']
    
    print(f"✅ PDF Extraído: {sigla} | Bs. {monto} | Trab: {trabajadores}")

    # 3. INTEGRACIÓN CON TU EXCEL_MANAGER
    manager = ExcelBox(ruta_excel)
    if manager.load_workbook(): # Usamos tu método de carga
        ws = manager.get_sheet('CUADRE TXT BENITEZ')
        
        if ws:
            # Definimos columnas de 1ERA QCNA (Ajusta según tu Excel real)
            COL_SIGLA = 2   # Columna B
            COL_LISTADO = 5 # Columna E (Listado Banc)
            COL_TRAB = 6    # Columna F (Trabajadores)

            encontrado = False
            # Buscamos la fila de la sigla
            for row in range(1, ws.max_row + 1):
                if ws.cell(row=row, column=COL_SIGLA).value == sigla:
                    # Escribimos usando la instancia de la hoja
                    ws.cell(row=row, column=COL_LISTADO).value = monto
                    ws.cell(row=row, column=COL_TRAB).value = trabajadores
                    encontrado = True
                    break
            
            if encontrado:
                manager.save_workbook() # Usamos tu método de guardado
                print(f"🚀 Excel actualizado exitosamente para {sigla}.")
            else:
                print(f"❌ No se encontró la sigla {sigla} en la columna {COL_SIGLA}")
        else:
            print("❌ No se pudo encontrar la hoja 'CUADRE TXT BENITEZ'")
else:
    print(f"❌ Error en PDF: {res['error']}")