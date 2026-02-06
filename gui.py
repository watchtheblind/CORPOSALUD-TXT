import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
from openpyxl import load_workbook

# Importamos tus clases motor
from extractor import PDFWorker
from excel_manager import ExcelBox

# Configuración visual (la tuya)
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("GENERADOR DE CUADROS TXT")
        self.geometry("900x650")
        
        # --- VARIABLES DE RUTA (Tus variables originales) ---
        self.path_q1 = ctk.StringVar()
        self.path_q2 = ctk.StringVar()
        self.path_vacacional = ctk.StringVar()
        self.path_nocturno = ctk.StringVar()
        self.path_domingos = ctk.StringVar()
        self.path_primas = ctk.StringVar()
        self.path_excel = ctk.StringVar()

        # Mapeo de Periodo -> (Variable de UI, Nombre exacto en Excel)
        self.config_mapeo = {
            "1ERA QCNA": (self.path_q1, "1ERA QCNA"),
            "2DA QCNA": (self.path_q2, "2DA QCNA"),
            "RECREACIONAL": (self.path_vacacional, "RECREACIONAL"),
            "BONO NOCTURNO": (self.path_nocturno, "BONO NOCTURNO"),
            "DOMINGOS Y FERIADOS": (self.path_domingos, "DOMINGOS Y FERIADOS"),
            "PRIMAS": (self.path_primas, "PRIMAS")
        }

        self._build_ui()

    # ... [Tus métodos _build_ui, _create_section_header, _create_input_block y _open_dialog se mantienen IGUAL] ...

    def _process_files(self):
        """Lógica de integración para procesar todos los cuadros."""
        ruta_plantilla = self.path_excel.get()
        
        if not ruta_plantilla:
            messagebox.showerror("Falta Archivo", "Por favor, seleccione la plantilla Excel.")
            return

        # Diccionario Mapper que ya perfeccionamos
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

        try:
            print("--- INICIANDO PROCESO DE CARGA ---")
            wb = load_workbook(ruta_plantilla)
            ws = wb.active # O wb['NOMBRE_DE_TU_HOJA']

            # Iteramos por cada sección de tu UI
            for periodo, (var_ruta, ancla_excel) in self.config_mapeo.items():
                ruta_carpeta = var_ruta.get()
                
                if not ruta_carpeta or not os.path.exists(ruta_carpeta):
                    print(f"⏩ Saltando {periodo}: Carpeta no seleccionada.")
                    continue

                print(f"📦 Procesando {periodo}...")
                # Instanciamos el cuadro de Excel para esta ancla
                cuadro = ExcelBox(ws, ancla_excel, NOMINA_MAPPER)
                
                # Procesamos PDFs en la carpeta
                archivos_pdf = [f for f in os.listdir(ruta_carpeta) if f.lower().endswith('.pdf')]
                
                for archivo in archivos_pdf:
                    ruta_pdf = os.path.join(ruta_carpeta, archivo)
                    worker = PDFWorker(ruta_pdf)
                    
                    resultado = worker.process(NOMINA_MAPPER)
                    
                    if resultado["success"]:
                        cuadro.fill(
                            resultado['sigla'], 
                            resultado['monto'], 
                            resultado['trabajadores'], 
                            resultado['estado']
                        )
                    else:
                        print(f"❌ Error en {archivo}: {resultado['error']}")

            # Guardar resultados
            wb.save(ruta_plantilla)
            messagebox.showinfo("Proceso Finalizado", "Se han cargado todos los datos con éxito.")
            print("--- PROCESO COMPLETADO ---")

        except PermissionError:
            messagebox.showerror("Error", "Cierre el archivo Excel antes de procesar.")
        except Exception as e:
            messagebox.showerror("Error Crítico", f"Detalle: {str(e)}")

    # (Incluye aquí tus métodos helper originales: _create_section_header, etc.)