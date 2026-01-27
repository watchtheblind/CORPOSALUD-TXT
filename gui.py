import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
from extractor import PDFWorker
from excel_manager import ExcelBox
from openpyxl import load_workbook

# Configuración de apariencia
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Extractor de Nómina POO - Corposalud")
        self.geometry("700x500")

        # Variables de ruta
        self.excel_path = ctk.StringVar()
        self.pdf_folder = ctk.StringVar()

        self._setup_gui()

    def _setup_gui(self):
        self.grid_columnconfigure(0, weight=1)
        
        # --- TÍTULO ---
        self.label_title = ctk.CTkLabel(self, text="Panel de Extracción de Nómina", font=ctk.CTkFont(size=20, weight="bold"))
        self.label_title.pack(pady=20)

        # --- SELECCIÓN DE EXCEL ---
        self.btn_excel = ctk.CTkButton(self, text="1. Seleccionar Plantilla Excel", command=self._select_excel)
        self.btn_excel.pack(pady=10)
        self.lbl_excel = ctk.CTkLabel(self, textvariable=self.excel_path, font=ctk.CTkFont(size=10))
        self.lbl_excel.pack()

        # --- SELECCIÓN DE CARPETA PDF ---
        self.btn_pdfs = ctk.CTkButton(self, text="2. Seleccionar Carpeta de PDFs", command=self._select_folder)
        self.btn_pdfs.pack(pady=10)
        self.lbl_folder = ctk.CTkLabel(self, textvariable=self.pdf_folder, font=ctk.CTkFont(size=10))
        self.lbl_folder.pack()

        # --- LOG DE ACTIVIDAD ---
        self.textbox_log = ctk.CTkTextbox(self, width=600, height=150)
        self.textbox_log.pack(pady=20, padx=20)
        self._log("Esperando selección de archivos...")

        # --- BOTÓN PROCESAR ---
        self.btn_run = ctk.CTkButton(self, text="INICIAR PROCESAMIENTO", fg_color="green", hover_color="#006400", command=self._start_process)
        self.btn_run.pack(pady=20)

    def _log(self, message):
        self.textbox_log.insert("end", f"> {message}\n")
        self.textbox_log.see("end")

    def _select_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.excel_path.set(path)
            self._log(f"Excel seleccionado: {os.path.basename(path)}")

    def _select_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.pdf_folder.set(path)
            self._log(f"Carpeta de PDFs seleccionada: {path}")

    def _start_process(self):
        if not self.excel_path.get() or not self.pdf_folder.get():
            messagebox.showwarning("Atención", "Por favor selecciona ambos: el Excel y la carpeta de PDFs.")
            return

        self._log("Iniciando proceso...")
        
        try:
            # 1. Cargar el libro de Excel una sola vez
            wb = load_workbook(self.excel_path.get())
            ws = wb.active # O busca la hoja por nombre si es necesario
            
            # 2. Instanciar los cuadros (Usando la lógica de anclaje de ExcelBox)
            # Aquí aplicamos el principio de responsabilidad única
            cuadro_1era = ExcelBox(ws, "1ERA QCNA")
            
            # 3. Listar archivos y procesar
            files = [f for f in os.listdir(self.pdf_folder.get()) if f.endswith(".pdf")]
            
            for f in files:
                full_path = os.path.join(self.pdf_folder.get(), f)
                worker = PDFWorker(full_path)
                data = worker.process()
                
                if data:
                    self._log(f"Procesando {f}: {data['gremio']} - {data['monto']} Bs.")
                    # 4. Llenar en el cuadro correspondiente (Elegimos 1ERA QCNA para esta prueba)
                    cuadro_1era.fill(data['gremio'], data['monto'], data['trabajadores'])
                else:
                    self._log(f"Omitido {f}: No cumple criterios o grupo no mapeado.")

            # 5. Guardar cambios
            wb.save(self.excel_path.get())
            self._log("!!! PROCESO FINALIZADO CON ÉXITO !!!")
            messagebox.showinfo("Éxito", "El archivo Excel ha sido actualizado.")

        except Exception as e:
            self._log(f"ERROR CRÍTICO: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error inesperado: {e}")

if __name__ == "__main__":
    app = App()
    app.mainloop()