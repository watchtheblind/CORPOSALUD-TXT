import customtkinter as ctk
from tkinter import filedialog

# Configuración visual
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Ventana Principal
        self.title("GENERADOR DE CUADROS TXT")
        self.geometry("900x650")
        
        # --- VARIABLES DE RUTA ---
        # Izquierda
        self.path_q1 = ctk.StringVar()
        self.path_q2 = ctk.StringVar()
        self.path_vacacional = ctk.StringVar()
        self.path_nocturno = ctk.StringVar()
        
        # Derecha
        self.path_domingos = ctk.StringVar()
        self.path_primas = ctk.StringVar()
        self.path_excel = ctk.StringVar()

        self._build_ui()

    def _build_ui(self):
        # TÍTULO PRINCIPAL
        lbl_title = ctk.CTkLabel(
            self, text="GENERADOR DE CUADROS TXT", 
            font=ctk.CTkFont(size=26, weight="bold")
        )
        lbl_title.pack(pady=20)

        # CONTENEDOR PRINCIPAL (2 COLUMNAS)
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Configurar columnas (50% cada una)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)

        # ==================================================
        # SECCIÓN IZQUIERDA
        # ==================================================
        left_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        left_frame.grid(row=0, column=0, sticky="nsew", padx=20)

        # --- BLOQUE 1: QUINCENAS ---
        # Icono y Label Secundario
        self._create_section_header(left_frame, "📅", "Quincenas")
        
        # Label Terciario + Campo 1 (Primera Quincena)
        self._create_input_block(left_frame, "Primera Quincena", self.path_q1, is_folder=True)
        
        # Label Terciario + Campo 2 (Segunda Quincena)
        self._create_input_block(left_frame, "Segunda Quincena", self.path_q2, is_folder=True)

        # Separador visual
        ctk.CTkFrame(left_frame, height=2, fg_color="gray30").pack(fill="x", pady=20)

        # --- BLOQUE 2: BONOS ---
        # Icono y Label Secundario
        self._create_section_header(left_frame, "💵", "Bonos")

        # Label Terciario + Campo (Bono Vacacional)
        self._create_input_block(left_frame, "Bono Vacacional", self.path_vacacional, is_folder=True)

        # Label Terciario + Campo (Bono Nocturno)
        self._create_input_block(left_frame, "Bono Nocturno", self.path_nocturno, is_folder=True)


        # ==================================================
        # SECCIÓN DERECHA
        # ==================================================
        right_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        right_frame.grid(row=0, column=1, sticky="nsew", padx=20)

        # --- BLOQUE 3: DOMINGOS ---
        # Icono y Label Secundario
        self._create_section_header(right_frame, "☀️", "Domingo y Feriado")
        
        # Campo directo (sin label terciario según instrucción)
        self._create_input_block(right_frame, None, self.path_domingos, is_folder=True)

        # Separador
        ctk.CTkFrame(right_frame, height=2, fg_color="gray30").pack(fill="x", pady=20)

        # --- BLOQUE 4: PRIMAS ---
        # Label Secundario
        self._create_section_header(right_frame, "👤", "Prima por hijos y becas")
        
        # Campo directo
        self._create_input_block(right_frame, None, self.path_primas, is_folder=True)

        # Espacio flexible para empujar lo siguiente al fondo si fuera necesario, 
        # pero aquí pondremos la plantilla Excel abajo como pediste.
        
        ctk.CTkFrame(right_frame, height=2, fg_color="gray30").pack(fill="x", pady=20)

        # --- BLOQUE 5: PLANTILLA EXCEL ---
        # Label Secundario
        lbl_excel = ctk.CTkLabel(right_frame, text="Plantilla de Excel", font=ctk.CTkFont(size=18, weight="bold"))
        lbl_excel.pack(anchor="w", pady=(10, 5))
        
        # Campo Excel (Selección de ARCHIVO, no carpeta)
        self._create_input_block(right_frame, None, self.path_excel, is_folder=False)

        # --- BOTÓN PROCESAR ---
        self.btn_process = ctk.CTkButton(
            right_frame,
            text="PROCESAR ARCHIVOS",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=50,
            fg_color="#27ae60", hover_color="#2ecc71",
            command=self._process_files
        )
        self.btn_process.pack(fill="x", pady=(40, 0))


    # --- HELPERS PARA NO REPETIR CÓDIGO ---

    def _create_section_header(self, parent, icon, text):
        """Crea el encabezado con icono y texto secundario"""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", pady=(10, 5))
        
        lbl_icon = ctk.CTkLabel(frame, text=icon, font=ctk.CTkFont(size=28))
        lbl_icon.pack(side="left", padx=(0, 10))
        
        lbl_text = ctk.CTkLabel(frame, text=text, font=ctk.CTkFont(size=18, weight="bold"))
        lbl_text.pack(side="left")

    def _create_input_block(self, parent, label_text, variable, is_folder=True):
        """Crea el label terciario (opcional) + Entry + Botón de 3 puntos"""
        
        # Label Terciario (si existe)
        if label_text:
            lbl = ctk.CTkLabel(parent, text=label_text, font=ctk.CTkFont(size=12))
            lbl.pack(anchor="w", pady=(5, 0))

        # Contenedor Entry + Botón
        input_frame = ctk.CTkFrame(parent, fg_color="transparent")
        input_frame.pack(fill="x", pady=(0, 10))

        entry = ctk.CTkEntry(
            input_frame, 
            textvariable=variable, 
            placeholder_text="Seleccione carpeta..." if is_folder else "Seleccione archivo Excel...",
            state="readonly"
        )
        entry.pack(side="left", fill="x", expand=True, padx=(0, 5))

        btn = ctk.CTkButton(
            input_frame, 
            text="...", 
            width=40, 
            fg_color="#34495e",
            command=lambda: self._open_dialog(variable, is_folder)
        )
        btn.pack(side="right")

    def _open_dialog(self, variable, is_folder):
        if is_folder:
            path = filedialog.askdirectory()
        else:
            path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xlsm")])
        
        if path:
            variable.set(path)

    def _process_files(self):
        print("--- INICIANDO PROCESO ---")
        print(f"1era Quincena: {self.path_q1.get()}")
        print(f"2da Quincena: {self.path_q2.get()}")
        print(f"Vacacional: {self.path_vacacional.get()}")
        print(f"Nocturno: {self.path_nocturno.get()}")
        print(f"Domingos: {self.path_domingos.get()}")
        print(f"Primas: {self.path_primas.get()}")
        print(f"Excel Base: {self.path_excel.get()}")

if __name__ == "__main__":
    app = App()
    app.mainloop()