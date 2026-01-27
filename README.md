# CORPOSALUD-TXT

## 🛠 Tech Stack y Librerías

Para lograr una herramienta intuitiva y funcional, estas son las piezas clave:

### 1. Interfaz Gráfica (GUI)

* **Tkinter:** Incluida en la librería estándar de Python. Es ligera y suficiente para herramientas internas.
* **CustomTkinter (Opcional pero recomendado):** Una evolución estética de Tkinter que le dará a tu programa un aspecto moderno (modo oscuro/claro, bordes redondeados) similar a las apps actuales.
* **FileDialog:** Para las ventanas emergentes de "Seleccionar archivo" o "Seleccionar carpeta".

### 2. Procesamiento de PDFs

* **PyMuPDF (fitz):** Actualmente es la más rápida y versátil para extraer texto.
* **pdfplumber:** Excelente si los PDFs contienen tablas, ya que permite extraer datos estructurados con mucha precisión.

### 3. Manejo de Datos y Excel

* **pandas:** Es el estándar para manipular datos. Nos servirá para leer la plantilla de Excel y escribir los resultados de forma eficiente.
* **openpyxl:** La librería que utiliza pandas por debajo para interactuar con archivos `.xlsx`.

### 4. Logging y Utilidades

* **logging:** Para crear ese registro de eventos que mencionas (errores, archivos procesados con éxito, etc.).
* **pathlib:** Para gestionar rutas de archivos de forma inteligente, sin importar si usas Windows, Mac o Linux.

---

## 📂 Propuesta de Estructura de Archivos (POO)

Para mantener el código limpio, seguiremos este esquema:

```text
📁 PDF_Extractor_Tool/
│
├── 📄 main.py              # Punto de entrada de la aplicación
├── 📄 gui.py               # Clase Interfaz (Hereda de Tkinter)
├── 📄 extractor.py         # Clase Procesador (Lógica de extracción de PDFs)
├── 📄 excel_manager.py     # Clase Excel (Lectura/Escritura de la plantilla)
├── 📁 logs/                # Carpeta donde se guardarán los registros .log
├── 📁 assets/              # Iconos o imágenes de la interfaz
└── 📄 README.md            # Documentación del proyecto

```

---

# Extractor de Datos PDF a Excel

Este software permite automatizar la extracción de información específica desde múltiples archivos PDF, consolidando los resultados en una plantilla de Excel predefinida.

### 🚀 Características

* **Interfaz Intuitiva:** Selección de archivos y carpetas mediante explorador.
* **Mapeo Inteligente:** Relación de campos específicos entre PDF y Excel.
* **Sistema de Logs:** Registro detallado de cada acción y error para asegurar la integridad de los datos.
* **Arquitectura POO:** Código modular, fácil de mantener y escalar.

### 📋 Requisitos Previos

* Python 3.8+
* Dependencias: `pip install pandas openpyxl pymupdf customtkinter`

---
