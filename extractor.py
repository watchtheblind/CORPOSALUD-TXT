import fitz  # PyMuPDF
import re
import os
NOMINA_MAPPER = {
    "Empleado Fijo Administrativo": "EF ADM",
    "Empleado Fijo Enfermeros": "EF ENFER",
    "Empleado Fijo Médicos": "EF MEDICO",
    "Empleado Fijo Otros Gremios": "EF OTROS GRE",
    "Obrero Fijo": "OF OBRERO",
    "Empleado Contratado Administrativo": "EC ADM",
    "Empleado Contratado Enfermeros": "EC ENFER",
    "Empleado Contratado Médicos": "EC MEDICO",
    "Empleado Contratado Otros Gremios": "EC OTROS GRE",
    "Obrero Contratado": "OC OBRERO"
}

VALID_HEADER = "Personal Con Cuenta Bancaria Activa"
class PDFWorker:
    """S de SOLID: Responsabilidad única - Extraer datos de PDF."""
    
    def __init__(self, filepath):
        self.filepath = filepath
        
        self.filename = os.path.basename(filepath)

    def process(self):
        """Retorna un dict con los datos si es válido, sino None."""
        try:
            doc = fitz.open(self.filepath)
            text = "".join([page.get_text() for page in doc])
            
            # Validación de seguridad
            if VALID_HEADER not in text:
                return None

            # Extracción de Grupo de Nómina
            grupo_match = re.search(r"Grupo de Nómina:\s*(.*)", text)
            grupo_raw = grupo_match.group(1).strip() if grupo_match else None
            
            # Traducción mediante el mapper
            gremio_excel = NOMINA_MAPPER.get(grupo_raw)
            
            if not gremio_excel:
                return None

            # Extracción de Totales (buscamos la última ocurrencia)
            monto_matches = re.findall(r"Total General Bs\.:\s*([\d\.,]+)", text)
            trab_matches = re.findall(r"Total Trabajadores:\s*(\d+)", text)
            
            return {
                "gremio": gremio_excel,
                "monto": float(monto_matches[-1].replace('.', '').replace(',', '.')) if monto_matches else 0.0,
                "trabajadores": int(trab_matches[-1]) if trab_matches else 0
            }
        except Exception as e:
            print(f"Error procesando {self.filename}: {e}")
            return None