import os
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer

class PDFWorker:
    def __init__(self, path):
        self.path = path

    def _get_page_elements(self, page):
        """Convierte una página de PDF en lista de diccionarios (Atomicidad)."""
        elements = []
        for element in page:
            if isinstance(element, LTTextContainer):
                elements.append({
                    "texto": " ".join(element.get_text().split()),
                    "y0": element.y0,
                    "x0": element.x0
                })
        return elements

    def _detectar_estado(self, elementos_p1):
        """Detecta si el reporte es de activos o inactivos."""
        for el in elementos_p1:
            if "Cuenta Bancaria Inactiva" in el["texto"]:
                return "INACTIVO"
        return "ACTIVO"

    def _find_gremio(self, elementos_pag1, mapper):
        """Busca el nombre del gremio (antes llamado _identificar_gremio)."""
        for el in elementos_pag1:
            if el["texto"] in mapper:
                return el["texto"]
        return None

    def _parse_totals(self, elementos_ultima_pag):
        """Extrae montos y trabajadores (antes llamado _extraer_totales_economicos)."""
        elementos_ordenados = sorted(elementos_ultima_pag, key=lambda x: x["y0"])
        for el in elementos_ordenados:
            partes = el["texto"].split()
            if len(partes) >= 2 and "," in partes[-2]:
                try:
                    monto = float(partes[-2].replace('.', '').replace(',', '.'))
                    trabajadores = int(partes[-1])
                    return monto, trabajadores
                except (ValueError, IndexError):
                    continue
        return None, None

    def process(self, mapper):
        """Orquestador con nombres de métodos unificados."""
        try:
            pages = list(extract_pages(self.path))
            if not pages: return {"success": False, "error": "PDF vacío"}

            # Ahora todos los nombres coinciden
            datos_pag1 = self._get_page_elements(pages[0])
            datos_ultima = self._get_page_elements(pages[-1])

            nombre_gremio = self._find_gremio(datos_pag1, mapper)
            estado_nomina = self._detectar_estado(datos_pag1) 

            monto, trabajadores = self._parse_totals(datos_ultima)
            
            if monto is None:
                return {"success": False, "error": "No se hallaron totales"}

            return {
                "success": True,
                "sigla": mapper[nombre_gremio][0],
                "monto": monto,
                "trabajadores": trabajadores,
                "estado": estado_nomina 
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def generar_log_auditoria(self, carpeta_logs="logs_debug"):
        if not os.path.exists(carpeta_logs): os.makedirs(carpeta_logs)
        nombre_pdf = os.path.basename(self.path)
        ruta_log = os.path.join(carpeta_logs, f"LOG_{nombre_pdf}.txt")
        try:
            paginas = list(extract_pages(self.path))
            with open(ruta_log, "w", encoding="utf-8") as f:
                f.write(f"AUDITORIA: {nombre_pdf}\n{'='*40}\n")
                for i, pag in enumerate(paginas):
                    # Usamos el mismo método unificado aquí también
                    elementos = self._get_page_elements(pag)
                    f.write(f"\n--- PÁGINA {i+1} ---\n")
                    for el in elementos:
                        f.write(f"[Y: {round(el['y0'],2)}] -> {el['texto']}\n")
            return True
        except Exception as e:
            print(f"Error en log: {e}"); return False