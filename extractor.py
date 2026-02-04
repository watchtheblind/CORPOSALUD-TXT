import os
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer

class PDFWorker:
    def __init__(self, path):
        self.path = path

    def _obtener_texto_limpio(self, pagina):
        """Convierte los elementos de una página en una lista de diccionarios con texto y coordenadas."""
        elementos = []
        for element in pagina:
            if isinstance(element, LTTextContainer):
                elementos.append({
                    "texto": " ".join(element.get_text().split()),
                    "y0": element.y0,
                    "x0": element.x0
                })
        return elementos

    def _identificar_gremio(self, elementos_pag1, mapper):
        """Busca el nombre del gremio dentro de los elementos de la primera página."""
        for el in elementos_pag1:
            if el["texto"] in mapper:
                return el["texto"]
        return None

    def _extraer_totales_economicos(self, elementos_ultima_pag):
        """Analiza la última página para encontrar el bloque de montos y trabajadores."""
        # Ordenamos de abajo hacia arriba (Y menor a mayor)
        elementos_ordenados = sorted(elementos_ultima_pag, key=lambda x: x["y0"])

        for el in elementos_ordenados:
            partes = el["texto"].split()
            # Criterio: Al menos 2 partes y la penúltima debe ser un monto (contiene coma)
            if len(partes) >= 2 and "," in partes[-2]:
                try:
                    monto = float(partes[-2].replace('.', '').replace(',', '.'))
                    trabajadores = int(partes[-1])
                    return monto, trabajadores
                except (ValueError, IndexError):
                    continue
        return None, None

    def process(self, mapper):
        """Orquestador del proceso de extracción."""
        try:
            paginas = list(extract_pages(self.path))
            if not paginas: return {"success": False, "error": "PDF vacío"}

            # 1. Preparar datos de páginas clave
            datos_pag1 = self._obtener_texto_limpio(paginas[0])
            datos_ultima = self._obtener_texto_limpio(paginas[-1])

            # 2. Identificar Gremio
            nombre_gremio = self._identificar_gremio(datos_pag1, mapper)
            if not nombre_gremio:
                return {"success": False, "error": "Gremio no identificado"}

            # 3. Extraer Totales
            monto, trabajadores = self._extraer_totales_economicos(datos_ultima)
            if monto is None:
                return {"success": False, "error": "No se hallaron totales numéricos"}

            # 4. Resultado final
            return {
                "success": True,
                "sigla": mapper[nombre_gremio][0],
                "monto": monto,
                "trabajadores": trabajadores
            }

        except Exception as e:
            return {"success": False, "error": f"Error crítico: {str(e)}"}

    def generar_log_auditoria(self, carpeta_logs="logs_debug"):
        """Método de diagnóstico (se mantiene igual)."""
        if not os.path.exists(carpeta_logs): os.makedirs(carpeta_logs)
        nombre_pdf = os.path.basename(self.path)
        ruta_log = os.path.join(carpeta_logs, f"LOG_{nombre_pdf}.txt")
        try:
            paginas = list(extract_pages(self.path))
            with open(ruta_log, "w", encoding="utf-8") as f:
                f.write(f"AUDITORIA: {nombre_pdf}\n{'='*40}\n")
                for i, pag in enumerate(paginas):
                    elementos = self._obtener_texto_limpio(pag)
                    f.write(f"\n--- PÁGINA {i+1} ---\n")
                    for el in elementos:
                        f.write(f"[Y: {round(el['y0'],2)}] -> {el['texto']}\n")
            return True
        except Exception as e:
            print(f"Error en log: {e}"); return False