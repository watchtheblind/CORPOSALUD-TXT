import os
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer

class PDFWorker:
    """
    Módulo encargado de la lectura geométrica de PDFs de nómina.
    Utiliza pdfminer para ubicar datos basados en su posición física en el documento.
    """
    def __init__(self, path):
        self.path = path

    def generar_log_auditoria(self, carpeta_logs="logs_debug"):
        """
        Exporta todo el texto del PDF a un TXT detallando coordenadas (X, Y).
        Útil para depurar cuando un PDF cambia de formato o da resultados inesperados.
        """
        if not os.path.exists(carpeta_logs):
            os.makedirs(carpeta_logs)
            
        nombre_pdf = os.path.basename(self.path)
        ruta_log = os.path.join(carpeta_logs, f"LOG_{nombre_pdf}.txt")
        
        try:
            paginas = list(extract_pages(self.path))
            with open(ruta_log, "w", encoding="utf-8") as f:
                f.write(f"AUDITORIA COMPLETA: {nombre_pdf}\n{'='*60}\n")
                for i, pagina in enumerate(paginas):
                    f.write(f"\n--- PÁGINA {i+1} ---\n")
                    for element in pagina:
                        if isinstance(element, LTTextContainer):
                            # Limpiamos saltos de línea para que el log sea legible
                            texto = " ".join(element.get_text().split())
                            f.write(f"[Y: {round(element.y0, 2)} | X: {round(element.x0, 2)}] -> {texto}\n")
            return True
        except Exception as e:
            print(f"Error generando log: {e}")
            return False

    def process(self, mapper):
        """
        Extrae el nombre del gremio y los totales (monto y trabajadores).
        Utiliza una lógica de 'Última Línea Numérica' para máxima precisión.
        """
        try:
            paginas = list(extract_pages(self.path))
            if not paginas: 
                return {"success": False, "error": "PDF vacío"}
            
            # OPTIMIZACIÓN: Solo procesamos primera y última hoja
            hoja_inicio = paginas[0]
            hoja_final = paginas[-1]
            
            nombre_detectado = None
            
            # --- PASO 1: IDENTIFICACIÓN DEL GREMIO (Página 1) ---
            for element in hoja_inicio:
                if isinstance(element, LTTextContainer):
                    texto = " ".join(element.get_text().split())
                    if texto in mapper:
                        nombre_detectado = texto
                        break
            
            if not nombre_detectado:
                return {"success": False, "error": "Gremio no hallado en encabezado"}

            # --- PASO 2: EXTRACCIÓN DE TOTALES (Última Página) ---
            bloques_finales = []
            for element in hoja_final:
                if isinstance(element, LTTextContainer):
                    bloques_finales.append({
                        "texto": " ".join(element.get_text().split()),
                        "y0": element.y0
                    })
            
            # Ordenamos por Y (de abajo hacia arriba) para priorizar el pie de página
            bloques_finales.sort(key=lambda x: x["y0"])

            monto_final = 0.0
            trabajadores_final = 0
            exito_extraccion = False

            # Recorremos los bloques desde el más bajo (Y menor) buscando números
            for bloque in bloques_finales:
                partes = bloque["texto"].split()
                
                # Criterio: La línea debe tener al menos 2 partes y el penúltimo ser moneda (con coma)
                if len(partes) >= 2 and "," in partes[-2]:
                    try:
                        # Limpieza de formato moneda (Ej: 1.827,82 -> 1827.82)
                        monto_str = partes[-2].replace('.', '').replace(',', '.')
                        monto_final = float(monto_str)
                        
                        # El último elemento es siempre el conteo de trabajadores
                        trabajadores_final = int(partes[-1])
                        
                        exito_extraccion = True
                        break # Detenemos al encontrar el primer bloque numérico válido
                    except ValueError:
                        continue # Si no es numérico, ignoramos y seguimos subiendo

            if not exito_extraccion:
                return {"success": False, "error": "No se hallaron totales numéricos válidos"}

            return {
                "success": True,
                "sigla": mapper[nombre_detectado][0],
                "monto": monto_final,
                "trabajadores": trabajadores_final
            }

        except Exception as e:
            return {"success": False, "error": str(e)}