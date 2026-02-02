import re
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer

class PDFWorker:
    def __init__(self, path):
        self.path = path

    def generar_log_auditoria(self, carpeta_logs):
        """Metodo temporal para volcar el contenido del PDF a un TXT"""
        import os
        if not os.path.exists(carpeta_logs):
            os.makedirs(carpeta_logs)
            
        nombre_pdf = os.path.basename(self.path)
        ruta_log = os.path.join(carpeta_logs, f"LOG_{nombre_pdf}.txt")
        
        try:
            paginas = list(extract_pages(self.path))
            with open(ruta_log, "w", encoding="utf-8") as f:
                f.write(f"AUDITORIA DE COORDENADAS: {nombre_pdf}\n")
                f.write("="*50 + "\n\n")
                
                for element in paginas[0]:
                    if isinstance(element, LTTextContainer):
                        texto = " ".join(element.get_text().split())
                        x0 = round(element.x0, 2)
                        y0 = round(element.y0, 2)
                        # Escribimos el texto y su ubicación exacta
                        f.write(f"[Y0: {y0} | X0: {x0}] -> '{texto}'\n")
            return True
        except Exception as e:
            print(f"Error generando log para {nombre_pdf}: {e}")
            return False

    def process(self, mapper):

        try:
            paginas = list(extract_pages(self.path))
            if not paginas: return {"success": False, "error": "PDF vacío"}
            
            elementos = []
            for element in paginas[0]:
                if isinstance(element, LTTextContainer):
                    elementos.append({
                        "texto": " ".join(element.get_text().split()),
                        "x0": round(element.x0, 2),
                        "y0": round(element.y0, 2)
                    })

            # 1. VERIFICADOR DE ACTIVIDAD
            if not any("Personal Con Cuenta Bancaria Activa" in el["texto"] for el in elementos):
                return {"success": False, "error": "Documento no válido"}

            # 2. UBICAR EL GRUPO
            nombre_detectado = next((el["texto"] for el in elementos if el["texto"] in mapper), None)
            
            if not nombre_detectado:
                return {"success": False, "error": "No se halló la clave en el mapper"}

            # 3. CAPTURAR TOTALES (X0: 563.88 Y0: 450.61)
            monto_final = 0.0
            trabajadores_final = 0
            
            for el in elementos:
                # Buscamos el bloque que está en la coordenada Y de los totales
                if abs(el["y0"] - 450.61) < 1.0 and el["x0"] > 500:
                    # El texto es '175,38 2'
                    partes = el["texto"].split() 
                    if len(partes) >= 2:
                        # Limpiamos el monto: quitamos puntos de miles y cambiamos coma por punto
                        monto_str = partes[0].replace('.', '').replace(',', '.')
                        monto_final = float(monto_str)
                        # El último elemento suele ser el total de trabajadores
                        trabajadores_final = int(partes[-1])
                        break

            return {
                "success": True,
                "nombre_leido": nombre_detectado,
                "sigla": mapper[nombre_detectado][0],
                "monto": monto_final,
                "trabajadores": trabajadores_final
            }

        except Exception as e:
            return {"success": False, "error": str(e)}