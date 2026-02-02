from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
class ExcelBox:
    # VARIABLES DE EXPANSIÓN (Fáciles de modificar aquí)
    FILAS_HACIA_ARRIBA = 5
    FILAS_HACIA_ABAJO  = 10  # Le pongo 15 para que cubra todos los gremios del mapper
    COLUMNAS_DERECHA   = 9

    def __init__(self, sheet, anchor_name, mapper):
        self.sheet = sheet
        self.anchor_name = anchor_name
        self.mapper = mapper
        self.anchor_cell = self._find_anchor()
        
        self.col_indices = {} 
        self.data_rows = {}    
        
        if self.anchor_cell:
            self._map_box_structure_manual()

    def _find_anchor(self):
        for row in self.sheet.iter_rows():
            for cell in row:
                if str(cell.value).strip() == self.anchor_name:
                    return cell
        return None

    def _map_box_structure_manual(self):
        r_orig = self.anchor_cell.row
        c_orig = self.anchor_cell.column
        
        # --- AJUSTE DE RANGOS ---
        # Si el ancla es 1ERA QCNA y los gremios están debajo, 
        # asegúrate de que FILAS_HACIA_ABAJO sea suficiente (ej. 20)
        fila_inicio = r_orig - self.FILAS_HACIA_ARRIBA
        fila_fin    = r_orig + self.FILAS_HACIA_ABAJO
        col_fin     = c_orig + self.COLUMNAS_DERECHA

        # 1. IDENTIFICAR COLUMNAS (Soportando celdas combinadas)
        for r in range(fila_inicio, r_orig + 5): # Buscamos un poco más abajo del ancla también
            for c in range(c_orig, col_fin + 1):
                cell = self.sheet.cell(row=r, column=c)
                val = str(cell.value or "").upper().strip()
                
                # Si la celda está vacía pero es parte de una combinación, openpyxl a veces devuelve None
                # Buscamos los textos clave:
                if "GREMIOS" in val:
                    self.col_indices['gremios'] = c
                if "LISTADOS DE BANCO" in val:
                    self.col_indices['monto'] = c
                if "CANT, DE TRABAJADORES" in val:
                    self.col_indices['cant'] = c

        # 2. MAPEADO DE FILAS (El truco para celdas combinadas)
        if 'gremios' in self.col_indices:
            col_g = self.col_indices['gremios']
            siglas_validas = [v[0] for v in self.mapper.values()]
            
            for r in range(fila_inicio, fila_fin + 1):
                cell = self.sheet.cell(row=r, column=col_g)
                val = cell.value
                
                # --- TRUCO CELDAS COMBINADAS ---
                # Si la celda es None, verificamos si es una celda combinada
                if val is None:
                    for merged_range in self.sheet.merged_cells.ranges:
                        if cell.coordinate in merged_range:
                            # Tomamos el valor de la celda superior izquierda del rango combinado
                            val = merged_range.start_cell.value
                            break
                
                val_str = str(val or "").strip()
                
                if val_str in siglas_validas:
                    self.data_rows[val_str] = r
        print(f"📍 Área: {get_column_letter(c_orig)} a {get_column_letter(col_fin)}")
        print(f"📍 Filas: {fila_inicio} a {fila_fin}")
        print(f"✅ Se mapearon {len(self.data_rows)} gremios dentro del rango.")

    def fill(self, sigla, monto, cant):
        if sigla in self.data_rows:
            row = self.data_rows[sigla]
            col_m = self.col_indices.get('monto')
            col_c = self.col_indices.get('cant')
            
            if col_m: self.sheet.cell(row=row, column=col_m).value = monto
            if col_c: self.sheet.cell(row=row, column=col_c).value = cant
            
            print(f"✍️ [OK] {self.anchor_name}: {sigla} -> Celda {get_column_letter(col_m)}{row}")
        else:
            print(f"⚠️ [!] {sigla} no está en el rango de {self.anchor_name}")