from openpyxl import load_workbook
from openpyxl.styles import Border
class ExcelBox:
    """O de SOLID: Maneja la lógica de un cuadro. Se instancia por sección."""
    
    def __init__(self, sheet, anchor_name):
        self.sheet = sheet
        self.anchor_name = anchor_name
        self.anchor_cell = self._find_anchor()
        self.data_rows = self._map_box_structure()

    def _find_anchor(self):
        """Busca la celda que dice '1ERA QCNA', '2DA QCNA', etc."""
        for row in self.sheet.iter_rows():
            for cell in row:
                if cell.value == self.anchor_name:
                    return cell
        return None

    def _map_box_structure(self):
        """
        Algoritmo de bordes: Identifica qué filas pertenecen al cuadro
        basándose en si tienen bordes y el nombre del gremio.
        """
        mapping = {}
        if not self.anchor_cell:
            return mapping

        # Empezamos la búsqueda desde la columna de 'GREMIOS' 
        # (Basado en tu wireframe, el gremio suele estar a la derecha del anclaje o en la misma fila)
        start_row = self.anchor_cell.row
        
        # Escaneamos hacia abajo hasta que no haya más bordes
        for row_idx in range(start_row, start_row + 20): # Límite razonable de 20 filas por cuadro
            # Buscamos en la columna de GREMIOS (ejemplo: columna D es la 4)
            # Aquí ajustaremos según el wireframe real
            for col_idx in range(1, 15): 
                cell = self.sheet.cell(row=row_idx, column=col_idx)
                if cell.value in NOMINA_MAPPER.values():
                    # Si tiene borde, es parte de nuestro cuadro
                    if cell.border.left.style or cell.border.bottom.style:
                        mapping[cell.value] = row_idx
        return mapping

    def fill(self, gremio, monto, cant):
        """Escribe los datos en las columnas de Cantidad y Listado Banco."""
        if gremio in self.data_rows:
            target_row = self.data_rows[gremio]
            # TODO: Localizar columnas exactas de 'CANT. DE TRABAJADORES' y 'LISTADOS DE BANCO'
            # Por ahora, un placeholder:
            print(f"Escribiendo en {self.anchor_name} -> {gremio}: {cant} trab, {monto} Bs.")