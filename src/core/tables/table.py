from typing import List, Optional, Dict, Any
from src.api.models import ApiTable, ApiCell, CellType
from src.core.cells.base_cell import BaseCell
from src.core.cells.value_cell import ValueCell
from src.core.cells.formula_cell import FormulaCell
from src.core.cells.reference_handler import ReferenceHandler

class Table:
    def __init__(self, api_table: ApiTable):
        self.id = api_table.id
        self.name = api_table.name
        self.cells: List[BaseCell] = []
        self._process_cells(api_table.cells)

    def _process_cells(self, api_cells: List[ApiCell]):
        for api_cell in api_cells:
            cell = self._create_cell(api_cell)
            if cell:
                self.cells.append(cell)

    def _create_cell(self, api_cell: ApiCell) -> Optional[BaseCell]:
        if api_cell.cell_type == CellType.VALUE:
            return ValueCell(api_cell)
        elif api_cell.cell_type == CellType.FORMULA:
            return FormulaCell(api_cell)
        elif api_cell.cell_type == CellType.LINK:
            return ReferenceHandler(api_cell)
        elif api_cell.cell_type == CellType.EMPTY:
            return None
        else:
            return ValueCell(api_cell)

    def get_linked_table_ids(self) -> List[str]:
        ids = set()
        for cell in self.cells:
            if isinstance(cell, ReferenceHandler):
                ids.update(cell.get_referenced_table_ids())
        return list(ids)