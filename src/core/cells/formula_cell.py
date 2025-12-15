from src.core.cells.base_cell import BaseCell
from src.api.models import ApiCell

class FormulaCell(BaseCell):
    def __init__(self, api_cell: ApiCell):
        super().__init__(api_cell)
        self.formula = api_cell.formula
        if isinstance(self.formula, str) and not self.formula.startswith('='):
            self.formula = f"={self.formula}"

    def get_value(self):
        return self.formula