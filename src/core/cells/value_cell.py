from src.core.cells.base_cell import BaseCell
from src.api.models import ApiCell

class ValueCell(BaseCell):
    def __init__(self, api_cell: ApiCell):
        super().__init__(api_cell)
        self.value = api_cell.value

    def get_value(self):
        return self.value