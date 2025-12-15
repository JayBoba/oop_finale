from abc import ABC, abstractmethod
from src.api.models import ApiCell

class BaseCell(ABC):
    def __init__(self, api_cell: ApiCell):
        self.row = api_cell.row
        self.column = api_cell.column
        self.api_cell = api_cell

    @abstractmethod
    def get_value(self):
        pass

    @property
    def coordinate(self) -> str:
        return self.api_cell.excel_address