from enum import Enum
from typing import *
from datetime import datetime
from pydantic import BaseModel, Field, field_validator

class CellType(str, Enum):
    VALUE = "value"
    FOMRULA = "formula"
    LINK = "link"
    EMPTY = "empty"

class FormatType(str, Enum):
    NUMBER = "number"
    CURRENCY = "currency"
    PERCENTAGE = "percentage"
    DATE = "date"
    TEXT = "text"

class CellReference(BaseModel):
    table_id: str
    cell_address: str
    sheet_name: Optional[str] = None

    def __str__(self) -> str:
        if self.sheet_name:
            return f"'{self.sheet_name}'!{self.cell_address}"
        return f"{self.table_id}!{self.cell_address}"
    
class ApiCell(BaseModel):
    id: str
    row: int #потом подумаю как сделать лимиты как в экселе, для строк от 1 до 1048576
    column: int # для столбцов максимум 16384
    value: Optional[Any] = None
    cell_type: CellType
    formula: Optional[Any] = None
    format_type: Optional[FormatType] = None
    reference: List[CellReference] = Field(default_factory=list)
    metadata: Dict[str, Any] = Field(default_factory=dict)

    @field_validator('references')
    def validate_references(cls, v, info):
        if info.data.get('cell_type') == CellType.LINK and not v:
            pass
        return v
    
    @property
    def excel_address(self) -> str:
        column_letter = self._column_to_letter(self.column)
        return f"{column_letter}{self.row}"
    
    @staticmethod
    def _column_to_letter(col: int) -> str:
        result = ""
        while col>0:
            col, remainder = divmod(col -1,  26)
            result = chr(65 + remainder) + result 
        return result
    
class ApiTable(BaseModel):
    id: str
    name: str
    description: Optional[str] = None
    cells: List[ApiCell] = Field(default_factory=list)
    metadata: Dict[str, Any] = Field(default_factory=dict)
    created_at: datetime = Field(default_factory=datetime.now)
    updated_at: datetime = Field(default_factory=datetime.now)
    
    @property
    def formula_cells(self) -> List[ApiCell]:
        return [cell for cell in self.cells if cell.cell_type == CellType.FORMULA]
    
    @property
    def link_cells(self) -> List[ApiCell]:
        return [cell for cell in self.cells if cell.cell_type == CellType.LINK]

class ApiResponse(BaseModel):
    success: bool
    data: Optional[Union[Dict, List]] = None
    error: Optional[str] = None
    request_id: Optional[str] = None
    timestamp: datetime = Field(default_factory=datetime.now)