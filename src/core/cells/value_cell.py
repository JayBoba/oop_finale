from typing import Any, Dict, Optional
from decimal import Decimal, InvalidOperation

from .base_cell import BaseCell, ExcelAddress, CellType, FormatType
from src.utils.validators import validate_numeric_value #TBD

class ValueCell(BaseCell):
     def __init__(
        self,
        address: ExcelAddress,
        value: Any,
        format_type: Optional[FormatType] = None
    ):
        super().__init__(address, value, CellType.VALUE, format_type)
        self._validate_value(value)
        
        '''def _validate_value(self, value: Any):
            if self._format_type == FormatType.NUMBER:
                validate_numeric_value(value)
            elif self._format_type == FormatType.PERCENTAGE:
                try:
                    float_val = float(value)
                    if not 0 <= float_val <= 1:
                        raise ValueError(f"Percentage must be between 0 and 1, got {float_val}")
                except (ValueError, TypeError):
                    pass'''

        def _do_evaluate(self, context: Dict[str, Any]) -> Any:
            if self._format_type in [FormatType.NUMBER, FormatType.CURRENCY]:
                try:
                    return Decimal(str(self._raw_value))
                except InvalidOperation:
                    return self._raw_value
            
            if self._format_type == FormatType.PERCENTAGE:
                return self._raw_value
            
            return self._raw_value
    
        def get_excel_formula(self) -> None:
            return None

        def __add__(self, other: Any) -> Any:
            if isinstance(other, ValueCell):
                return self.evaluate() + other.evaluate()
            try:
                return self.evaluate() + other
            except TypeError:
                return NotImplemented
            
        #other algo operations incoming gotta go slep now bb
            