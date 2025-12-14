from abc import ABC, abstractmethod
from typing import Any, Dict, Optional, Set, List
from dataclasses import dataclass, field
import re
import logging
from enum import Enum

from src.api.models import CellType, FormatType, CellReference
from src.utils.validators import validate_cell_address #TBD

logger = logging.getLogger(__name__)

class CellEvalError(Exception):
    pass

class CellState(Enum):
    PENDING = "pending"
    CALCULATING = "calculating"
    COMPLETED = "completed"
    ERROR = "error"

@dataclass(frozen=True, eq=True, order=True)
class ExcelAddress:
    row: int
    column: int

    def __post_init__(self):
        if self.row < 1:
            raise ValueError(f"Row must be positive, got {self.row}")
        if self.row > 1048576:
            raise ValueError(f"Row exceeds Excel limit, got {self.row}")
        if self.column < 1:
            raise ValueError(f"Column must be positive, got {self.column}")
        if self.column > 16384:
            raise ValueError(f"Column exceeds Excel limit, got {self.column}")
        
    def __str__(self) -> str:
        return self.to_excel_format()
    
    def to_excel_format(self) -> str:
        col = self.column
        result = ""
        while col > 0:
            col, remainder = divmod(col - 1, 26)
            result = chr(65 + remainder) + result
        return f"{result}{self.row}"
    
    def __equal__(self, other: object) -> bool:
        if not isinstance(other, ExcelAddress):
            return NotImplemented
        return self.row == other.row and self.column == other.column
    
    def __notequal__(self, other: object) -> bool:
        return not self.__eq__(other)
    
    def __lt__(self, other: object) -> bool:
        if not isinstance(other, ExcelAddress):
            return NotImplemented
        if self.row != other.row:
            return self.row < other.row
        return self.column < other.column
    
    def __hash__(self) -> int:
        return hash((self.row, self.column))
    
    def __add__(self, other: object) -> "ExcelAddress":
        if isinstance(other, tuple) and len(other) == 2:
            row_offset, col_offset = other
            return ExcelAddress(self.row + row_offset, self.column + col_offset)
        return NotImplemented
    
    def __sub__(self, other: object) -> "ExcelAddress":
        if isinstance(other, tuple) and len(other) == 2:
            row_offset, col_offset = other
            return ExcelAddress(self.row - row_offset, self.column - col_offset)
        return NotImplemented
    
    @classmethod
    def from_string(cls, address: str) -> "ExcelAddress":
        if '!' in address:
            address = address.split('!')[1]
        
        match = re.match(r"^([A-Z]+)(\d+)$", address.upper())
        if not match:
            raise ValueError(f"Invalid Excel address format: {address}")
        
        col_str, row_str = match.groups()
        
        col = 0
        for char in col_str:
            col = col * 26 + (ord(char) - 64)
        
        return cls(row=int(row_str), col=col)
    
class BaseCell(ABC):
    
    def __init__(
        self,
        address: ExcelAddress,
        raw_value: Any,
        cell_type: CellType,
        format_type: Optional[FormatType] = None
    ):
        self._address = address
        self._raw_value = raw_value
        self._cell_type = cell_type
        self._format_type = format_type
        self._evaluation_state: CellState = CellState.PENDING
        self._cached_value: Optional[Any] = None
        self._error_message: Optional[str] = None
        self._dependencies: Set[ExcelAddress] = set()
        self._external_refs: List[CellReference] = []
        
    @property
    def address(self) -> ExcelAddress:
        return self._address
    
    @property
    def raw_value(self) -> Any:
        return self._raw_value
    
    @property
    def cell_type(self) -> CellType:
        return self._cell_type
    
    @property
    def format_type(self) -> Optional[FormatType]:
        return self._format_type
    
    @property
    def evaluation_state(self) -> CellState:
        return self._evaluation_state
    
    @property
    def cached_value(self) -> Optional[Any]:
        return self._cached_value
    
    @property
    def error_message(self) -> Optional[str]:
        return self._error_message
    
    @property
    def dependencies(self) -> Set[ExcelAddress]:
        return self._dependencies.copy()
    
    @property
    def external_references(self) -> List[CellReference]:
        return self._external_refs.copy()
    
    def evaluate(self, context: Optional[Dict[str, Any]] = None) -> Any:
        if self._evaluation_state == CellState.CALCULATING:
            raise CellEvalError(
                f"Circular dependency detected at {self.address}"
            )
        
        if self._evaluation_state == CellState.COMPLETED and not context:
            return self._cached_value
        
        try:
            self._evaluation_state = CellState.CALCULATING
            self._cached_value = self._do_evaluate(context or {})
            self._evaluation_state = CellState.COMPLETED
            self._error_message = None
            return self._cached_value
            
        except Exception as e:
            self._evaluation_state = CellState.ERROR
            self._error_message = str(e)
            logger.error(f"Error evaluating cell {self.address}: {e}")
            raise CellEvalError(f"Error in cell {self.address}: {e}")
    
    @abstractmethod
    def _do_evaluate(self, context: Dict[str, Any]) -> Any:
        pass
    
    @abstractmethod
    def get_excel_formula(self) -> Optional[str]:
        pass
    
    def get_excel_value(self) -> Any:
        if self._evaluation_state == CellState.COMPLETED:
            return self._cached_value
        return self._raw_value
    
    def reset_evaluation(self):
        self._evaluation_state = CellState.PENDING
        self._cached_value = None
        self._error_message = None
    
    def add_dependency(self, address: ExcelAddress):
        self._dependencies.add(address)
    
    def add_external_reference(self, ref: CellReference):
        self._external_refs.append(ref)
    
    def __str__(self) -> str:
        return f"{self.address}: {self.raw_value}"
    
    def __repr__(self) -> str:
        return (f"{self.__class__.__name__}("
                f"address={self.address}, "
                f"type={self.cell_type}, "
                f"value={self.raw_value})")
    
    def __eq__(self, other: object) -> bool:
        if not isinstance(other, BaseCell):
            return NotImplemented
        return (self.address == other.address and 
                self.raw_value == other.raw_value and
                self.cell_type == other.cell_type)
    
    def __hash__(self) -> int:
        return hash((self.address, self.raw_value, self.cell_type))