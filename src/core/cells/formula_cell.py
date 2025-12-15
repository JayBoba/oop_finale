import re
import ast
import math
import operator
from typing import Any, Dict, List, Set, Optional, Union
from decimal import Decimal, InvalidOperation
from collections import defaultdict

from .base_cell import BaseCell, ExcelAddress, CellType, FormatType, CellEvalError
from ..tables.table import Table
from src.utils.validators import validate_formula_syntax #TBD
import logging

logger = logging.getLogger(__name__)


class FormulaCell(BaseCell):
    SUPPORTED_FUNCTIONS = {
        'SUM': sum,
        'AVERAGE': lambda args: sum(args) / len(args) if args else 0,
        'AVG': lambda args: sum(args) / len(args) if args else 0,
        'MIN': min,
        'MAX': max,
        'ABS': abs,
        'ROUND': round,
        'SQRT': math.sqrt,
        'POWER': math.pow,
        'POW': math.pow,
        'EXP': math.exp,
        'LN': math.log,
        'LOG': math.log10,
        'LOG10': math.log10,
        'SIN': math.sin,
        'COS': math.cos,
        'TAN': math.tan,
        'ASIN': math.asin,
        'ACOS': math.acos,
        'ATAN': math.atan,
        'PI': lambda: math.pi,
        'E': lambda: math.e,
        'INT': int,
        'MOD': lambda args: args[0] % args[1] if len(args) == 2 else None,
        'COUNT': len,
        'COUNTA': len,
        'IF': lambda args: args[1] if args[0] else args[2] if len(args) == 3 else None,
    }
    
    SUPPORTED_OPERATORS = {
        ast.Add: operator.add,
        ast.Sub: operator.sub,
        ast.Mult: operator.mul,
        ast.Div: operator.truediv,
        ast.FloorDiv: operator.floordiv,
        ast.Pow: operator.pow,
        ast.Mod: operator.mod,
        ast.USub: operator.neg,
        ast.UAdd: operator.pos,
        ast.Eq: operator.eq,
        ast.NotEq: operator.ne,
        ast.Lt: operator.lt,
        ast.LtE: operator.le,
        ast.Gt: operator.gt,
        ast.GtE: operator.ge,
    }
    
    def __init__(
        self,
        address: ExcelAddress,
        formula: str,
        table: Table,
        format_type: Optional[FormatType] = None
    ):
        super().__init__(address, formula, CellType.FORMULA, format_type)
        
        self._table = table
        self._parsed_formula = self._parse_formula(formula)
        self._dependencies = self._extract_dependencies(formula)
        self._external_refs = self._extract_external_references(formula)
        self._range_dependencies = self._extract_range_dependencies(formula)
        
        try:
            validate_formula_syntax(formula)
        except ValueError as e:
            logger.warning(f"Syntax error in {address}: {formula}. Error: {e}")
            self._syntax_error = str(e)
        else:
            self._syntax_error = None
        
        logger.debug(f"Created FormulaCell {address} with: {formula}")
    
    def _parse_formula(self, formula: str) -> str:
        formula = formula.strip()
        if formula.startswith('='):
            return formula[1:].strip()
        return formula
    
    def _extract_dependencies(self, formula: str) -> Set[ExcelAddress]:
        dependencies = set()
        
        formula_without_strings = re.sub(r'"[^"]*"|\'[^\']*\'', '""', formula)
        
        patterns = [
            r'\b([A-Z]{1,3}\$?[1-9][0-9]{0,6})\b',
            r"'[^']*'!([A-Z]{1,3}\$?[1-9][0-9]{0,6})",
            r"(\b[A-Za-z_][A-Za-z0-9_]*)!([A-Z]{1,3}\$?[1-9][0-9]{0,6})",
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, formula_without_strings, re.IGNORECASE)
            for match in matches:
                if isinstance(match, tuple):
                    cell_ref = match[-1]
                else:
                    cell_ref = match
                
                cell_ref = cell_ref.replace('$', '')
                
                try:
                    address = ExcelAddress.from_string(cell_ref.upper())
                    dependencies.add(address)
                except ValueError as e:
                    logger.debug(f"Failed to parse an adress {cell_ref}: {e}")
                    continue
        
        return dependencies
    
    def _extract_range_dependencies(self, formula: str) -> Dict[str, List[ExcelAddress]]:
        range_deps = {}
        
        range_pattern = r'([A-Z]{1,3}\$?[1-9][0-9]{0,6})\s*:\s*([A-Z]{1,3}\$?[1-9][0-9]{0,6})'
        
        matches = re.findall(range_pattern, formula, re.IGNORECASE)
        for start_str, end_str in matches:
            start_str = start_str.replace('$', '')
            end_str = end_str.replace('$', '')
            
            try:
                start_addr = ExcelAddress.from_string(start_str.upper())
                end_addr = ExcelAddress.from_string(end_str.upper())
                
                addresses = []
                for row in range(min(start_addr.row, end_addr.row), 
                               max(start_addr.row, end_addr.row) + 1):
                    for col in range(min(start_addr.column, end_addr.column), 
                                   max(start_addr.column, end_addr.column) + 1):
                        addresses.append(ExcelAddress(row, col))
                
                range_key = f"{start_str}:{end_str}"
                range_deps[range_key] = addresses
                
                self._dependencies.update(addresses)
                
            except ValueError as e:
                logger.debug(f"Failed to parse diap: {start_str}:{end_str}: {e}")
                continue
        
        return range_deps
    
    def _extract_external_references(self, formula: str) -> List[str]:
        external_refs = []
        
        patterns = [
            r'\b([A-Za-z_][A-Za-z0-9_]*)!([A-Z]{1,3}[1-9][0-9]{0,6})\b',
            r"'([^']+)'!([A-Z]{1,3}[1-9][0-9]{0,6})\b",
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, formula)
            for table_name, cell_ref in matches:
                ref = f"{table_name}!{cell_ref}"
                external_refs.append(ref)
        
        return external_refs
    
    def _do_evaluate(self, context: Dict[str, Any]) -> Any:
        if self._syntax_error:
            raise CellEvalError(f"Syntax error: {self._syntax_error}")
        
        stack_key = f"{self._table.id}:{self.address}"
        
        if "evaluation_stack" not in context:
            context["evaluation_stack"] = set()
        
        if stack_key in context["evaluation_stack"]:
            raise CellEvalError(f"Found loop dependency: {stack_key}")
        
        context["evaluation_stack"].add(stack_key)
        
        try:
            evaluated_formula = self._substitute_cell_references(self._parsed_formula, context)
            result = self._safe_evaluate_expression(evaluated_formula, context)
            if self._format_type == FormatType.PERCENTAGE and isinstance(result, (int, float)):
                result = Decimal(str(result))
            
            return result
            
        except CellEvalError:
            raise
        except Exception as e:
            raise CellEvalError(f"Error while calculation formula: {str(e)}")
        finally:
            context["evaluation_stack"].remove(stack_key)
    
    def _substitute_cell_references(self, formula: str, context: Dict[str, Any]) -> str:
        for range_key, addresses in self._range_dependencies.items():
            if f"SUM({range_key})" in formula.upper():
                values = []
                for addr in addresses:
                    cell = self._table.get_cell(addr)
                    if cell:
                        try:
                            values.append(cell.evaluate(context))
                        except:
                            values.append(0)
                sum_value = sum(v for v in values if isinstance(v, (int, float, Decimal)))
                formula = formula.replace(f"SUM({range_key})", str(sum_value), 1)
        
        for dependency in self._dependencies:
            cell = self._table.get_cell(dependency)
            if cell:
                try:
                    cell_value = cell.evaluate(context)
                    ref_str = str(dependency)
                    for pattern in [ref_str, f"${ref_str[0]}${ref_str[1:]}", 
                                   f"${ref_str[0]}{ref_str[1:]}", f"{ref_str[0]}${ref_str[1:]}"]:
                        if pattern in formula:
                            formula = formula.replace(pattern, str(cell_value))
                except Exception as e:
                    logger.debug(f"Failed to calculate dependency {dependency}: {e}")
        
        return formula
    
    def _safe_evaluate_expression(self, expression: str, context: Dict[str, Any]) -> Any:
        expression = expression.strip()
        
        if not expression:
            return ""
        
        try:
            return int(expression)
        except ValueError:
            pass
        
        try:
            return float(expression)
        except ValueError:
            pass
        
        try:
            return self._evaluate_with_ast(expression, context)
        except Exception as e:
            logger.debug(f"Could not calculate '{expression}' as a formula: {e}")
            return expression
    
    def _evaluate_with_ast(self, expression: str, context: Dict[str, Any]) -> Any:
        safe_env = {
            '__builtins__': {},
            'math': math,
        }
        
        for func_name, func in self.SUPPORTED_FUNCTIONS.items():
            safe_env[func_name] = func
            safe_env[func_name.upper()] = func
            safe_env[func_name.lower()] = func
        
        try:
            tree = ast.parse(expression, mode='eval')
            result = self._evaluate_ast(tree.body, safe_env)
            return result
        except Exception as e:
            raise CellEvalError(f"Failed to calculate AST: {str(e)}")
    
    def _evaluate_ast(self, node: ast.AST, env: Dict[str, Any]) -> Any:
        if isinstance(node, ast.Constant):
            return node.value
        
        elif isinstance(node, ast.Constant):
            return node.value
        
        elif isinstance(node, ast.Name):
            if node.id in env:
                if callable(env[node.id]):
                    return env[node.id]
                else:
                    return env[node.id]()
            else:
                raise CellEvalError(f"Unknown id: {node.id}")
        
        elif isinstance(node, ast.BinOp):
            left = self._evaluate_ast(node.left, env)
            right = self._evaluate_ast(node.right, env)
            
            op_type = type(node.op)
            if op_type in self.SUPPORTED_OPERATORS:
                try:
                    return self.SUPPORTED_OPERATORS[op_type](left, right)
                except Exception as e:
                    raise CellEvalError(f"Operation error {op_type}: {e}")
            else:
                raise CellEvalError(f"Unsupported operation: {op_type}")
        
        elif isinstance(node, ast.UnaryOp):
            operand = self._evaluate_ast(node.operand, env)
            op_type = type(node.op)
            
            if op_type in self.SUPPORTED_OPERATORS:
                return self.SUPPORTED_OPERATORS[op_type](operand)
            else:
                raise CellEvalError(f"Unsupported operation: {op_type}")
        
        elif isinstance(node, ast.Call):
            if isinstance(node.func, ast.Name):
                func_name = node.func.id
                if func_name in env:
                    func = env[func_name]
                else:
                    raise CellEvalError(f"Unknown function: {func_name}")
            else:
                raise CellEvalError("Nested func calls are not yet supported")
            
            args = []
            for arg in node.args:
                args.append(self._evaluate_ast(arg, env))
            
            try:
                return func(*args)
            except Exception as e:
                raise CellEvalError(f"Failed to call a function {func_name}: {e}")
        
        elif isinstance(node, ast.Compare):
            left = self._evaluate_ast(node.left, env)
            
            for op, comparator in zip(node.ops, node.comparators):
                right = self._evaluate_ast(comparator, env)
                op_type = type(op)
                
                if op_type in self.SUPPORTED_OPERATORS:
                    if not self.SUPPORTED_OPERATORS[op_type](left, right):
                        return False
                else:
                    raise CellEvalError(f"Unsupported operation: {op_type}")
                
                left = right
            
            return True
        
        else:
            raise CellEvalError(f"Unsupported AST node: {type(node)}")
    
    def get_excel_formula(self) -> str:
        return f"={self._parsed_formula}"
    
    @property
    def dependencies(self) -> Set[ExcelAddress]:
        return self._dependencies.copy()
    
    @property
    def external_references(self) -> List[str]:
        return self._external_refs.copy()
    
    @property
    def range_dependencies(self) -> Dict[str, List[ExcelAddress]]:
        return self._range_dependencies.copy()
    
    @property
    def has_syntax_error(self) -> bool:
        return self._syntax_error is not None
    
    @property
    def syntax_error(self) -> Optional[str]:
        return self._syntax_error
    
    def __str__(self) -> str:
        base_str = super().__str__()
        if self.has_syntax_error:
            return f"{base_str} [SYNTAX ERROR: {self._syntax_error}]"
        return base_str
    
    def __repr__(self) -> str:
        deps_count = len(self._dependencies)
        ext_refs_count = len(self._external_refs)
        return (f"FormulaCell(address={self.address}, "
                f"formula={self._raw_value[:50]}{'...' if len(self._raw_value) > 50 else ''}, "
                f"deps={deps_count}, ext_refs={ext_refs_count})")

def extract_all_references(formula: str) -> Dict[str, List[str]]:
    result = {
        'simple': [],
        'absolute': [],
        'mixed': [],
        'ranges': [],
        'external': [],
        'sheet_refs': [],
    }
    
    formula_no_strings = re.sub(r'"[^"]*"|\'[^\']*\'', '""', formula)
    
    cell_pattern = r'\b([$]?[A-Z]{1,3}[$]?[1-9][0-9]{0,6})\b'
    for match in re.finditer(cell_pattern, formula_no_strings, re.IGNORECASE):
        ref = match.group(1)
        if ref.startswith('$') and ref.endswith('$'):
            result['absolute'].append(ref)
        elif '$' in ref:
            result['mixed'].append(ref)
        else:
            result['simple'].append(ref)
    
    range_pattern = r'([A-Z]{1,3}[1-9][0-9]{0,6})\s*:\s*([A-Z]{1,3}[1-9][0-9]{0,6})'
    for match in re.finditer(range_pattern, formula_no_strings, re.IGNORECASE):
        result['ranges'].append(match.group(0))
    
    ext_pattern = r'\b([A-Za-z_][A-Za-z0-9_]*)!([A-Z]{1,3}[1-9][0-9]{0,6})\b'
    for match in re.finditer(ext_pattern, formula_no_strings):
        result['external'].append(match.group(0))
    
    sheet_pattern = r"'([^']+)'!([A-Z]{1,3}[1-9][0-9]{0,6})\b"
    for match in re.finditer(sheet_pattern, formula_no_strings):
        result['sheet_refs'].append(match.group(0))
    
    return result


def validate_formula_references(formula: str, available_cells: Set[str]) -> Dict[str, Any]:
    refs = extract_all_references(formula)
    
    valid_simple = []
    invalid_simple = []
    
    for ref in refs['simple']:
        if ref.upper() in available_cells:
            valid_simple.append(ref)
        else:
            invalid_simple.append(ref)

    valid_ranges = []
    invalid_ranges = []
    
    for range_str in refs['ranges']:
        start_end = range_str.split(':')
        if len(start_end) == 2:
            start, end = start_end
            if start.upper() in available_cells and end.upper() in available_cells:
                valid_ranges.append(range_str)
            else:
                invalid_ranges.append(range_str)
    
    return {
        'total_references': sum(len(v) for v in refs.values()),
        'valid_simple': valid_simple,
        'invalid_simple': invalid_simple,
        'valid_ranges': valid_ranges,
        'invalid_ranges': invalid_ranges,
        'external_references': refs['external'],
        'sheet_references': refs['sheet_refs'],
        'has_invalid_references': len(invalid_simple) > 0 or len(invalid_ranges) > 0,
    }