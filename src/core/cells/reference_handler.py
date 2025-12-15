from typing import List
from src.core.cells.base_cell import BaseCell
from src.api.models import ApiCell

class ReferenceHandler(BaseCell):
    def __init__(self, api_cell: ApiCell):
        super().__init__(api_cell)
        self.references = api_cell.reference

    def get_value(self):
        #uhhh ???? gonna think abt dis later
        if not self.references:
            return "Reference Error"

        #taking the first reference for simplicity idgaf
        ref = self.references[0]
        #gon need to resolve table_id to sheet name later idk
        return f"LINK:{ref.table_id}!{ref.cell_address}"

    def get_referenced_table_ids(self) -> List[str]:
        return [ref.table_id for ref in self.references]