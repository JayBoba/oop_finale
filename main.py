import argparse
import sys
import os
from dotenv import load_dotenv
from typing import Dict, Set
from src.api.client import APIClient
from src.core.tables.table import Table
from src.core.tables.excel_writer import ExcelWriter
from src.api.models import ApiTable, ApiCell, CellType, CellReference

load_dotenv()

def main():
    parser = argparse.ArgumentParser(description="Export buildin.ai tables to Excel.")
    parser.add_argument("--table-id", help="The ID of the table to export.")
    parser.add_argument("--list", action="store_true", help="List all available tables.")
    parser.add_argument("--output", default="output.xlsx", help="Output Excel file name.")
    parser.add_argument("--test", action="store_true", help="Run in test mode with mock data.")

    args = parser.parse_args()

    token = os.getenv("BUILDIN_AI_TOKEN")
    if not token and not args.test:
        print("Error: BUILDIN_AI_TOKEN environment variable not set.")
        sys.exit(1)

        token = "dummy_token"

    client = APIClient(token=token)

    if args.test:
        print("Running in TEST mode with mock data.")
        mock_data = generate_mock_data()
        client.set_mock_mode(mock_data)

    if args.list:
        print("Fetching list of tables...")
        tables = client.get_tables()
        if tables:
            print(f"Found {len(tables)} tables:")
            for t in tables:
                if isinstance(t, dict):
                     t_id = t.get("id", "N/A")
                     t_name = t.get("name", "N/A")
                     print(f"- {t_name} (ID: {t_id})")
                else:
                    print(f"- {t}")
        else:
            print("No tables found or failed to fetch list.")
        return

    if not args.table_id:
        print("Error: --table-id is required unless --list is used.")
        parser.print_help()
        sys.exit(1)

    if args.test and args.table_id not in mock_data:
         print(f"Mock data does not contain {args.table_id}, but continuing (might fail). Available: {list(mock_data.keys())}")

    print(f"Fetching table {args.table_id} and dependencies...")

    tables_to_process = [args.table_id]
    processed_table_ids: Set[str] = set()
    fetched_tables: Dict[str, Table] = {}

    while tables_to_process:
        current_id = tables_to_process.pop(0)
        if current_id in processed_table_ids:
            continue

        print(f"Fetching {current_id}...")
        api_table = client.get_table(current_id)

        if not api_table:
            print(f"Failed to fetch table {current_id}. Skipping.")
            continue

        table = Table(api_table)
        fetched_tables[current_id] = table
        processed_table_ids.add(current_id)

        linked_ids = table.get_linked_table_ids()
        for link_id in linked_ids:
            if link_id not in processed_table_ids and link_id not in tables_to_process:
                tables_to_process.append(link_id)

    if not fetched_tables:
        print("No tables fetched. Exiting.")
        return

    print(f"Fetched {len(fetched_tables)} tables. Writing to {args.output}...")
    writer = ExcelWriter(args.output)
    writer.write_tables(list(fetched_tables.values()))
    print("Done.")

def generate_mock_data() -> Dict[str, ApiTable]:
    t1 = ApiTable(
        id="table_1",
        name="Budget",
        cells=[
            ApiCell(id="c1", row=1, column=1, cell_type=CellType.VALUE, value="Item"),
            ApiCell(id="c2", row=1, column=2, cell_type=CellType.VALUE, value="Cost"),
            ApiCell(id="c3", row=2, column=1, cell_type=CellType.VALUE, value="Rent"),
            ApiCell(id="c4", row=2, column=2, cell_type=CellType.VALUE, value=1000),
            ApiCell(id="c5", row=3, column=1, cell_type=CellType.VALUE, value="Food"),
            ApiCell(id="c6", row=3, column=2, cell_type=CellType.VALUE, value=500),
            ApiCell(id="c7", row=4, column=1, cell_type=CellType.VALUE, value="Total"),
            ApiCell(id="c8", row=4, column=2, cell_type=CellType.FORMULA, formula="=SUM(B2:B3)"),
            ApiCell(id="c9", row=5, column=1, cell_type=CellType.VALUE, value="Link to Info"),
            ApiCell(id="c10", row=5, column=2, cell_type=CellType.LINK,
                    reference=[CellReference(table_id="table_2", cell_address="A1")])
        ]
    )

    t2 = ApiTable(
        id="table_2",
        name="Info",
        cells=[
            ApiCell(id="c2_1", row=1, column=1, cell_type=CellType.VALUE, value="Important Info"),
            ApiCell(id="c2_2", row=2, column=1, cell_type=CellType.VALUE, value="Details..."),
            ApiCell(id="c2_3", row=3, column=1, cell_type=CellType.LINK,
                    reference=[CellReference(table_id="table_1", cell_address="A1")])

        ]
    )

    return {"table_1": t1, "table_2": t2}

if __name__ == "__main__":
    main()