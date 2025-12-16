import requests
from typing import Optional, List, Dict
from src.api.models import ApiTable, ApiResponse, ApiCell, CellType
from datetime import datetime
import json

class APIClient:
    def __init__(self, token: str, base_url: str = "https://api.buildin.ai"):
        self.token = token
        self.base_url = base_url.rstrip("/")
        self.headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        self._mock_data = {}
        self._is_mock = False

    def set_mock_mode(self, mock_data: Dict[str, ApiTable]):
        self._is_mock = True
        self._mock_data = mock_data

    def get_tables(self) -> List[Dict[str, str]]:
        if self._is_mock:
            return [{"id": t.id, "name": t.name} for t in self._mock_data.values()]

        #GET /tables
        url = f"{self.base_url}/tables"
        try:
            response = requests.get(url, headers=self.headers, timeout=10)
            if response.status_code == 200:
                data = response.json()
                if "tables" in data:
                    return data["tables"]
                return data
            elif response.status_code == 404:
                # if 404, maybe i should try /v1/tables or something??? idk for now just raising an error
                print(f"Warning: {url} returned 404. Using empty list.")
                return []
            else:
                response.raise_for_status()
        except Exception as e:
            print(f"Error fetching tables: {e}")
            return []

    def get_table(self, table_id: str) -> Optional[ApiTable]:
        if self._is_mock:
            return self._mock_data.get(table_id)

        #GET /tables/{table_id}
        url = f"{self.base_url}/tables/{table_id}"
        try:
            response = requests.get(url, headers=self.headers, timeout=10)
            if response.status_code == 200:
                data = response.json()
                return ApiTable(**data)
            elif response.status_code == 404:
                print(f"Table {table_id} not found.")
                return None
            else:
                response.raise_for_status()
        except Exception as e:
            print(f"Error fetching table {table_id}: {e}")
            return None