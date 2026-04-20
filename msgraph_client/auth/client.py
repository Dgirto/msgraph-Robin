import requests
import logging
from msgraph_client.auth import GraphAuth

logger = logging.getLogger("msgraph_client")


class GraphClient:
    """Hace las llamadas HTTP a Microsoft Graph API"""

    def __init__(self, auth: GraphAuth):
        self.auth = auth
        self.base_url = GraphAuth.GRAPH_ENDPOINT

    def get(self, endpoint: str) -> dict:
        """Hace un GET y devuelve el JSON"""
        url = f"{self.base_url}/{endpoint}"
        response = requests.get(url, headers=self.auth.get_headers())
        self._handle_error(response)
        return response.json()

    def get_all_pages(self, endpoint: str) -> list:
        """Hace GET manejando paginación automáticamente"""
        all_items = []
        url = f"{self.base_url}/{endpoint}"

        while url:
            response = requests.get(url, headers=self.auth.get_headers())
            self._handle_error(response)
            data = response.json()
            all_items.extend(data.get("value", []))
            url = data.get("@odata.nextLink")  # siguiente página o None

        return all_items

    def get_content(self, endpoint: str) -> bytes:
        """Descarga contenido binario (archivos)"""
        url = f"{self.base_url}/{endpoint}"
        headers = self.auth.get_headers()
        headers.pop("Content-Type", None)
        response = requests.get(url, headers=headers)
        self._handle_error(response)
        return response.content

    def post(self, endpoint: str, data: dict):
        """Hace un POST"""
        url = f"{self.base_url}/{endpoint}"
        response = requests.post(url, headers=self.auth.get_headers(), json=data)
        self._handle_error(response)
        if response.status_code in [202, 204]:
            return None
        return response.json()

    def _handle_error(self, response: requests.Response):
        """Lanza excepción si la respuesta es un error"""
        if response.status_code >= 400:
            raise Exception(
                f"Error {response.status_code} en Graph API: {response.text}"
            )