import time
import requests
import logging
from msgraph_client.auth import GraphAuth

logger = logging.getLogger("msgraph_client")


class GraphClient:
    """Hace las llamadas HTTP a Microsoft Graph API"""

    MAX_RETRIES = 3  # intentos máximos ante un error

    def __init__(self, auth: GraphAuth):
        self.auth = auth
        self.base_url = GraphAuth.GRAPH_ENDPOINT

    def get(self, endpoint: str) -> dict:
        """Hace un GET y devuelve el JSON"""
        url = f"{self.base_url}/{endpoint}"
        response = self._request_with_retry("GET", url)
        return response.json()

    def get_all_pages(self, endpoint: str) -> list:
        """Hace GET manejando paginación automáticamente"""
        all_items = []
        url = f"{self.base_url}/{endpoint}"

        while url:
            response = self._request_with_retry("GET", url)
            data = response.json()
            all_items.extend(data.get("value", []))
            url = data.get("@odata.nextLink")

        return all_items

    def get_content(self, endpoint: str) -> bytes:
        """Descarga contenido binario (archivos)"""
        url = f"{self.base_url}/{endpoint}"
        headers = self.auth.get_headers()
        headers.pop("Content-Type", None)
        response = self._request_with_retry("GET", url, headers=headers)
        return response.content

    def post(self, endpoint: str, data: dict):
        """Hace un POST"""
        url = f"{self.base_url}/{endpoint}"
        response = self._request_with_retry("POST", url, json=data)
        if response.status_code in [202, 204]:
            return None
        return response.json()

    def _request_with_retry(self, method: str, url: str, headers: dict = None, **kwargs) -> requests.Response:
        """
        Ejecuta un request con reintentos automáticos.
        - Si recibe 429 (rate limit): espera el tiempo que indica Microsoft y reintenta
        - Si recibe 401 (token expirado): renueva el token y reintenta
        - Si falla por otro error: reintenta hasta MAX_RETRIES veces
        """
        if headers is None:
            headers = self.auth.get_headers()

        for attempt in range(1, self.MAX_RETRIES + 1):
            response = requests.request(method, url, headers=headers, **kwargs)

            # Rate limit — Microsoft dice "espera antes de seguir"
            if response.status_code == 429:
                wait = int(response.headers.get("Retry-After", 10))
                logger.warning(f"Rate limit alcanzado. Esperando {wait} segundos... (intento {attempt}/{self.MAX_RETRIES})")
                time.sleep(wait)
                continue

            # Token expirado — renovar y reintentar
            if response.status_code == 401:
                logger.warning("Token inválido, renovando...")
                self.auth._authenticate()
                headers = self.auth.get_headers()
                continue

            # Error del servidor — esperar y reintentar
            if response.status_code >= 500:
                wait = 2 ** attempt  # 2, 4, 8 segundos
                logger.warning(f"Error del servidor {response.status_code}. Reintentando en {wait}s...")
                time.sleep(wait)
                continue

            # Cualquier otro error — lanzar excepción
            if response.status_code >= 400:
                raise Exception(f"Error {response.status_code} en Graph API: {response.text}")

            return response

        raise Exception(f"Se agotaron los {self.MAX_RETRIES} intentos para {url}")