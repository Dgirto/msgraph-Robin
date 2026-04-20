import msal
import logging
from datetime import datetime, timedelta

logger = logging.getLogger("msgraph_client")


class GraphAuth:
    """Maneja la autenticación con Microsoft Graph API"""

    GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0"
    SCOPE = ["https://graph.microsoft.com/.default"]

    def __init__(self, client_id: str, client_secret: str, tenant_id: str):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.access_token = None
        self.token_expires_at = None
        self._authenticate()

    def _authenticate(self):
        """Obtiene un token de acceso"""
        authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        app = msal.ConfidentialClientApplication(
            self.client_id,
            authority=authority,
            client_credential=self.client_secret
        )
        result = app.acquire_token_for_client(scopes=self.SCOPE)

        if "access_token" in result:
            self.access_token = result["access_token"]
            expires_in = result.get("expires_in", 3600)
            self.token_expires_at = datetime.now() + timedelta(seconds=expires_in - 300)
            logger.info("Autenticación exitosa con Microsoft Graph")
        else:
            error = result.get("error_description", result.get("error"))
            raise Exception(f"Error de autenticación: {error}")

    def ensure_valid_token(self):
        """Renueva el token si está por vencer"""
        if not self.token_expires_at or datetime.now() >= self.token_expires_at:
            logger.info("Renovando token...")
            self._authenticate()

    def get_headers(self) -> dict:
        """Devuelve los headers listos para usar en requests"""
        self.ensure_valid_token()
        return {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }