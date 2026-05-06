import logging
import requests
from io import BytesIO
from typing import Dict, List, Optional
import pandas as pd
from msgraph_client.auth.client import GraphClient

logger = logging.getLogger("msgraph_client")


class DriveWriter:
    """Cliente para escribir y actualizar archivos Excel en OneDrive via Graph API"""

    def __init__(self, client: GraphClient, user_email: str):
        """
        Args:
            client: instancia de GraphClient
            user_email: email del usuario dueño del drive
        """
        self.client = client
        self.user_email = user_email

    def upload_sheet(
        self,
        path: str,
        data: List[Dict],
        sheet_name: str = "Hoja1",
        index: bool = False,
    ) -> str:
        """
        Convierte una lista de dicts en Excel y lo sube a OneDrive.
        Si el archivo ya existe, lo reemplaza.

        Args:
            path: ruta destino en OneDrive (ej: "Reportes/ventas.xlsx")
            data: lista de dicts a convertir en Excel
            sheet_name: nombre de la hoja
            index: si True, incluye el índice del DataFrame

        Returns:
            str: ID del archivo subido

        Ejemplo:
            writer.upload_sheet(
                path="Reportes/ventas.xlsx",
                data=[{"Producto": "A", "Ventas": 100}],
                sheet_name="Mayo"
            )
        """
        df = pd.DataFrame(data)
        return self._upload_dataframe(path, df, sheet_name, index)

    def upload_dataframe(
        self,
        path: str,
        df: pd.DataFrame,
        sheet_name: str = "Hoja1",
        index: bool = False,
    ) -> str:
        """
        Sube un DataFrame como Excel a OneDrive.
        Si el archivo ya existe, lo reemplaza.

        Args:
            path: ruta destino en OneDrive (ej: "Reportes/ventas.xlsx")
            df: DataFrame a subir
            sheet_name: nombre de la hoja
            index: si True, incluye el índice del DataFrame

        Returns:
            str: ID del archivo subido
        """
        return self._upload_dataframe(path, df, sheet_name, index)

    def append_rows(
        self,
        path: str,
        data: List[Dict],
        sheet_name: Optional[str] = None,
        index: bool = False,
    ) -> str:
        """
        Descarga un Excel existente, agrega filas al final y lo vuelve a subir.

        Args:
            path: ruta del archivo en OneDrive
            data: lista de dicts con las filas nuevas
            sheet_name: hoja donde agregar. Si es None, usa la primera hoja
            index: si True, incluye el índice en el resultado

        Returns:
            str: ID del archivo actualizado

        Ejemplo:
            writer.append_rows(
                path="Reportes/ventas.xlsx",
                data=[{"Producto": "C", "Ventas": 300}],
                sheet_name="Mayo"
            )
        """
        # Descargar el Excel actual
        endpoint = f"users/{self.user_email}/drive/root:/{path}:/content"
        content = self.client.get_content(endpoint)
        existing_df = pd.read_excel(BytesIO(content), sheet_name=sheet_name)

        # Si sheet_name era None, pandas devuelve dict con todas las hojas
        if isinstance(existing_df, dict):
            sheet_key = list(existing_df.keys())[0]
            existing_df = existing_df[sheet_key]
            resolved_sheet = sheet_key
        else:
            resolved_sheet = sheet_name or "Hoja1"

        # Concatenar filas nuevas
        new_df = pd.DataFrame(data)
        combined_df = pd.concat([existing_df, new_df], ignore_index=True)

        return self._upload_dataframe(path, combined_df, resolved_sheet, index)

    def update_sheet(
        self,
        path: str,
        data: List[Dict],
        sheet_name: str,
        index: bool = False,
    ) -> str:
        """
        Reemplaza el contenido de una hoja específica manteniendo las demás intactas.

        Args:
            path: ruta del archivo en OneDrive
            data: lista de dicts con los nuevos datos
            sheet_name: nombre de la hoja a reemplazar
            index: si True, incluye el índice

        Returns:
            str: ID del archivo actualizado
        """
        # Descargar todas las hojas
        endpoint = f"users/{self.user_email}/drive/root:/{path}:/content"
        content = self.client.get_content(endpoint)
        all_sheets: dict = pd.read_excel(BytesIO(content), sheet_name=None)

        # Reemplazar solo la hoja indicada
        all_sheets[sheet_name] = pd.DataFrame(data)

        # Reconstruir el Excel con todas las hojas
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for sname, sdf in all_sheets.items():
                sdf.to_excel(writer, sheet_name=sname, index=index)
        excel_bytes = output.getvalue()

        return self._upload_bytes(path, excel_bytes)

    def create_folder(self, folder_path: str) -> str:
        """
        Crea una carpeta en OneDrive. Si ya existe, no hace nada.

        Args:
            folder_path: ruta de la nueva carpeta (ej: "Reportes/2025")

        Returns:
            str: ID de la carpeta

        Ejemplo:
            writer.create_folder("Reportes/2025")
        """
        parts = folder_path.rstrip("/").rsplit("/", 1)
        parent = parts[0] if len(parts) > 1 else ""
        folder_name = parts[-1]

        if parent:
            parent_ep = f"users/{self.user_email}/drive/root:/{parent}:/children"
        else:
            parent_ep = f"users/{self.user_email}/drive/root/children"

        payload = {
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "fail",
        }

        try:
            result = self.client.post(parent_ep, payload)
            logger.info(f"✓ Carpeta creada: {folder_path}")
            return result["id"]
        except Exception as e:
            if "409" in str(e) or "nameAlreadyExists" in str(e):
                logger.info(f"Carpeta ya existe: {folder_path}")
                if parent:
                    meta = self.client.get(f"users/{self.user_email}/drive/root:/{folder_path}")
                else:
                    meta = self.client.get(f"users/{self.user_email}/drive/root:/{folder_name}")
                return meta["id"]
            raise

    # ─── Métodos internos ───────────────────────────────────────────────────────

    def _upload_dataframe(
        self,
        path: str,
        df: pd.DataFrame,
        sheet_name: str,
        index: bool,
    ) -> str:
        """Convierte un DataFrame a bytes Excel y lo sube"""
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=index)
        return self._upload_bytes(path, output.getvalue())

    def _upload_bytes(self, path: str, file_bytes: bytes) -> str:
        """
        Sube bytes a OneDrive.
        Usa PUT simple para archivos <= 4MB, upload session para archivos mayores.
        """
        MAX_SIMPLE = 4 * 1024 * 1024  # 4 MB

        if len(file_bytes) <= MAX_SIMPLE:
            return self._simple_upload(path, file_bytes)
        else:
            return self._session_upload(path, file_bytes)

    def _simple_upload(self, path: str, file_bytes: bytes) -> str:
        """PUT directo para archivos <= 4 MB"""
        from msgraph_client.auth import GraphAuth
        url = f"{GraphAuth.GRAPH_ENDPOINT}/users/{self.user_email}/drive/root:/{path}:/content"

        self.client.auth._authenticate()
        headers = {
            "Authorization": f"Bearer {self.client.auth.token}",
            "Content-Type": "application/octet-stream",
        }
        response = requests.put(url, headers=headers, data=file_bytes)

        if response.status_code in [200, 201]:
            logger.info(f"✓ Archivo subido: {path}")
            return response.json()["id"]
        raise Exception(f"Error subiendo archivo: {response.status_code} - {response.text}")

    def _session_upload(self, path: str, file_bytes: bytes) -> str:
        """Upload session para archivos > 4 MB"""
        session_ep = f"users/{self.user_email}/drive/root:/{path}:/createUploadSession"
        session = self.client.post(session_ep, {
            "item": {"@microsoft.graph.conflictBehavior": "replace"}
        })
        upload_url = session["uploadUrl"]

        CHUNK = 5 * 1024 * 1024  # chunks de 5 MB
        total = len(file_bytes)
        offset = 0
        response = None

        while offset < total:
            chunk = file_bytes[offset: offset + CHUNK]
            end = offset + len(chunk) - 1
            headers = {
                "Content-Range": f"bytes {offset}-{end}/{total}",
                "Content-Length": str(len(chunk)),
            }
            response = requests.put(upload_url, headers=headers, data=chunk)
            if response.status_code not in [200, 201, 202]:
                raise Exception(f"Error en chunk upload: {response.status_code} - {response.text}")
            offset += len(chunk)

        logger.info(f"✓ Archivo grande subido: {path}")
        return response.json()["id"]