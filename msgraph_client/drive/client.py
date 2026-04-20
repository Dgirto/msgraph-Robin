import logging
from io import BytesIO
from typing import List, Dict, Any, Optional
import pandas as pd
from msgraph_client.auth.client import GraphClient
from msgraph_client.models import FileObject

logger = logging.getLogger("msgraph_client")


class DriveClient:
    """Cliente para leer archivos desde OneDrive / SharePoint via Graph API"""

    def __init__(self, client: GraphClient, user_email: str):
        """
        Args:
            client: instancia de GraphClient
            user_email: email del usuario dueño del drive
        """
        self.client = client
        self.user_email = user_email

    def get_file(self, file_id: str, drive_id: Optional[str] = None) -> FileObject:
        """
        Obtiene metadata de un archivo por su ID.

        Args:
            file_id: ID del archivo en OneDrive
            drive_id: ID del drive (opcional, usa el drive del usuario si no se da)
        """
        if drive_id:
            endpoint = f"drives/{drive_id}/items/{file_id}"
        else:
            endpoint = f"users/{self.user_email}/drive/items/{file_id}"

        raw = self.client.get(endpoint)
        return self._parse_file(raw)

    def get_file_by_path(self, path: str) -> FileObject:
        """
        Obtiene metadata de un archivo por su ruta.

        Args:
            path: ruta del archivo (ej: "Reportes/ventas.xlsx")
        """
        endpoint = f"users/{self.user_email}/drive/root:/{path}"
        raw = self.client.get(endpoint)
        return self._parse_file(raw)

    def read_sheet(
        self,
        sheet_name: str,
        file_id: Optional[str] = None,
        path: Optional[str] = None,
        has_header: bool = True,
        drive_id: Optional[str] = None
    ) -> List[Dict]:
        """
        Lee una hoja completa de un archivo Excel.

        Args:
            sheet_name: nombre de la hoja
            file_id: ID del archivo (usar esto o path)
            path: ruta del archivo (usar esto o file_id)
            has_header: si True usa la primera fila como nombres de columnas
            drive_id: ID del drive (opcional)
        """
        df = self._load_excel(file_id=file_id, path=path, sheet_name=sheet_name, drive_id=drive_id)

        if not has_header:
            # Renombrar columnas a letras: A, B, C...
            df.columns = [self._col_index_to_letter(i) for i in range(len(df.columns))]

        return df.to_dict(orient="records")

    def read_column(
        self,
        sheet_name: str,
        column: str,
        file_id: Optional[str] = None,
        path: Optional[str] = None,
        has_header: bool = True,
        drive_id: Optional[str] = None
    ) -> List[Any]:
        """
        Lee una columna específica de una hoja Excel.

        Args:
            sheet_name: nombre de la hoja
            column: nombre de la columna (si has_header=True) o letra (A, B, C...)
            file_id: ID del archivo
            path: ruta del archivo
            has_header: si True busca por nombre, si False busca por letra
            drive_id: ID del drive (opcional)
        """
        df = self._load_excel(file_id=file_id, path=path, sheet_name=sheet_name, drive_id=drive_id)

        if not has_header:
            df.columns = [self._col_index_to_letter(i) for i in range(len(df.columns))]

        if column not in df.columns:
            raise ValueError(f"Columna '{column}' no encontrada. Columnas disponibles: {list(df.columns)}")

        return df[column].tolist()

    def read_rows(
        self,
        sheet_name: str,
        limit: int,
        file_id: Optional[str] = None,
        path: Optional[str] = None,
        has_header: bool = True,
        drive_id: Optional[str] = None
    ) -> List[Dict]:
        """
        Lee las primeras N filas de una hoja Excel.

        Args:
            sheet_name: nombre de la hoja
            limit: cantidad de filas a leer
            file_id: ID del archivo
            path: ruta del archivo
            has_header: si True usa la primera fila como nombres de columnas
            drive_id: ID del drive (opcional)
        """
        df = self._load_excel(file_id=file_id, path=path, sheet_name=sheet_name, drive_id=drive_id)

        if not has_header:
            df.columns = [self._col_index_to_letter(i) for i in range(len(df.columns))]

        return df.head(limit).to_dict(orient="records")

    # ─── Métodos internos ───────────────────────────────────────────────────────

    def _load_excel(
        self,
        sheet_name: str,
        file_id: Optional[str] = None,
        path: Optional[str] = None,
        drive_id: Optional[str] = None
    ) -> pd.DataFrame:
        """Descarga el Excel y lo carga en un DataFrame"""
        if file_id:
            if drive_id:
                endpoint = f"drives/{drive_id}/items/{file_id}/content"
            else:
                endpoint = f"users/{self.user_email}/drive/items/{file_id}/content"
        elif path:
            endpoint = f"users/{self.user_email}/drive/root:/{path}:/content"
        else:
            raise ValueError("Debes proporcionar file_id o path")

        content = self.client.get_content(endpoint)
        return pd.read_excel(BytesIO(content), sheet_name=sheet_name)

    def _parse_file(self, raw: dict) -> FileObject:
        """Convierte el JSON crudo de Microsoft en un objeto FileObject"""
        return FileObject(
            id=raw.get("id", ""),
            name=raw.get("name", ""),
            path=raw.get("parentReference", {}).get("path", ""),
            size=raw.get("size", 0),
            mime_type=raw.get("file", {}).get("mimeType"),
            drive_id=raw.get("parentReference", {}).get("driveId")
        )

    def _col_index_to_letter(self, index: int) -> str:
        """Convierte índice numérico a letra de columna (0=A, 1=B, ...)"""
        letters = ""
        while index >= 0:
            letters = chr(index % 26 + ord("A")) + letters
            index = index // 26 - 1
        return letters