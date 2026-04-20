import logging
import time
from typing import List, Callable, Optional
from datetime import datetime
from msgraph_client.auth.client import GraphClient
from msgraph_client.models import Email, Attachment

logger = logging.getLogger("msgraph_client")


class EmailClient:
    """Cliente para leer correos desde Microsoft Outlook via Graph API"""

    def __init__(self, client: GraphClient, mailbox: str):
        """
        Args:
            client: instancia de GraphClient
            mailbox: email de la bandeja a leer (ej: "usuario@empresa.com")
        """
        self.client = client
        self.mailbox = mailbox

    def get_new_emails(
        self,
        only_unread: bool = True,
        from_address: Optional[str] = None,
        subject_contains: Optional[str] = None,
        include_attachments: bool = False,
        limit: int = 10
    ) -> List[Email]:
        """
        Obtiene correos de la bandeja de entrada.

        Args:
            only_unread: si True, solo trae los no leídos
            from_address: filtra por remitente
            subject_contains: filtra por palabra en el asunto
            include_attachments: si True, descarga los adjuntos
            limit: cantidad máxima de correos a traer
        """
        filters = []
        if only_unread:
            filters.append("isRead eq false")
        if from_address:
            safe = from_address.replace("'", "''")
            filters.append(f"from/emailAddress/address eq '{safe}'")

        if filters:
            endpoint = f"users/{self.mailbox}/mailFolders/inbox/messages?$top={limit}&$filter=" + " and ".join(filters)
        else:
            endpoint = f"users/{self.mailbox}/mailFolders/inbox/messages?$top={limit}&$orderby=receivedDateTime DESC"

        data = self.client.get(endpoint)
        raw_emails = data.get("value", [])

        # Filtro local por asunto (Graph API no soporta 'contains' en asunto fácilmente)
        if subject_contains:
            raw_emails = [
                e for e in raw_emails
                if subject_contains.lower() in e.get("subject", "").lower()
            ]

        emails = []
        for raw in raw_emails:
            attachments = []
            if include_attachments and raw.get("hasAttachments"):
                attachments = self._get_attachments(raw["id"])

            emails.append(self._parse_email(raw, attachments))

        return emails

    def start_polling(self, callback: Callable, poll_interval: int = 60, **kwargs):
        """
        Revisa correos nuevos cada cierto tiempo y llama callback con cada uno.

        Args:
            callback: función que recibe un Email
            poll_interval: segundos entre cada revisión
            **kwargs: mismos parámetros que get_new_emails()
        """
        logger.info(f"Iniciando polling cada {poll_interval} segundos...")
        while True:
            try:
                emails = self.get_new_emails(**kwargs)
                for email in emails:
                    callback(email)
            except Exception as e:
                logger.error(f"Error en polling: {e}")
            time.sleep(poll_interval)

    def _get_attachments(self, message_id: str) -> List[Attachment]:
        """Descarga los adjuntos de un correo"""
        endpoint = f"users/{self.mailbox}/messages/{message_id}/attachments"
        data = self.client.get(endpoint)
        attachments = []
        for att in data.get("value", []):
            attachments.append(Attachment(
                filename=att.get("name", ""),
                content_type=att.get("contentType", ""),
                content_bytes=att.get("contentBytes", "")
            ))
        return attachments

    def _parse_email(self, raw: dict, attachments: List[Attachment]) -> Email:
        """Convierte el JSON crudo de Microsoft en un objeto Email"""
        to_list = [
            r.get("emailAddress", {}).get("address", "")
            for r in raw.get("toRecipients", [])
        ]
        received_raw = raw.get("receivedDateTime", "")
        try:
            received_at = datetime.fromisoformat(received_raw.replace("Z", "+00:00"))
        except Exception:
            received_at = datetime.now()

        return Email(
            id=raw.get("id", ""),
            subject=raw.get("subject", ""),
            from_address=raw.get("from", {}).get("emailAddress", {}).get("address", ""),
            to=to_list,
            body=raw.get("body", {}).get("content", ""),
            received_at=received_at,
            attachments=attachments
        )