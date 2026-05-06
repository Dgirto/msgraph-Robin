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
        self.client = client
        self.mailbox = mailbox

    def get_new_emails(self, only_unread=True, from_address=None, subject_contains=None, include_attachments=False, limit=10):
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

        if subject_contains:
            raw_emails = [e for e in raw_emails if subject_contains.lower() in e.get("subject", "").lower()]

        emails = []
        for raw in raw_emails:
            attachments = []
            if include_attachments and raw.get("hasAttachments"):
                attachments = self._get_attachments(raw["id"])
            emails.append(self._parse_email(raw, attachments))

        return emails

    def start_polling(self, callback: Callable, poll_interval: int = 60, **kwargs):
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
        endpoint = f"users/{self.mailbox}/messages/{message_id}/attachments"
        data = self.client.get(endpoint)
        return [Attachment(filename=a.get("name",""), content_type=a.get("contentType",""), content_bytes=a.get("contentBytes","")) for a in data.get("value", [])]

    def _parse_email(self, raw: dict, attachments: List[Attachment]) -> Email:
        to_list = [r.get("emailAddress", {}).get("address", "") for r in raw.get("toRecipients", [])]
        try:
            received_at = datetime.fromisoformat(raw.get("receivedDateTime", "").replace("Z", "+00:00"))
        except Exception:
            received_at = datetime.now()
        return Email(id=raw.get("id",""), subject=raw.get("subject",""), from_address=raw.get("from",{}).get("emailAddress",{}).get("address",""), to=to_list, body=raw.get("body",{}).get("content",""), received_at=received_at, attachments=attachments)