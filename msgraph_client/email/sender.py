import logging
from typing import List, Optional
from msgraph_client.auth.client import GraphClient

logger = logging.getLogger("msgraph_client")


class EmailSender:
    """Cliente para enviar correos desde Microsoft Outlook via Graph API"""

    def __init__(self, client: GraphClient, mailbox: str):
        """
        Args:
            client: instancia de GraphClient
            mailbox: email del buzón desde el que se envía (ej: "robot@empresa.com")
        """
        self.client = client
        self.mailbox = mailbox

    def send(
        self,
        to: List[str],
        subject: str,
        body: str,
        body_type: str = "HTML",
        cc: Optional[List[str]] = None,
        bcc: Optional[List[str]] = None,
        save_to_sent: bool = True,
    ) -> None:
        """
        Envía un correo electrónico.

        Args:
            to: lista de destinatarios principales
            subject: asunto del correo
            body: cuerpo del correo (HTML o texto plano)
            body_type: "HTML" o "Text". Por defecto "HTML"
            cc: lista de destinatarios en copia (opcional)
            bcc: lista de destinatarios en copia oculta (opcional)
            save_to_sent: si True, guarda el correo en Enviados

        Ejemplo:
            sender.send(
                to=["destino@empresa.com"],
                subject="Reporte listo",
                body="<p>El reporte está disponible ✅</p>",
                cc=["jefe@empresa.com"]
            )
        """
        payload = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": body_type,
                    "content": body,
                },
                "toRecipients": [
                    {"emailAddress": {"address": addr}} for addr in to
                ],
            },
            "saveToSentItems": save_to_sent,
        }

        if cc:
            payload["message"]["ccRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in cc
            ]

        if bcc:
            payload["message"]["bccRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in bcc
            ]

        endpoint = f"users/{self.mailbox}/sendMail"
        self.client.post(endpoint, payload)
        logger.info(f"✓ Correo enviado a {to} | Asunto: {subject}")

    def send_with_attachment(
        self,
        to: List[str],
        subject: str,
        body: str,
        filename: str,
        file_bytes: bytes,
        content_type: str = "application/octet-stream",
        body_type: str = "HTML",
        cc: Optional[List[str]] = None,
    ) -> None:
        """
        Envía un correo con un archivo adjunto.

        Args:
            to: lista de destinatarios
            subject: asunto del correo
            body: cuerpo del correo
            filename: nombre del archivo adjunto (ej: "reporte.xlsx")
            file_bytes: contenido del archivo en bytes
            content_type: tipo MIME del archivo
            body_type: "HTML" o "Text"
            cc: lista de destinatarios en copia (opcional)

        Ejemplo:
            with open("reporte.xlsx", "rb") as f:
                sender.send_with_attachment(
                    to=["destino@empresa.com"],
                    subject="Reporte",
                    body="<p>Adjunto el reporte.</p>",
                    filename="reporte.xlsx",
                    file_bytes=f.read(),
                    content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        """
        import base64

        payload = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": body_type,
                    "content": body,
                },
                "toRecipients": [
                    {"emailAddress": {"address": addr}} for addr in to
                ],
                "attachments": [
                    {
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "name": filename,
                        "contentType": content_type,
                        "contentBytes": base64.b64encode(file_bytes).decode("utf-8"),
                    }
                ],
            },
            "saveToSentItems": True,
        }

        if cc:
            payload["message"]["ccRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in cc
            ]

        endpoint = f"users/{self.mailbox}/sendMail"
        self.client.post(endpoint, payload)
        logger.info(f"✓ Correo con adjunto '{filename}' enviado a {to}")