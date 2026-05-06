import logging
from datetime import datetime
from typing import List, Optional
from msgraph_client.auth.client import GraphClient

logger = logging.getLogger("msgraph_client")


class TeamsClient:
    """Cliente para interactuar con Microsoft Teams via Graph API"""

    def __init__(self, client: GraphClient):
        self.client = client

    # ─── Equipos y canales ──────────────────────────────────────────────────────

    def list_teams(self) -> List[dict]:
        """Lista todos los equipos de la organización."""
        result = self.client.get("teams?$select=id,displayName,description")
        return result.get("value", [])

    def list_channels(self, team_id: str) -> List[dict]:
        """Lista los canales de un equipo."""
        result = self.client.get(f"teams/{team_id}/channels")
        return result.get("value", [])

    def get_team(self, display_name: str) -> Optional[dict]:
        """Busca un equipo por nombre."""
        for team in self.list_teams():
            if team.get("displayName", "").lower() == display_name.lower():
                return team
        return None

    def get_channel(self, team_id: str, channel_name: str) -> Optional[dict]:
        """Busca un canal dentro de un equipo por nombre."""
        for channel in self.list_channels(team_id):
            if channel.get("displayName", "").lower() == channel_name.lower():
                return channel
        return None

    # ─── Mensajes a canales (Incoming Webhook) ──────────────────────────────────

    def send_webhook_message(
        self,
        webhook_url: str,
        message: str,
        title: Optional[str] = None,
        color: Optional[str] = None,
    ) -> bool:
        """
        Envía un mensaje a un canal via Incoming Webhook.
        No requiere permisos de Azure.

        Para obtener la URL:
            Canal de Teams → ... → Conectores → Incoming Webhook → Configurar

        Args:
            webhook_url: URL del Incoming Webhook del canal
            message: texto del mensaje
            title: título de la tarjeta (opcional)
            color: color del borde en hex (ej: "FF363A") (opcional)

        Ejemplo:
            teams.send_webhook_message(
                webhook_url="https://xxx.webhook.office.com/...",
                message="✅ El proceso terminó.",
                title="Robin AI",
                color="FF363A"
            )
        """
        import requests

        if title or color:
            payload = {
                "@type": "MessageCard",
                "@context": "http://schema.org/extensions",
                "themeColor": color or "0076D7",
                "summary": message,
                "sections": [{
                    "activityTitle": title or "",
                    "activityText": message,
                }]
            }
        else:
            payload = {"text": message}

        response = requests.post(webhook_url, json=payload)
        if response.status_code == 200:
            logger.info("✓ Mensaje enviado via webhook")
            return True
        raise Exception(f"Error enviando webhook: {response.status_code} - {response.text}")

    # ─── Mensajes a canales (Graph API) ─────────────────────────────────────────

    def send_message(
        self,
        team_id: str,
        channel_id: str,
        message: str,
        message_type: str = "text",
    ) -> dict:
        """
        Envía un mensaje a un canal usando IDs.
        Requiere permiso delegado: ChannelMessage.Send
        """
        endpoint = f"teams/{team_id}/channels/{channel_id}/messages"
        payload = {"body": {"contentType": message_type, "content": message}}
        result = self.client.post(endpoint, payload)
        logger.info(f"✓ Mensaje enviado al canal {channel_id}")
        return result

    def send_message_by_name(
        self,
        team_name: str,
        channel_name: str,
        message: str,
        message_type: str = "text",
    ) -> Optional[dict]:
        """
        Envía un mensaje a un canal buscando por nombre, sin necesitar IDs.
        Requiere permiso delegado: ChannelMessage.Send
        """
        team = self.get_team(team_name)
        if not team:
            logger.error(f"❌ Equipo '{team_name}' no encontrado.")
            return None
        channel = self.get_channel(team["id"], channel_name)
        if not channel:
            logger.error(f"❌ Canal '{channel_name}' no encontrado en '{team_name}'.")
            return None
        return self.send_message(team["id"], channel["id"], message, message_type)

    def reply_to_message(
        self,
        team_id: str,
        channel_id: str,
        message_id: str,
        reply: str,
        message_type: str = "text",
    ) -> dict:
        """Responde a un mensaje existente en un canal."""
        endpoint = f"teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies"
        payload = {"body": {"contentType": message_type, "content": reply}}
        result = self.client.post(endpoint, payload)
        logger.info(f"✓ Respuesta enviada al mensaje {message_id}")
        return result

    # ─── Mensajes directos ──────────────────────────────────────────────────────

    def send_direct_message(
        self,
        user_email: str,
        message: str,
        message_type: str = "text",
    ) -> Optional[dict]:
        """
        Envía un mensaje directo (1 a 1) a un usuario.
        Requiere permiso: Chat.ReadWrite.All

        Ejemplo:
            teams.send_direct_message(
                user_email="dorian@empresa.com",
                message="Reporte subido ✅"
            )
        """
        try:
            chat = self._get_or_create_chat(user_email)
            result = self.client.post(
                f"chats/{chat['id']}/messages",
                {"body": {"contentType": message_type, "content": message}}
            )
            logger.info(f"✓ Mensaje directo enviado a {user_email}")
            return result
        except Exception as e:
            logger.error(f"❌ Error enviando mensaje directo: {e}")
            return None

    # ─── Leer mensajes ──────────────────────────────────────────────────────────

    def get_channel_messages(
        self,
        team_id: str,
        channel_id: str,
        limit: int = 10,
    ) -> List[dict]:
        """
        Lee los mensajes más recientes de un canal.
        Requiere permiso: ChannelMessage.Read.All

        Ejemplo:
            mensajes = teams.get_channel_messages(team_id="...", channel_id="...", limit=10)
            for m in mensajes:
                print(m["body"]["content"])
        """
        result = self.client.get(f"teams/{team_id}/channels/{channel_id}/messages?$top={limit}")
        return result.get("value", [])

    def get_chat_messages(self, user_email: str, limit: int = 10) -> List[dict]:
        """
        Lee los chats recientes de un usuario.
        Requiere permiso: Chat.ReadWrite.All
        """
        result = self.client.get(f"users/{user_email}/chats?$top={limit}")
        return result.get("value", [])

    # ─── Reuniones ──────────────────────────────────────────────────────────────

    def create_meeting(
        self,
        user_email: str,
        subject: str,
        start: datetime,
        end: datetime,
        attendees: Optional[List[str]] = None,
    ) -> dict:
        """
        Crea una reunión de Teams y devuelve el link para unirse.
        Requiere permiso: OnlineMeetings.ReadWrite.All

        Args:
            user_email: email del organizador
            subject: título de la reunión
            start: fecha/hora de inicio (datetime con timezone)
            end: fecha/hora de fin (datetime con timezone)
            attendees: lista de emails de participantes (opcional)

        Returns:
            dict: reunión con joinWebUrl, id, subject

        Ejemplo:
            from datetime import datetime, timezone, timedelta

            reunion = teams.create_meeting(
                user_email="dorian@empresa.com",
                subject="Demo Robin AI",
                start=datetime.now(timezone.utc) + timedelta(hours=1),
                end=datetime.now(timezone.utc) + timedelta(hours=2),
                attendees=["cliente@empresa.com"]
            )
            print(reunion["joinWebUrl"])
        """
        payload = {
            "subject": subject,
            "startDateTime": start.strftime("%Y-%m-%dT%H:%M:%S.0000000Z"),
            "endDateTime": end.strftime("%Y-%m-%dT%H:%M:%S.0000000Z"),
        }
        if attendees:
            payload["participants"] = {
                "attendees": [{"upn": email, "role": "attendee"} for email in attendees]
            }

        result = self.client.post(f"users/{user_email}/onlineMeetings", payload)
        logger.info(f"✓ Reunión creada: {subject} | {result.get('joinWebUrl')}")
        return result

    def list_meetings(self, user_email: str) -> List[dict]:
        """
        Lista las reuniones online de un usuario.
        Requiere permiso: OnlineMeetings.ReadWrite.All
        """
        result = self.client.get(f"users/{user_email}/onlineMeetings")
        return result.get("value", [])

    def get_meeting(self, user_email: str, meeting_id: str) -> dict:
        """
        Obtiene los detalles de una reunión por su ID.
        Requiere permiso: OnlineMeetings.ReadWrite.All
        """
        return self.client.get(f"users/{user_email}/onlineMeetings/{meeting_id}")

    def cancel_meeting(self, user_email: str, meeting_id: str) -> bool:
        """
        Cancela una reunión de Teams.
        Requiere permiso: OnlineMeetings.ReadWrite.All
        """
        import requests
        from msgraph_client.auth import GraphAuth
        self.client.auth._authenticate()
        url = f"{GraphAuth.GRAPH_ENDPOINT}/users/{user_email}/onlineMeetings/{meeting_id}"
        response = requests.delete(url, headers=self.client.auth.get_headers())
        if response.status_code == 204:
            logger.info(f"✓ Reunión {meeting_id} cancelada")
            return True
        raise Exception(f"Error cancelando reunión: {response.status_code} - {response.text}")

    # ─── Métodos internos ───────────────────────────────────────────────────────

    def _get_or_create_chat(self, user_email: str) -> dict:
        """Crea o recupera un chat 1:1 con un usuario"""
        payload = {
            "chatType": "oneOnOne",
            "members": [{
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                "roles": ["owner"],
                "user@odata.bind": f"https://graph.microsoft.com/v1.0/users/{user_email}",
            }],
        }
        result = self.client.post("chats", payload)
        if result:
            return result
        chats = self.client.get(f"users/{user_email}/chats?$filter=chatType eq 'oneOnOne'")
        items = chats.get("value", [])
        if items:
            return items[0]
        raise Exception(f"No se pudo crear ni encontrar chat con {user_email}")