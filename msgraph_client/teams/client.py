import logging
from typing import List, Optional
from msgraph_client.auth.client import GraphClient

logger = logging.getLogger("msgraph_client")


class TeamsClient:
    """Cliente para interactuar con Microsoft Teams via Graph API"""

    def __init__(self, client: GraphClient):
        """
        Args:
            client: instancia de GraphClient
        """
        self.client = client

    # ─── Equipos y canales ──────────────────────────────────────────────────────

    def list_teams(self) -> List[dict]:
        """
        Lista todos los equipos de la organización.

        Returns:
            list: equipos con id, displayName, description
        """
        result = self.client.get("teams?$select=id,displayName,description")
        return result.get("value", [])

    def list_channels(self, team_id: str) -> List[dict]:
        """
        Lista los canales de un equipo.

        Args:
            team_id: ID del equipo

        Returns:
            list: canales con id, displayName
        """
        result = self.client.get(f"teams/{team_id}/channels")
        return result.get("value", [])

    def get_team(self, display_name: str) -> Optional[dict]:
        """
        Busca un equipo por nombre.

        Args:
            display_name: nombre del equipo (ej: "Operaciones")

        Returns:
            dict | None: equipo encontrado o None
        """
        teams = self.list_teams()
        for team in teams:
            if team.get("displayName", "").lower() == display_name.lower():
                return team
        return None

    def get_channel(self, team_id: str, channel_name: str) -> Optional[dict]:
        """
        Busca un canal dentro de un equipo por nombre.

        Args:
            team_id: ID del equipo
            channel_name: nombre del canal (ej: "General")

        Returns:
            dict | None: canal encontrado o None
        """
        channels = self.list_channels(team_id)
        for channel in channels:
            if channel.get("displayName", "").lower() == channel_name.lower():
                return channel
        return None

    # ─── Enviar mensajes a canales ──────────────────────────────────────────────

    def send_message(
        self,
        team_id: str,
        channel_id: str,
        message: str,
        message_type: str = "text",
    ) -> dict:
        """
        Envía un mensaje a un canal de Teams usando IDs.

        Args:
            team_id: ID del equipo
            channel_id: ID del canal
            message: texto del mensaje (puede ser HTML si message_type="html")
            message_type: "text" o "html"

        Returns:
            dict: metadatos del mensaje enviado
        """
        endpoint = f"teams/{team_id}/channels/{channel_id}/messages"
        payload = {
            "body": {
                "contentType": message_type,
                "content": message,
            }
        }
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

        Args:
            team_name: nombre del equipo (ej: "Operaciones")
            channel_name: nombre del canal (ej: "General")
            message: texto del mensaje
            message_type: "text" o "html"

        Returns:
            dict | None: metadatos del mensaje o None si no se encontró el equipo/canal

        Ejemplo:
            teams.send_message_by_name(
                team_name="Operaciones",
                channel_name="General",
                message="✅ El proceso terminó correctamente."
            )
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
        """
        Responde a un mensaje existente en un canal.

        Args:
            team_id: ID del equipo
            channel_id: ID del canal
            message_id: ID del mensaje al que se responde
            reply: texto de la respuesta
            message_type: "text" o "html"

        Returns:
            dict: metadatos de la respuesta enviada
        """
        endpoint = f"teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies"
        payload = {
            "body": {
                "contentType": message_type,
                "content": reply,
            }
        }
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
        Envía un mensaje directo (1 a 1) a un usuario de Teams.

        Args:
            user_email: email del destinatario
            message: texto del mensaje
            message_type: "text" o "html"

        Returns:
            dict | None: metadatos del mensaje enviado

        Ejemplo:
            teams.send_direct_message(
                user_email="dorian@empresa.com",
                message="Reporte subido a OneDrive ✅"
            )
        """
        try:
            chat = self._get_or_create_chat(user_email)
            endpoint = f"chats/{chat['id']}/messages"
            payload = {
                "body": {
                    "contentType": message_type,
                    "content": message,
                }
            }
            result = self.client.post(endpoint, payload)
            logger.info(f"✓ Mensaje directo enviado a {user_email}")
            return result
        except Exception as e:
            logger.error(f"❌ Error enviando mensaje directo a {user_email}: {e}")
            return None

    # ─── Leer mensajes ──────────────────────────────────────────────────────────

    def get_messages(self, team_id: str, channel_id: str, limit: int = 10) -> List[dict]:
        """
        Lee los mensajes más recientes de un canal.

        Args:
            team_id: ID del equipo
            channel_id: ID del canal
            limit: cantidad de mensajes a traer (máx 50)

        Returns:
            list: mensajes con id, body, from, createdDateTime
        """
        endpoint = f"teams/{team_id}/channels/{channel_id}/messages?$top={limit}"
        result = self.client.get(endpoint)
        return result.get("value", [])

    # ─── Métodos internos ───────────────────────────────────────────────────────

    def _get_or_create_chat(self, user_email: str) -> dict:
        """Crea o recupera un chat 1:1 con un usuario"""
        payload = {
            "chatType": "oneOnOne",
            "members": [
                {
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "roles": ["owner"],
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users/{user_email}",
                }
            ],
        }
        result = self.client.post("chats", payload)
        if result:
            return result

        # Si ya existe, buscarlo
        chats = self.client.get(f"users/{user_email}/chats?$filter=chatType eq 'oneOnOne'")
        items = chats.get("value", [])
        if items:
            return items[0]

        raise Exception(f"No se pudo crear ni encontrar chat con {user_email}")