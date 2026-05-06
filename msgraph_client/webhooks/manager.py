import json
import logging
from datetime import datetime, timezone, timedelta
from typing import List, Optional
import requests
from msgraph_client.auth.client import GraphClient

logger = logging.getLogger("msgraph_client")


class WebhookManager:
    """
    Gestiona suscripciones a notificaciones push de Microsoft Graph.

    En lugar de hacer polling ("¿hay correos nuevos?"),
    Graph API avisa a tu endpoint cuando ocurre algo (correo nuevo, archivo cambiado, etc.).

    Flujo:
        1. subscribe_*()  → Microsoft empieza a notificar tu endpoint
        2. Tu endpoint recibe las notificaciones y las procesa
        3. renew_subscription() antes de que expire
        4. delete_subscription() cuando ya no se necesite
    """

    def __init__(self, client: GraphClient):
        """
        Args:
            client: instancia de GraphClient
        """
        self.client = client

    # ─── Crear suscripciones ────────────────────────────────────────────────────

    def subscribe_mail(
        self,
        mailbox: str,
        notification_url: str,
        secret: str,
        change_types: str = "created",
        expiration_minutes: int = 4230,
    ) -> dict:
        """
        Suscribe a notificaciones de correos nuevos en un buzón.

        Args:
            mailbox: email del buzón a monitorear
            notification_url: URL HTTPS pública que recibirá las notificaciones
            secret: clave secreta para validar notificaciones (mín 10 chars)
            change_types: "created", "updated", "deleted" o combinados con coma
            expiration_minutes: minutos hasta expiración (máx 4230 ≈ 3 días)

        Returns:
            dict: suscripción creada con id, expirationDateTime, resource

        Ejemplo:
            sub = manager.subscribe_mail(
                mailbox="robot@empresa.com",
                notification_url="https://miapp.com/webhooks/graph",
                secret="mi_clave_secreta_min10"
            )
            print(sub["id"])  # guardar para renovar/eliminar
        """
        resource = f"users/{mailbox}/mailFolders/inbox/messages"
        return self._create_subscription(resource, notification_url, secret, change_types, expiration_minutes)

    def subscribe_onedrive(
        self,
        user_email: str,
        notification_url: str,
        secret: str,
        change_types: str = "updated",
        expiration_minutes: int = 4230,
    ) -> dict:
        """
        Suscribe a notificaciones de cambios en el OneDrive de un usuario.

        Args:
            user_email: email del usuario cuyo OneDrive se monitorea
            notification_url: URL HTTPS pública receptora
            secret: clave secreta de validación
            change_types: "created", "updated", "deleted" o combinados
            expiration_minutes: minutos hasta expiración

        Returns:
            dict: suscripción creada
        """
        resource = f"users/{user_email}/drive/root"
        return self._create_subscription(resource, notification_url, secret, change_types, expiration_minutes)

    def subscribe_teams_channel(
        self,
        team_id: str,
        channel_id: str,
        notification_url: str,
        secret: str,
        change_types: str = "created",
    ) -> dict:
        """
        Suscribe a notificaciones de mensajes nuevos en un canal de Teams.

        Args:
            team_id: ID del equipo
            channel_id: ID del canal
            notification_url: URL HTTPS pública receptora
            secret: clave secreta de validación
            change_types: "created", "updated", "deleted"

        Returns:
            dict: suscripción creada
        """
        resource = f"teams/{team_id}/channels/{channel_id}/messages"
        return self._create_subscription(
            resource, notification_url, secret, change_types,
            expiration_minutes=60  # Teams tiene máximo 60 minutos
        )

    # ─── Gestionar suscripciones ────────────────────────────────────────────────

    def list_subscriptions(self) -> List[dict]:
        """
        Lista todas las suscripciones activas de la aplicación.

        Returns:
            list: suscripciones con id, resource, expirationDateTime
        """
        result = self.client.get("subscriptions")
        return result.get("value", [])

    def renew_subscription(self, subscription_id: str, expiration_minutes: int = 4230) -> dict:
        """
        Renueva una suscripción antes de que expire.
        Se recomienda llamar esta función con al menos 1 hora de anticipación.

        Args:
            subscription_id: ID de la suscripción a renovar
            expiration_minutes: nuevos minutos hasta expiración

        Returns:
            dict: suscripción actualizada con nueva expirationDateTime
        """
        from msgraph_client.auth import GraphAuth
        expiration = self._build_expiration(expiration_minutes)
        url = f"{GraphAuth.GRAPH_ENDPOINT}/subscriptions/{subscription_id}"
        payload = {"expirationDateTime": expiration}

        self.client.auth._authenticate()
        headers = self.client.auth.get_headers()
        response = requests.patch(url, headers=headers, json=payload)

        if response.status_code == 200:
            logger.info(f"✓ Suscripción {subscription_id} renovada hasta {expiration}")
            return response.json()
        raise Exception(f"Error renovando suscripción: {response.status_code} - {response.text}")

    def delete_subscription(self, subscription_id: str) -> bool:
        """
        Elimina una suscripción activa.

        Args:
            subscription_id: ID de la suscripción a eliminar

        Returns:
            bool: True si se eliminó correctamente
        """
        from msgraph_client.auth import GraphAuth
        self.client.auth._authenticate()
        url = f"{GraphAuth.GRAPH_ENDPOINT}/subscriptions/{subscription_id}"
        response = requests.delete(url, headers=self.client.auth.get_headers())

        if response.status_code == 204:
            logger.info(f"✓ Suscripción {subscription_id} eliminada")
            return True
        raise Exception(f"Error eliminando suscripción: {response.status_code} - {response.text}")

    def delete_all_subscriptions(self) -> int:
        """
        Elimina todas las suscripciones activas.

        Returns:
            int: cantidad de suscripciones eliminadas
        """
        subs = self.list_subscriptions()
        deleted = 0
        for sub in subs:
            if self.delete_subscription(sub["id"]):
                deleted += 1
        logger.info(f"✓ {deleted}/{len(subs)} suscripciones eliminadas")
        return deleted

    # ─── Validar notificaciones entrantes ───────────────────────────────────────

    @staticmethod
    def parse_notification(raw_body: bytes) -> List[dict]:
        """
        Parsea el cuerpo de una notificación entrante de Graph API.

        Args:
            raw_body: bytes del cuerpo de la request HTTP

        Returns:
            list: notificaciones con resource, changeType, clientState

        Ejemplo en Flask:
            notifications = WebhookManager.parse_notification(request.get_data())
        """
        data = json.loads(raw_body.decode("utf-8"))
        return data.get("value", [])

    @staticmethod
    def validate_notification(client_state: str, received_client_state: str) -> bool:
        """
        Valida que una notificación proviene de Microsoft Graph.
        Compara el secret enviado al crear la suscripción con el que llegó.

        Args:
            client_state: el secret que usaste al crear la suscripción
            received_client_state: el clientState que llegó en la notificación

        Returns:
            bool: True si la notificación es válida

        Ejemplo en Flask:
            for notif in notifications:
                if not WebhookManager.validate_notification(
                    client_state=SECRET,
                    received_client_state=notif.get("clientState", "")
                ):
                    return "Unauthorized", 401
        """
        import hmac
        return hmac.compare_digest(client_state, received_client_state)

    # ─── Métodos internos ───────────────────────────────────────────────────────

    def _create_subscription(
        self,
        resource: str,
        notification_url: str,
        secret: str,
        change_types: str,
        expiration_minutes: int,
    ) -> dict:
        """Crea una suscripción en Graph API"""
        payload = {
            "changeType": change_types,
            "notificationUrl": notification_url,
            "resource": resource,
            "expirationDateTime": self._build_expiration(expiration_minutes),
            "clientState": secret,
        }
        result = self.client.post("subscriptions", payload)
        logger.info(f"✓ Suscripción creada para '{resource}'")
        return result

    @staticmethod
    def _build_expiration(minutes: int) -> str:
        """Construye el datetime de expiración en formato ISO 8601 UTC"""
        expiration = datetime.now(timezone.utc) + timedelta(minutes=minutes)
        return expiration.strftime("%Y-%m-%dT%H:%M:%S.0000000Z")