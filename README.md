# Módulos nuevos — Extensiones Fase 2

Agrega esto al README.md existente, reemplazando la sección 9 (Futuras extensiones).

---

## Módulos disponibles

| Módulo | Clase | Descripción |
|---|---|---|
| `email_reader.py` | `EmailReader` | Lectura de correos desde Outlook/Exchange |
| `email_sender.py` | `EmailSender` | Envío de correos con CC, BCC y adjuntos |
| `onedrive_reader.py` | `OneDriveReader` | Lectura de archivos y Excel desde OneDrive/SharePoint |
| `onedrive_writer.py` | `OneDriveWriter` | Escritura y actualización de Excel en OneDrive |
| `teams_client.py` | `TeamsClient` | Envío de mensajes a canales y chats de Teams |
| `webhook_manager.py` | `WebhookManager` | Notificaciones push (reemplaza polling) |

---

## Permisos requeridos en Azure

| Permiso | Módulo | Para qué sirve |
|---|---|---|
| `Mail.Read` | EmailReader | Leer correos de cualquier buzón |
| `Mail.Send` | EmailSender | Enviar correos desde cualquier buzón |
| `Files.Read.All` | OneDriveReader | Leer archivos de OneDrive/SharePoint |
| `Files.ReadWrite.All` | OneDriveWriter | Escribir/subir archivos a OneDrive/SharePoint |
| `Team.ReadBasic.All` | TeamsClient | Listar equipos de la organización |
| `Channel.ReadBasic.All` | TeamsClient | Listar canales de un equipo |
| `ChannelMessage.Send` | TeamsClient | Enviar mensajes a canales |
| `Chat.ReadWrite.All` | TeamsClient | Enviar mensajes directos (1 a 1) |
| `User.Read.All` | GraphClient | Listar usuarios de la organización |

---

## Ejemplos de uso

### Envío de correos (EmailSender)

```python
from lib.graph.graph_client import GraphClient
from lib.graph.email_sender import EmailSender

client = GraphClient()
sender = EmailSender(client)

sender.send_email(
    from_email="robot@empresa.com",
    to_email="destino@empresa.com",
    subject="Reporte generado",
    body_html="<p>El reporte está listo ✅</p>",
    cc=["jefe@empresa.com"],
)
```

### Escritura en Excel / OneDrive (OneDriveWriter)

```python
from lib.graph.graph_client import GraphClient
from lib.graph.onedrive_writer import OneDriveWriter
import pandas as pd

client = GraphClient()
writer = OneDriveWriter(client)

df = pd.DataFrame({"Producto": ["A", "B"], "Ventas": [100, 200]})

# Subir nuevo archivo
writer.upload_dataframe(
    user_email="usuario@empresa.com",
    file_path="Reportes/ventas_mayo.xlsx",
    df=df,
    sheet_name="Mayo"
)

# Agregar filas a un Excel existente
nuevas_filas = pd.DataFrame({"Producto": ["C"], "Ventas": [300]})
writer.append_rows(
    user_email="usuario@empresa.com",
    file_path="Reportes/ventas_mayo.xlsx",
    df=nuevas_filas,
    sheet_name="Mayo"
)
```

### Integración con Teams (TeamsClient)

```python
from lib.graph.graph_client import GraphClient
from lib.graph.teams_client import TeamsClient

client = GraphClient()
teams = TeamsClient(client)

# Enviar a canal por nombre (sin necesitar IDs)
teams.send_channel_message_by_name(
    team_name="Operaciones",
    channel_name="General",
    message="✅ El proceso terminó correctamente."
)

# Enviar mensaje directo a un usuario
teams.send_direct_message(
    user_email="dorian@empresa.com",
    message="Reporte generado y subido a OneDrive."
)
```

### Webhooks — notificaciones push (WebhookManager)

```python
from lib.graph.graph_client import GraphClient
from lib.graph.webhook_manager import WebhookManager

client = GraphClient()
manager = WebhookManager(client)

# Suscribirse a correos nuevos
sub = manager.subscribe_mail(
    user_email="robot@empresa.com",
    notification_url="https://miapp.com/webhooks/graph",  # debe ser HTTPS público
    secret="mi_clave_secreta_min10chars",
)
print(f"Suscripción creada: {sub['id']}")

# Listar suscripciones activas
subs = manager.list_subscriptions()

# Renovar antes de que expire
manager.renew_subscription(sub["id"])

# Eliminar
manager.delete_subscription(sub["id"])
```

#### Endpoint receptor de webhooks (ejemplo Flask)

```python
from flask import Flask, request, jsonify
from lib.graph.webhook_manager import WebhookManager

app = Flask(__name__)
SECRET = "mi_clave_secreta_min10chars"

@app.route("/webhooks/graph", methods=["POST"])
def receive_notification():
    # Microsoft valida el endpoint con un GET primero
    validation_token = request.args.get("validationToken")
    if validation_token:
        return validation_token, 200, {"Content-Type": "text/plain"}

    # Procesar notificaciones
    notifications = WebhookManager.parse_notification(request.get_data())
    for notif in notifications:
        if not WebhookManager.validate_notification(
            raw_body=request.get_data(),
            client_state=SECRET,
            received_client_state=notif.get("clientState", ""),
        ):
            return "Unauthorized", 401

        resource = notif.get("resource")
        change_type = notif.get("changeType")
        print(f"Notificación: {change_type} en {resource}")
        # Aquí va tu lógica de negocio

    return jsonify({}), 202
```