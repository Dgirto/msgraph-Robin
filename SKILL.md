---
name: msgraph
description: >
  Usa la librería msgraph_client (Microsoft Graph) para leer y enviar correo
  de Outlook, leer y escribir archivos en OneDrive (incluido Excel), enviar
  mensajes y crear reuniones en Teams, y suscribirse a webhooks de Graph.
  Úsala cuando el usuario pida automatizar Microsoft 365: Outlook, OneDrive,
  SharePoint, Teams o notificaciones push de Graph.
triggers:
- microsoft 365
- msgraph
- graph api
- outlook
- onedrive
- teams
- sharepoint
- excel onedrive
---

# Conector Microsoft 365 (msgraph_client)

Librería Python para Microsoft Graph con flujo **client credentials** (aplicación, no usuario interactivo). Está **preinstalada en el runtime** cuando el conector está configurado. Si no:

```bash
pip install git+https://github.com/Dgirto/msgraph-Robin.git
```

El paquete está en la raíz del repo (no uses `#subdirectory=lib`).

## Regla crítica de credenciales

El código generado **NUNCA hardcodea credenciales**. Siempre se leen de variables de entorno, disponibles cuando el conector `msgraph` está configurado:

| Variable | Contenido |
|----------|-----------|
| `RUVIC_MSGRAPH_CLIENT_ID` | Application (client) ID de Entra ID |
| `RUVIC_MSGRAPH_CLIENT_SECRET` | Client secret (Value, no Secret ID) |
| `RUVIC_MSGRAPH_TENANT_ID` | Directory (tenant) ID |
| `RUVIC_MSGRAPH_DEFAULT_MAILBOX` | (opcional) buzón por defecto, p. ej. `robot@empresa.com` |

Si estas variables NO existen, el conector no está configurado: no generes código que lo use; indica al usuario que lo configure en **Settings → Conectores**.

El código generado **NUNCA** usa nombres de variable con segmento de alias (`_DEFAULT_`, `_TEST_`, `_PRODUCCION_`, etc.) — siempre `RUVIC_MSGRAPH_{CAMPO}` tal cual.

## Conexión (siempre igual)

`msgraph_client` no lee el entorno solo: hay que pasar las variables a `GraphAuth`.

```python
import os
from msgraph_client import GraphAuth, GraphClient

auth = GraphAuth(
    client_id=os.environ["RUVIC_MSGRAPH_CLIENT_ID"],
    client_secret=os.environ["RUVIC_MSGRAPH_CLIENT_SECRET"],
    tenant_id=os.environ["RUVIC_MSGRAPH_TENANT_ID"],
)
client = GraphClient(auth)
mailbox = os.environ.get("RUVIC_MSGRAPH_DEFAULT_MAILBOX") or "robot@empresa.com"
```

## Capacidad 1 — Leer correos (EmailClient)

```python
from msgraph_client import EmailClient

email_client = EmailClient(client, mailbox=mailbox)
emails = email_client.get_new_emails(
    only_unread=True,
    from_address=None,
    subject_contains=None,
    include_attachments=False,
    limit=10,
)
for email in emails:
    print(email.id, email.subject, email.from_address)
```

## Capacidad 2 — Enviar correos (EmailSender)

```python
from msgraph_client import EmailSender

sender = EmailSender(client, mailbox=mailbox)
sender.send(
    to=["destino@empresa.com"],
    subject="Reporte listo",
    body="<p>El reporte está disponible</p>",
    body_type="HTML",
)
```

Con adjunto: `sender.send_with_attachment(...)` (filename, file_bytes, content_type).

## Capacidad 3 — OneDrive / Excel (DriveClient, DriveWriter)

```python
from msgraph_client import DriveClient, DriveWriter

drive = DriveClient(client, user_email=mailbox)
rows = drive.read_sheet("Sheet1", path="Reportes/ventas.xlsx", has_header=True)

writer = DriveWriter(client, user_email=mailbox)
writer.upload_sheet(
    path="Reportes/ventas.xlsx",
    data=[{"Producto": "A", "Ventas": 100}],
    sheet_name="Mayo",
)
```

También: `read_column`, `read_rows`, `get_file_by_path`, `append_rows`, `update_sheet`, `create_folder`, `upload_dataframe`.

## Capacidad 4 — Teams (TeamsClient)

```python
from msgraph_client import TeamsClient

teams = TeamsClient(client)
teams.send_direct_message(user_email="persona@empresa.com", message="Listo")
teams.send_message_by_name(
    team_name="Operaciones",
    channel_name="General",
    message="El proceso terminó",
)
```

Incoming Webhook (sin permisos de Graph): `teams.send_webhook_message(webhook_url=..., message=...)`.

Reuniones: `create_meeting`, `list_meetings`, `cancel_meeting` (permiso `OnlineMeetings.ReadWrite.All`).

## Capacidad 5 — Webhooks (WebhookManager)

```python
from msgraph_client import WebhookManager

manager = WebhookManager(client)
sub = manager.subscribe_mail(
    mailbox=mailbox,
    notification_url="https://miapp.com/webhooks/graph",
    secret="clave_secreta_min10",
)
```

La URL de notificación debe ser HTTPS público. Microsoft envía primero un `validationToken` que hay que devolver en texto plano.

## Permisos

Solo concede en Azure los Application permissions que el flujo vaya a usar (`Mail.Read`, `Mail.Send`, `Files.ReadWrite.All`, `Chat.ReadWrite.All`, `ChannelMessage.Send`, `OnlineMeetings.ReadWrite.All`, etc.) y otorga **admin consent**.
