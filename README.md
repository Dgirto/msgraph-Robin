# msgraph-client

Librería Python para integración con Microsoft Graph API.
Permite leer y enviar correos, leer y escribir archivos en OneDrive, enviar mensajes a Teams y recibir notificaciones push.

## Instalación

    pip install git+https://github.com/Dgirto/msgraph-Robin.git

## Configuración

Necesitas tres credenciales de tu aplicación en Azure:
- `client_id`
- `client_secret`
- `tenant_id`

## Uso básico

### Autenticación

    from msgraph_client import GraphAuth, GraphClient

    auth = GraphAuth(
        client_id="tu-client-id",
        client_secret="tu-client-secret",
        tenant_id="tu-tenant-id"
    )
    client = GraphClient(auth)

### Logging

Configura el nivel de detalle de los logs. Llámalo una sola vez al inicio de tu script.

    from msgraph_client.utils import setup_logging

    setup_logging("INFO")  # opciones: DEBUG, INFO, WARNING, ERROR

Niveles disponibles:
- DEBUG: muestra todo, cada request y token. Útil para desarrollo.
- INFO: muestra autenticaciones y acciones importantes.
- WARNING: muestra solo cuando algo falla pero se recupera.
- ERROR: muestra solo errores graves. Recomendado para producción.

## EmailClient

    from msgraph_client import EmailClient

    email_client = EmailClient(client, mailbox="usuario@empresa.com")

### get_new_emails()

Obtiene correos de la bandeja de entrada.

    emails = email_client.get_new_emails(
        only_unread=True,           # solo no leídos (default: True)
        from_address="x@gmail.com", # filtrar por remitente (opcional)
        subject_contains="Reporte", # filtrar por asunto (opcional)
        include_attachments=True,   # incluir adjuntos (default: False)
        limit=10                    # cantidad máxima (default: 10)
    )

    for email in emails:
        print(email.id)
        print(email.subject)
        print(email.from_address)
        print(email.to)
        print(email.body)
        print(email.received_at)
        for att in email.attachments:
            print(att.filename)
            print(att.content_type)
            print(att.content_bytes)  # base64

### start_polling()

Revisa correos nuevos cada cierto tiempo y llama a un callback por cada uno.

    def procesar(email):
        print(f"Nuevo correo: {email.subject}")

    email_client.start_polling(
        callback=procesar,
        poll_interval=60  # segundos entre cada revisión
    )

## EmailSender

    from msgraph_client import EmailSender

    sender = EmailSender(client, mailbox="robot@empresa.com")

### send()

Envía un correo electrónico.

    sender.send(
        to=["destino@empresa.com"],
        subject="Reporte listo",
        body="<p>El reporte está disponible ✅</p>",
        body_type="HTML",           # "HTML" o "Text" (default: "HTML")
        cc=["jefe@empresa.com"],    # copia (opcional)
        bcc=["otro@empresa.com"],   # copia oculta (opcional)
        save_to_sent=True           # guardar en Enviados (default: True)
    )

### send_with_attachment()

Envía un correo con un archivo adjunto.

    with open("reporte.xlsx", "rb") as f:
        sender.send_with_attachment(
            to=["destino@empresa.com"],
            subject="Reporte",
            body="<p>Adjunto el reporte.</p>",
            filename="reporte.xlsx",
            file_bytes=f.read(),
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

## DriveClient

    from msgraph_client import DriveClient

    drive_client = DriveClient(client, user_email="usuario@empresa.com")

### get_file()

Obtiene metadata de un archivo por su ID.

    file = drive_client.get_file("01SVOVJV53SKD2EOOHHZDIIVMZHIP2ZHZA")
    print(file.id)
    print(file.name)
    print(file.size)
    print(file.mime_type)
    print(file.path)
    print(file.drive_id)

### get_file_by_path()

Obtiene metadata de un archivo por su ruta.

    file = drive_client.get_file_by_path("Reportes/ventas.xlsx")

### read_sheet()

Lee una hoja completa de un archivo Excel.

    # Con encabezado (devuelve nombres de columnas)
    rows = drive_client.read_sheet(
        "Sheet1",
        path="Reportes/ventas.xlsx",
        has_header=True  # default: True
    )
    # Resultado: [{"col1": "valor1", "col2": "valor2"}]

    # Sin encabezado (devuelve letras A, B, C...)
    rows = drive_client.read_sheet(
        "Sheet1",
        path="Reportes/ventas.xlsx",
        has_header=False
    )
    # Resultado: [{"A": "valor1", "B": "valor2"}]

### read_column()

Lee una columna específica de una hoja Excel.

    # Por nombre de columna (si has_header=True)
    valores = drive_client.read_column(
        "Sheet1",
        column="Precio",
        path="Reportes/ventas.xlsx"
    )

    # Por letra (si has_header=False)
    valores = drive_client.read_column(
        "Sheet1",
        column="A",
        path="Reportes/ventas.xlsx",
        has_header=False
    )

### read_rows()

Lee las primeras N filas de una hoja Excel.

    filas = drive_client.read_rows(
        "Sheet1",
        limit=10,
        path="Reportes/ventas.xlsx"
    )

## DriveWriter

    from msgraph_client import DriveWriter

    writer = DriveWriter(client, user_email="usuario@empresa.com")

### upload_sheet()

Convierte una lista de dicts en Excel y lo sube a OneDrive. Si el archivo existe, lo reemplaza.

    writer.upload_sheet(
        path="Reportes/ventas.xlsx",
        data=[{"Producto": "A", "Ventas": 100}, {"Producto": "B", "Ventas": 200}],
        sheet_name="Mayo"
    )

### upload_dataframe()

Sube un DataFrame de pandas como Excel a OneDrive.

    import pandas as pd

    df = pd.DataFrame({"Producto": ["A", "B"], "Ventas": [100, 200]})
    writer.upload_dataframe(
        path="Reportes/ventas.xlsx",
        df=df,
        sheet_name="Mayo"
    )

### append_rows()

Descarga un Excel existente, agrega filas al final y lo vuelve a subir.

    writer.append_rows(
        path="Reportes/ventas.xlsx",
        data=[{"Producto": "C", "Ventas": 300}],
        sheet_name="Mayo"
    )

### update_sheet()

Reemplaza el contenido de una hoja específica manteniendo las demás hojas intactas.

    writer.update_sheet(
        path="Reportes/ventas.xlsx",
        data=[{"Producto": "A", "Ventas": 999}],
        sheet_name="Mayo"
    )

### create_folder()

Crea una carpeta en OneDrive. Si ya existe, no hace nada.

    writer.create_folder("Reportes/2025")

## TeamsClient

    from msgraph_client import TeamsClient

    teams = TeamsClient(client)

### send_webhook_message()

Envía un mensaje a un canal via Incoming Webhook. No requiere permisos de Azure.

Para obtener la URL del webhook: Canal de Teams → ... → Conectores → Incoming Webhook → Configurar

    teams.send_webhook_message(
        webhook_url="https://xxx.webhook.office.com/...",
        message="✅ El proceso terminó.",
        title="Robin AI",    # opcional
        color="FF363A"       # opcional
    )

### send_message_by_name()

Envía un mensaje a un canal buscando por nombre, sin necesitar IDs.
Requiere permiso delegado: ChannelMessage.Send

    teams.send_message_by_name(
        team_name="Operaciones",
        channel_name="General",
        message="✅ El proceso terminó correctamente."
    )

### send_message()

Envía un mensaje a un canal usando IDs directamente.

    teams.send_message(
        team_id="...",
        channel_id="...",
        message="Hola desde la librería",
        message_type="text"  # "text" o "html"
    )

### send_direct_message()

Envía un mensaje directo (1 a 1) a un usuario de Teams.
Requiere permiso: Chat.ReadWrite.All

    teams.send_direct_message(
        user_email="dorian@empresa.com",
        message="Reporte subido a OneDrive ✅"
    )

### reply_to_message()

Responde a un mensaje existente en un canal.

    teams.reply_to_message(
        team_id="...",
        channel_id="...",
        message_id="...",
        reply="Recibido ✅"
    )

### list_teams() / list_channels()

Lista equipos y canales de la organización.

    equipos = teams.list_teams()
    canales = teams.list_channels(team_id="...")

### get_channel_messages()

Lee los mensajes más recientes de un canal.
Requiere permiso: ChannelMessage.Read.All

    mensajes = teams.get_channel_messages(
        team_id="...",
        channel_id="...",
        limit=10
    )
    for m in mensajes:
        print(m["body"]["content"])

### get_chat_messages()

Lee los chats recientes de un usuario.
Requiere permiso: Chat.ReadWrite.All

    chats = teams.get_chat_messages(user_email="dorian@empresa.com", limit=10)

### create_meeting()

Crea una reunión de Teams y devuelve el link para unirse.
Requiere permiso: OnlineMeetings.ReadWrite.All

    from datetime import datetime, timezone, timedelta

    reunion = teams.create_meeting(
        user_email="dorian@empresa.com",
        subject="Demo Robin AI",
        start=datetime.now(timezone.utc) + timedelta(hours=1),
        end=datetime.now(timezone.utc) + timedelta(hours=2),
        attendees=["cliente@empresa.com"]  # opcional
    )
    print(reunion["joinWebUrl"])  # link para unirse

### list_meetings() / get_meeting() / cancel_meeting()

    reuniones = teams.list_meetings(user_email="dorian@empresa.com")
    reunion   = teams.get_meeting(user_email="dorian@empresa.com", meeting_id="...")
    teams.cancel_meeting(user_email="dorian@empresa.com", meeting_id="...")

## WebhookManager

En lugar de hacer polling ("¿hay correos nuevos?" cada X segundos), Graph API avisa a tu endpoint cuando ocurre algo.

    from msgraph_client import WebhookManager

    manager = WebhookManager(client)

### subscribe_mail()

Suscribe a notificaciones de correos nuevos en un buzón.

    sub = manager.subscribe_mail(
        mailbox="robot@empresa.com",
        notification_url="https://miapp.com/webhooks/graph",  # debe ser HTTPS público
        secret="mi_clave_secreta_min10",
        change_types="created",         # "created", "updated", "deleted"
        expiration_minutes=4230         # máx 4230 (~3 días)
    )
    print(sub["id"])  # guardar este ID para renovar/eliminar

### subscribe_onedrive()

Suscribe a notificaciones de cambios en OneDrive.

    sub = manager.subscribe_onedrive(
        user_email="usuario@empresa.com",
        notification_url="https://miapp.com/webhooks/graph",
        secret="mi_clave_secreta_min10"
    )

### subscribe_teams_channel()

Suscribe a notificaciones de mensajes nuevos en un canal de Teams.

    sub = manager.subscribe_teams_channel(
        team_id="...",
        channel_id="...",
        notification_url="https://miapp.com/webhooks/graph",
        secret="mi_clave_secreta_min10"
    )

### list_subscriptions() / renew_subscription() / delete_subscription()

    # Listar todas las activas
    subs = manager.list_subscriptions()

    # Renovar antes de que expire
    manager.renew_subscription(sub["id"])

    # Eliminar una
    manager.delete_subscription(sub["id"])

    # Eliminar todas
    manager.delete_all_subscriptions()

### Endpoint receptor (ejemplo Flask)

    from flask import Flask, request, jsonify
    from msgraph_client import WebhookManager

    app = Flask(__name__)
    SECRET = "mi_clave_secreta_min10"

    @app.route("/webhooks/graph", methods=["POST"])
    def recibir():
        # Microsoft valida el endpoint con este parámetro primero
        token = request.args.get("validationToken")
        if token:
            return token, 200, {"Content-Type": "text/plain"}

        notifs = WebhookManager.parse_notification(request.get_data())
        for n in notifs:
            if not WebhookManager.validate_notification(
                client_state=SECRET,
                received_client_state=n.get("clientState", "")
            ):
                return "Unauthorized", 401
            print(f"Notificación: {n.get('changeType')} en {n.get('resource')}")

        return jsonify({}), 202

    app.run(port=5000)

## Estructura del paquete

    msgraph_client/
    ├── auth/       autenticación y cliente HTTP
    ├── email/      lectura y envío de correos
    ├── drive/      lectura y escritura en OneDrive/SharePoint
    ├── teams/      mensajes, reuniones e integración con Teams
    ├── webhooks/   notificaciones push (reemplaza polling)
    ├── models/     clases Email, Attachment, FileObject
    └── utils/      utilidades internas

## Permisos requeridos en Azure

| Permiso | Módulo | Para qué sirve |
|---|---|---|
| Mail.Read | EmailClient | Leer correos de cualquier buzón |
| Mail.Send | EmailSender | Enviar correos desde cualquier buzón |
| Files.Read.All | DriveClient | Leer archivos de OneDrive/SharePoint |
| Files.ReadWrite.All | DriveWriter | Escribir archivos en OneDrive/SharePoint |
| Team.ReadBasic.All | TeamsClient | Listar equipos de la organización |
| Channel.ReadBasic.All | TeamsClient | Listar canales de un equipo |
| ChannelMessage.Read.All | TeamsClient | Leer mensajes de canales |
| Chat.ReadWrite.All | TeamsClient | Mensajes directos y leer chats |
| OnlineMeetings.ReadWrite.All | TeamsClient | Crear y cancelar reuniones de Teams |
| User.Read.All | GraphClient | Listar usuarios de la organización |
| Organization.Read.All | test_connection | Probar conexión (`GET /organization`) |

## Conector Ruvic

Este repo es el conector `msgraph` del catálogo Ruvic. Archivos de entrega:

- `manifest.json` — id `msgraph`, modo `azure_app`, prefijo `RUVIC_MSGRAPH_`
- `docs/index.html` — guía bilingüe del portal (`/catalog/msgraph`)
- `SKILL.md` — instrucciones para el agente
- `test_connection.py` — `GET /v1.0/organization` con las env vars `RUVIC_MSGRAPH_*`

Instalación en runtime (paquete en la raíz, sin `#subdirectory=lib`):

    pip install git+https://github.com/Dgirto/msgraph-Robin.git