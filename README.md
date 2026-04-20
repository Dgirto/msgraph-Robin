# msgraph-client

Librería Python para integración con Microsoft Graph API.
Permite leer correos de Outlook y archivos de OneDrive/SharePoint de forma sencilla.

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

## Estructura del paquete

    msgraph_client/
    ├── auth/       autenticación y cliente HTTP
    ├── email/      lectura de correos
    ├── drive/      lectura de archivos OneDrive/SharePoint
    ├── models/     clases Email, Attachment, FileObject
    └── utils/      utilidades internas

## Permisos requeridos en Azure

| Permiso | Módulo | Para qué sirve |
|---|---|---|
| Mail.Read | EmailClient | Leer correos de cualquier buzón |
| Files.Read.All | DriveClient | Leer archivos de OneDrive/SharePoint |
| User.Read.All | GraphClient | Listar usuarios de la organización |