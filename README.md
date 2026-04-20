# msgraph-client

Librería Python para integración con Microsoft Graph API.
Permite leer correos de Outlook y archivos de OneDrive/SharePoint de forma sencilla.

## Instalación

```bash
pip install git+https://github.com/tu-empresa/msgraph-client.git
```

## Configuración

Necesitas tres credenciales de tu aplicación en Azure:
- `client_id`
- `client_secret`
- `tenant_id`

## Uso básico

### Correos

```python
from msgraph_client import GraphAuth, GraphClient, EmailClient

auth = GraphAuth(
    client_id="tu-client-id",
    client_secret="tu-client-secret",
    tenant_id="tu-tenant-id"
)
client = GraphClient(auth)
email_client = EmailClient(client, mailbox="usuario@empresa.com")

# Obtener correos no leídos
emails = email_client.get_new_emails(only_unread=True)
for email in emails:
    print(email.subject, email.from_address)

# Obtener correos con adjuntos
emails = email_client.get_new_emails(include_attachments=True)

# Polling: revisar correos cada 60 segundos
def procesar(email):
    print(f"Nuevo correo: {email.subject}")

email_client.start_polling(callback=procesar, poll_interval=60)
```

### Archivos Excel (OneDrive)

```python
from msgraph_client import GraphAuth, GraphClient, DriveClient

auth = GraphAuth(
    client_id="tu-client-id",
    client_secret="tu-client-secret",
    tenant_id="tu-tenant-id"
)
client = GraphClient(auth)
drive_client = DriveClient(client, user_email="usuario@empresa.com")

# Leer hoja completa
rows = drive_client.read_sheet("Sheet1", path="Reportes/ventas.xlsx")

# Leer columna específica
precios = drive_client.read_column("Sheet1", column="Precio", path="Reportes/ventas.xlsx")

# Leer primeras 10 filas
filas = drive_client.read_rows("Sheet1", limit=10, path="Reportes/ventas.xlsx")
```

## Estructura del paquete

\```
msgraph_client/
├── auth/       → autenticación y cliente HTTP
├── email/      → lectura de correos
├── drive/      → lectura de archivos OneDrive/SharePoint
├── models/     → clases Email, Attachment, FileObject
└── utils/      → utilidades internas
\```

## Permisos requeridos en Azure

| Permiso | Módulo |
|---|---|
| `Mail.Read` | EmailClient |
| `Mail.ReadWrite` | EmailClient |
| `Files.Read.All` | DriveClient |