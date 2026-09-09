"""Prueba de conexión del conector Microsoft 365 (Graph API).

Firma estándar Ruvic: def test_connection() -> tuple[bool, str]
- Lee la configuración EXCLUSIVAMENTE de las env vars RUVIC_MSGRAPH_*.
- Nunca lanza excepciones; retorna (ok, mensaje).

Ejecutable también como script para pruebas locales:
    python test_connection.py
"""

from __future__ import annotations

import os


def test_connection() -> tuple[bool, str]:
    """Obtiene un token de aplicación y consulta /organization en Graph."""
    try:
        from msgraph_client import GraphAuth, GraphClient
    except ImportError:
        return (
            False,
            "La librería msgraph-client no está instalada. "
            "Instala con: pip install "
            "git+https://github.com/Dgirto/msgraph-Robin.git",
        )

    client_id = (os.environ.get("RUVIC_MSGRAPH_CLIENT_ID") or "").strip()
    client_secret = (os.environ.get("RUVIC_MSGRAPH_CLIENT_SECRET") or "").strip()
    tenant_id = (os.environ.get("RUVIC_MSGRAPH_TENANT_ID") or "").strip()
    missing = [
        name
        for name, value in (
            ("RUVIC_MSGRAPH_CLIENT_ID", client_id),
            ("RUVIC_MSGRAPH_CLIENT_SECRET", client_secret),
            ("RUVIC_MSGRAPH_TENANT_ID", tenant_id),
        )
        if not value
    ]
    if missing:
        return False, f"Faltan variables de entorno: {', '.join(missing)}"

    try:
        auth = GraphAuth(
            client_id=client_id,
            client_secret=client_secret,
            tenant_id=tenant_id,
        )
        client = GraphClient(auth)
        data = client.get("organization")
    except Exception as exc:  # red de seguridad: jamás propagar
        return False, f"Autenticación o consulta a Graph fallida: {exc}"

    orgs = data.get("value") if isinstance(data, dict) else None
    if isinstance(orgs, list) and orgs:
        display = orgs[0].get("displayName") or orgs[0].get("id") or "organización"
        return True, f"Conexión exitosa a Microsoft Graph ({display})"
    return True, "Conexión exitosa a Microsoft Graph"


if __name__ == "__main__":
    ok, message = test_connection()
    print(f"{'OK' if ok else 'FALLO'}: {message}")
    raise SystemExit(0 if ok else 1)
