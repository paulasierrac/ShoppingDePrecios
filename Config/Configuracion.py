from azure.identity import ClientSecretCredential
from azure.keyvault.secrets import SecretClient
from dotenv import load_dotenv
import os

load_dotenv()

_VAULT_URL     = os.getenv("VAULT_URL")
_TENANT_ID     = os.getenv("TENANT_ID")
_CLIENT_ID     = os.getenv("CLIENT_ID")
_CLIENT_SECRET = os.getenv("CLIENT_SECRET")


# ─────────────────────────────────────────────
#  CARGA DE SECRETOS DESDE VAULT
# ─────────────────────────────────────────────
def CargarVault(filtro_tags: dict = None, strip_prefix: str = None, nombres: list[str] = None) -> dict:
    """
    Conecta al Key Vault y descarga secretos.

    Parámetros:
      filtro_tags  : Solo descarga secretos que tengan TODAS estas etiquetas.
                     Ej: {"project": "AutorizacionesMasivo", "environment": "dev"}
      strip_prefix : Elimina este prefijo del nombre antes de convertirlo a clave.
                     Ej: "AutorizacionesMasivo-SAPUser" → "SAPUSER"

    Retorna dict { clave: valor }
    """
    credential = ClientSecretCredential(
        tenant_id     = _TENANT_ID,
        client_id     = _CLIENT_ID,
        client_secret = _CLIENT_SECRET,
    )
    client   = SecretClient(vault_url=_VAULT_URL, credential=credential)
    secretos = {}

    for prop in client.list_properties_of_secrets():

        # ── Filtrar por etiquetas ──
        if filtro_tags:
            tags_secreto = prop.tags or {}
            if not all(tags_secreto.get(k) == v for k, v in filtro_tags.items()):
                continue

        # ── Filtrar por nombre exacto ──
        if nombres:
            if prop.name not in nombres:
                continue

        # ── Limpiar prefijo del nombre ──
        nombre = prop.name
        if strip_prefix and nombre.upper().startswith(strip_prefix.upper()):
            nombre = nombre[len(strip_prefix):]
            nombre = nombre.lstrip("-")

        valor  = client.get_secret(prop.name).value
        clave  = nombre.replace("-", "_").upper()
        secretos[clave] = valor

    return secretos