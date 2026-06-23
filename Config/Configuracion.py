from pydantic_settings import BaseSettings
from pydantic import Field
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


# ─────────────────────────────────────────────
#  MODELO DE CONFIGURACIÓN
# ─────────────────────────────────────────────
class Config(BaseSettings):

    # Metadata del bot
    nombre_bot : str = "AutorizacionMasivo"
    ambiente   : str = "DEV"

    # SAP
    sap_user       : str = Field(alias="SAPUSER")
    sap_password   : str = Field(alias="SAPPASSWORD")
    sap_mandante   : str = "410"
    sap_idioma     : str = "ES"
    sap_ruta_logon : str = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplgpad.exe"
    sap_entorno    : str = Field(alias="SAPENTORNO")

    # Base de datos
    db_user     : str = Field(alias="DBUSER")
    db_password : str = Field(alias="DBPASSWORD")
    db_server   : str = Field(alias="DBSERVER")
    db_name     : str = Field(alias="DBNAME")
    db_schema   : str = Field(alias="DBSCHEMA")

    # Configuracion del correo
    tenant_id     : str = Field(alias="TENANTID") 
    client_id     : str = Field(alias="CLIENTID") 
    client_secret : str = Field(alias="SECRETVALUE")
    graph_sender  : str = Field(alias="CORREO")

    # ── Credenciales SMTP ─────────────────────────────────────────────
    smtp_host     : str = Field(alias="SMTP_HOST") 
    smtp_port     : str = Field(alias="SMTP_PORT") 
    smtp_user     : str = Field(alias="SMTP_USER") 
    smtp_password : str = Field(alias="SMTP_PASSWORD") 

    # ── Credenciales portal NuevaEPS ────────────────────────────────
    nuevaeps_tipodoc : str = Field(default="CC",  alias="NUEVAEPS_TIPODOC")
    nuevaeps_usuario : str = Field(alias="NUEVAEPS_USUARIO")
    nuevaeps_password: str = Field(alias="NUEVAEPS_PASSWORD")

    # ── Credenciales portal Salud Total ───────────────────────────────
    saludtotal_tipodocips  : str = Field(default="NIT", alias="SALUDTOTAL_TIPODOCIPS")
    saludtotal_nitips      : str = Field(alias="SALUDTOTAL_NITIPS")
    saludtotal_tipodocuser : str = Field(default="CEDULA DE CIUDADANIA", alias="SALUDTOTAL_TIPODOCUSER")
    saludtotal_docusuario  : str = Field(alias="SALUDTOTAL_DOCUSUARIO")
    saludtotal_password    : str = Field(alias="SALUDTOTAL_PASSWORD")

    # ── Credenciales portal Sanitas (deltaasalud.tech) ────────────────
    sanitas_usuario : str = Field(alias="SANITAS_USUARIO")
    sanitas_password: str = Field(alias="SANITAS_PASSWORD")

    # ── Credenciales portal SanitasUrg por sucursal (colsanitas.com) ──
    sanitasurg_infantiluser  : str = Field(alias="SANITASURG_INFANTILUSER")
    sanitasurg_infantilpass  : str = Field(alias="SANITASURG_INFANTILPASS")
    sanitasurg_romauser      : str = Field(alias="SANITASURG_ROMAUSER")
    sanitasurg_romapass      : str = Field(alias="SANITASURG_ROMAPASS")
    sanitasurg_fusagasugauser: str = Field(alias="SANITASURG_FUSAGASUGAUSER")
    sanitasurg_fusagasugapass: str = Field(alias="SANITASURG_FUSAGASUGAPASS")
    sanitasurg_girardotuser  : str = Field(alias="SANITASURG_GIRARDOTUSER")
    sanitasurg_girardotpass  : str = Field(alias="SANITASURG_GIRARDOTPASS")

    model_config = {
        "populate_by_name" : True,
        "case_sensitive"   : False,
        "env_file"         : ".env",
        "env_file_encoding": "utf-8",
        "extra"            : "ignore",
    }


# ─────────────────────────────────────────────
#  INICIALIZACIÓN — corre una sola vez al importar
# ─────────────────────────────────────────────
def InicializarSettings() -> Config:
    """
    Carga los secretos del vault en dos pasos:
      1. Secretos compartidos (DB) → shared=true + environment=dev
      2. Secretos del proyecto     → project=AutorizacionesMasivo + environment=dev
    Los secretos del proyecto tienen prioridad en caso de colisión de nombres.
    """
    try:
        secretos_shared = CargarVault(
            filtro_tags  = {"shared": "true", "environment": "dev"},
            strip_prefix = "Dev"
        )
        secretos_proyecto = CargarVault(
            filtro_tags  = {"project": "AutorizacionesMasivo"},
            strip_prefix = "AutorizacionesMasivo"
        )

        secretos_sin_tag = CargarVault(
            nombres = ["ClientID", "TenantID", "SecretValue"]
        )

        # Merge — proyecto sobreescribe shared si hay colisión
        secretos = {**secretos_sin_tag, **secretos_shared, **secretos_proyecto}

    except Exception as e:
        raise RuntimeError(
            f"No se pudo conectar al Azure Key Vault: {e}\n"
            "Verifica conectividad y credenciales del vault."
        ) from e

    try:
        return Config(**secretos)
    except Exception as e:
        raise RuntimeError(
            f"Error construyendo Config desde vault: {e}\n"
            "Verifica que los nombres de los secretos en el vault coincidan "
            "con los aliases definidos en Config.\n"
            f"Secretos disponibles: {list(secretos.keys())}"
        ) from e


settings = InicializarSettings()