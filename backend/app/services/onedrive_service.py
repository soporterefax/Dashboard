import io
import os
import threading
from dataclasses import dataclass
from typing import Optional
from urllib.parse import quote


GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


@dataclass
class ExcelCache:
    content: Optional[bytes] = None
    etag: Optional[str] = None


_cache = ExcelCache()
_cache_lock = threading.Lock()


# =========================================================
# VARIABLES DE ENTORNO
# =========================================================

def _required_env(name: str) -> str:
    value = os.getenv(name, "").strip()

    if not value:
        raise RuntimeError(
            f"Falta configurar la variable de entorno {name}"
        )

    return value


# =========================================================
# TOKEN MICROSOFT GRAPH
# =========================================================

def _get_access_token() -> str:
    import msal

    tenant_id = _required_env("MS_TENANT_ID")
    client_id = _required_env("MS_CLIENT_ID")
    client_secret = _required_env("MS_CLIENT_SECRET")

    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential=client_secret,
    )

    result = app.acquire_token_silent(
        scopes=[
            "https://graph.microsoft.com/.default"
        ],
        account=None,
    )

    if not result:
        result = app.acquire_token_for_client(
            scopes=[
                "https://graph.microsoft.com/.default"
            ]
        )

    token = result.get("access_token")

    if not token:
        detail = (
            result.get("error_description")
            or result.get("error")
            or "Error desconocido"
        )

        raise RuntimeError(
            f"No se pudo obtener token de Microsoft Graph: {detail}"
        )

    return token


# =========================================================
# URL DEL EXCEL
# =========================================================

def _item_url() -> str:
    drive_id = _required_env("MS_DRIVE_ID")

    item_id = os.getenv(
        "MS_EXCEL_ITEM_ID",
        ""
    ).strip()

    item_path = (
        os.getenv(
            "MS_EXCEL_PATH",
            ""
        )
        .strip()
        .strip("/")
    )

    if item_id:
        return (
            f"{GRAPH_BASE_URL}"
            f"/drives/{drive_id}"
            f"/items/{item_id}"
        )

    if item_path:
        encoded_path = quote(
            item_path,
            safe="/"
        )

        return (
            f"{GRAPH_BASE_URL}"
            f"/drives/{drive_id}"
            f"/root:/{encoded_path}"
        )

    raise RuntimeError(
        "Configura MS_EXCEL_ITEM_ID o MS_EXCEL_PATH"
    )


# =========================================================
# HEADERS
# =========================================================

def _headers(token: str) -> dict:
    return {
        "Authorization": f"Bearer {token}"
    }


# =========================================================
# DESCARGAR EXCEL
# =========================================================

def obtener_excel_onedrive(
    force_refresh: bool = False
) -> io.BytesIO:

    import requests

    """
    Obtiene el Excel privado desde SharePoint mediante
    Microsoft Graph.

    Usa eTag para evitar descargas innecesarias.
    """

    token = _get_access_token()
    item_url = _item_url()

    with _cache_lock:

        metadata_response = requests.get(
            (
                f"{item_url}"
                "?$select="
                "id,name,eTag,lastModifiedDateTime,size"
            ),
            headers=_headers(token),
            timeout=30,
        )

        metadata_response.raise_for_status()

        metadata = metadata_response.json()

        current_etag = metadata.get("eTag")

        # -----------------------------------------
        # CACHE
        # -----------------------------------------

        if (
            not force_refresh
            and _cache.content is not None
            and current_etag
            and current_etag == _cache.etag
        ):
            return io.BytesIO(
                _cache.content
            )

        # -----------------------------------------
        # DESCARGA
        # -----------------------------------------

        download_response = requests.get(
            f"{item_url}/content",
            headers=_headers(token),
            timeout=60,
            allow_redirects=True,
        )

        download_response.raise_for_status()

        _cache.content = (
            download_response.content
        )

        _cache.etag = current_etag

        return io.BytesIO(
            _cache.content
        )


# =========================================================
# ESTADO / METADATOS
# =========================================================

def estado_fuente_onedrive() -> dict:

    import requests

    """
    Devuelve metadatos del Excel sin exponer secretos.
    """

    token = _get_access_token()
    item_url = _item_url()

    response = requests.get(
        (
            f"{item_url}"
            "?$select="
            "id,name,eTag,lastModifiedDateTime,size"
        ),
        headers=_headers(token),
        timeout=30,
    )

    response.raise_for_status()

    data = response.json()

    return {
        "fuente": "onedrive",
        "archivo": data.get("name"),
        "ultima_modificacion":
            data.get("lastModifiedDateTime"),
        "tamano_bytes":
            data.get("size"),
        "etag":
            data.get("eTag"),
    }


# =========================================================
# SUBIR / ACTUALIZAR EXCEL
# =========================================================

def subir_excel_onedrive(
    contenido: bytes,
    etag_esperado: Optional[str] = None,
) -> dict:

    import requests

    """
    Reemplaza el contenido del Excel existente en SharePoint.

    Si se proporciona etag_esperado, usa If-Match para evitar
    sobrescribir cambios realizados por otra persona.
    """

    if not contenido:
        raise ValueError(
            "El contenido del Excel está vacío."
        )

    token = _get_access_token()
    item_url = _item_url()

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type":
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet",
    }

    # -----------------------------------------
    # CONTROL DE CONCURRENCIA
    # -----------------------------------------

    if etag_esperado:
        headers["If-Match"] = etag_esperado

    response = requests.put(
        f"{item_url}/content",
        headers=headers,
        data=contenido,
        timeout=120,
    )

    # -----------------------------------------
    # EL ARCHIVO CAMBIÓ ANTES DE GUARDAR
    # -----------------------------------------

    if response.status_code == 412:

        raise RuntimeError(
            "El archivo fue modificado por otra persona "
            "antes de guardar. Recarga la información "
            "e intenta nuevamente."
        )

    response.raise_for_status()

    data = response.json()

    nuevo_etag = data.get("eTag")

    # -----------------------------------------
    # ACTUALIZAR CACHE
    # -----------------------------------------

    with _cache_lock:

        _cache.content = contenido
        _cache.etag = nuevo_etag

    return {
        "ok": True,
        "archivo": data.get("name"),
        "etag": nuevo_etag,
        "ultima_modificacion":
            data.get("lastModifiedDateTime"),
        "tamano_bytes":
            data.get("size"),
    }


# =========================================================
# LIMPIAR CACHE
# =========================================================

def invalidar_cache_onedrive():

    global _cache

    with _cache_lock:

        _cache = ExcelCache()