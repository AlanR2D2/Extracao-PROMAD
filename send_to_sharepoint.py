import os
import time
import msal
import requests
from monitoring import log
from dotenv import load_dotenv

load_dotenv()

TENANT_ID = os.getenv("SP_TENANT_ID")
CLIENT_ID = os.getenv("SP_CLIENT_ID")
CLIENT_SECRET = os.getenv("SP_CLIENT_SECRET")
SITE_URL = os.getenv("SP_SITE_URL")  # https://doutoraconsumidor.sharepoint.com/sites/DoutoraConsumidor

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

_token_cache = None


def _get_access_token():
    """Obtém token de acesso via client credentials (application-level)."""
    global _token_cache
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(scopes=SCOPE)
    if "access_token" in result:
        _token_cache = result["access_token"]
        return _token_cache
    else:
        error_msg = result.get("error_description", result.get("error", "Erro desconhecido"))
        raise RuntimeError(f"Falha ao obter token SharePoint: {error_msg}")


def _get_site_id(token):
    """Obtém o site ID do SharePoint a partir da URL."""
    # Extrair hostname e path do site
    # URL: https://doutoraconsumidor.sharepoint.com/sites/DoutoraConsumidor
    from urllib.parse import urlparse
    parsed = urlparse(SITE_URL)
    hostname = parsed.hostname  # doutoraconsumidor.sharepoint.com
    site_path = parsed.path.rstrip("/")  # /sites/DoutoraConsumidor

    url = f"{GRAPH_BASE}/sites/{hostname}:{site_path}"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    site_id = resp.json()["id"]
    log.info(f"SharePoint Site ID obtido: {site_id}")
    return site_id


def _get_drive_id(token, site_id):
    """Obtém o drive ID da biblioteca 'Documentos Compartilhados'."""
    url = f"{GRAPH_BASE}/sites/{site_id}/drives"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    drives = resp.json().get("value", [])
    for drive in drives:
        # O nome pode ser "Documentos" ou "Documents" ou "Documentos Compartilhados"
        if drive.get("name") in ("Documentos Compartilhados", "Documents", "Documentos"):
            log.info(f"Drive encontrado: '{drive['name']}' (id: {drive['id']})")
            return drive["id"]
    # Fallback: pegar o primeiro drive
    if drives:
        log.warning(f"Drive padrão não encontrado, usando: '{drives[0]['name']}'")
        return drives[0]["id"]
    raise RuntimeError("Nenhum drive encontrado no site SharePoint")


def upload_file(local_path, remote_folder, remote_filename):
    """
    Faz upload de um arquivo local para o SharePoint, substituindo se já existir.
    
    Args:
        local_path: Caminho local do arquivo (.xlsx)
        remote_folder: Pasta no SharePoint (ex: "Planilhas Aerolines")
        remote_filename: Nome do arquivo no destino (ex: "Promad_aeroline.xlsx")
    """
    log.info(f"Iniciando upload para SharePoint: {remote_folder}/{remote_filename}")

    token = _get_access_token()
    site_id = _get_site_id(token)
    drive_id = _get_drive_id(token, site_id)

    # Upload via PUT (substitui se existir) - para arquivos até 4MB
    # Para arquivos maiores, usar upload session
    file_size = os.path.getsize(local_path)
    remote_path = f"{remote_folder}/{remote_filename}"

    if file_size < 4 * 1024 * 1024:  # < 4MB
        url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{remote_path}:/content"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }
        max_retries = 3
        for attempt in range(1, max_retries + 1):
            with open(local_path, "rb") as f:
                resp = requests.put(url, headers=headers, data=f)
            if resp.status_code == 423 and attempt < max_retries:
                log.warning(f"[RETRY] Arquivo bloqueado (423), tentativa {attempt}/{max_retries}. Aguardando 10s...")
                time.sleep(10)
                continue
            resp.raise_for_status()
            break
    else:
        # Upload session para arquivos grandes
        url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{remote_path}:/createUploadSession"
        headers = {"Authorization": f"Bearer {token}"}
        body = {"item": {"@microsoft.graph.conflictBehavior": "replace"}}
        resp = requests.post(url, headers=headers, json=body)
        resp.raise_for_status()
        upload_url = resp.json()["uploadUrl"]

        chunk_size = 10 * 1024 * 1024  # 10MB chunks
        with open(local_path, "rb") as f:
            start = 0
            while True:
                chunk = f.read(chunk_size)
                if not chunk:
                    break
                end = start + len(chunk) - 1
                content_range = f"bytes {start}-{end}/{file_size}"
                resp = requests.put(
                    upload_url,
                    headers={"Content-Range": content_range},
                    data=chunk,
                )
                resp.raise_for_status()
                start = end + 1

    log.info(f"[OK] Upload concluído: {remote_path}")
