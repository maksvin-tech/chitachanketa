"""
Spremanje rezultata i skenova u Google Drive.

Struktura mapa koju program stvara u vašem Google Driveu:

    OMR Ankete/
        Rezultati/<vrsta obrasca>/<datum_vrijeme>_<računalo>.xlsx
        Skenovi/<vrsta obrasca>/<datum_vrijeme>_<računalo>/<izvorne slike>

Prijava:
    - google_credentials.json  -> OAuth klijent iz Google Cloud konzole (vidi PODESAVANJE_GOOGLE_DRIVE.md)
    - google_token.json        -> nastaje sam nakon prve prijave na svakom računalu

Program koristi ovlast "drive.file": vidi SAMO datoteke koje je sam stvorio,
ne i ostatak vašeg Drivea.
"""

import io
import os
import socket
from datetime import datetime

import pandas as pd

APP_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_FILE = os.path.join(APP_DIR, "google_credentials.json")
TOKEN_FILE = os.path.join(APP_DIR, "google_token.json")

SCOPES = ["https://www.googleapis.com/auth/drive.file"]
ROOT_FOLDER = "OMR Ankete"
RESULTS_FOLDER = "Rezultati"
SCANS_FOLDER = "Skenovi"

FOLDER_MIME = "application/vnd.google-apps.folder"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class DriveNotConfigured(Exception):
    """Nedostaje google_credentials.json ili Google biblioteke nisu instalirane."""


# ------------------------------------------------------------------------------
# Prijava
# ------------------------------------------------------------------------------

def libraries_installed():
    try:
        import googleapiclient  # noqa: F401
        import google_auth_oauthlib  # noqa: F401
        return True
    except ImportError:
        return False


def is_configured():
    return os.path.exists(CREDENTIALS_FILE) and libraries_installed()


def has_saved_login():
    return os.path.exists(TOKEN_FILE)


def get_service(interactive=True):
    """Vraća Drive servis. Ako prijava ne postoji i interactive=True, otvara preglednik za prijavu.
    Ako interactive=False i nema valjane prijave, vraća None."""
    if not libraries_installed():
        raise DriveNotConfigured(
            "Google biblioteke nisu instalirane. Pokrenite: pip install -r requirements.txt"
        )
    if not os.path.exists(CREDENTIALS_FILE):
        raise DriveNotConfigured(
            f"U mapi programa nedostaje datoteka {os.path.basename(CREDENTIALS_FILE)}."
        )

    from google.auth.exceptions import RefreshError
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

    creds = None
    if os.path.exists(TOKEN_FILE):
        try:
            creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        except (ValueError, OSError):
            creds = None

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except RefreshError:
            creds = None  # Prijava je istekla ili je opozvana -> treba se ponovno prijaviti

    if not creds or not creds.valid:
        if not interactive:
            return None
        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
        creds = flow.run_local_server(port=0, open_browser=True)

    with open(TOKEN_FILE, "w", encoding="utf-8") as f:
        f.write(creds.to_json())

    return build("drive", "v3", credentials=creds, cache_discovery=False)


def logout():
    if os.path.exists(TOKEN_FILE):
        os.remove(TOKEN_FILE)


def account_email(service):
    try:
        about = service.about().get(fields="user(emailAddress)").execute()
        return about["user"]["emailAddress"]
    except Exception:
        return None


# ------------------------------------------------------------------------------
# Mape i datoteke
# ------------------------------------------------------------------------------

def _q(value):
    """Zaštita navodnika u Drive upitima."""
    return value.replace("\\", "\\\\").replace("'", "\\'")


def _list(service, query, fields="id, name, createdTime, webViewLink"):
    items, token = [], None
    while True:
        resp = service.files().list(
            q=query,
            spaces="drive",
            fields=f"nextPageToken, files({fields})",
            orderBy="createdTime",
            pageSize=1000,
            pageToken=token,
        ).execute()
        items.extend(resp.get("files", []))
        token = resp.get("nextPageToken")
        if not token:
            return items


def _find_folder(service, name, parent_id=None):
    query = f"name = '{_q(name)}' and mimeType = '{FOLDER_MIME}' and trashed = false"
    if parent_id:
        query += f" and '{parent_id}' in parents"
    found = _list(service, query)
    return found[0] if found else None  # Najstarija, ako ih je slučajno više


def _ensure_folder(service, name, parent_id=None):
    folder = _find_folder(service, name, parent_id)
    if folder:
        return folder
    body = {"name": name, "mimeType": FOLDER_MIME}
    if parent_id:
        body["parents"] = [parent_id]
    return service.files().create(body=body, fields="id, name, webViewLink").execute()


def _ensure_path(service, *names):
    parent = None
    for name in names:
        parent = _ensure_folder(service, name, parent["id"] if parent else None)
    return parent


def _upload(service, name, data, mimetype, parent_id):
    from googleapiclient.http import MediaIoBaseUpload
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mimetype, resumable=True)
    body = {"name": name, "parents": [parent_id]}
    return service.files().create(body=body, media_body=media, fields="id, name, webViewLink").execute()


def _replace_content(service, file_id, data, mimetype):
    from googleapiclient.http import MediaIoBaseUpload
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mimetype, resumable=True)
    return service.files().update(fileId=file_id, media_body=media, fields="id, name, webViewLink").execute()


def _guess_mime(filename):
    ext = os.path.splitext(filename)[1].lower()
    return {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg"}.get(ext, "application/octet-stream")


def dataframe_to_xlsx_bytes(df):
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


# ------------------------------------------------------------------------------
# Spremanje i učitavanje serija
# ------------------------------------------------------------------------------

def _safe_name(text, max_len=60):
    """Uklanja znakove koji smetaju u imenima datoteka (npr. kod preuzimanja na Windows)."""
    cleaned = "".join("_" if c in '\\/:*?"<>|' else c for c in text)
    return " ".join(cleaned.split())[:max_len].strip()


def new_batch_name(label="", now=None):
    now = now or datetime.now()
    computer = os.environ.get("COMPUTERNAME") or socket.gethostname() or "racunalo"
    name = f"{now:%Y-%m-%d_%H-%M-%S}_{computer}"
    label = _safe_name(label)
    return f"{name}_{label}" if label else name


def save_batch(service, form_type, results_df, scans=None, previous=None, label=""):
    """Sprema seriju anketa u Drive.

    results_df : tablica rezultata (sa stupcem "Datoteka")
    scans      : lista (ime_datoteke, bajtovi) ili None ako se skenovi ne spremaju
    previous   : rezultat ranijeg save_batch za ISTU seriju -> samo se zamijeni Excel
                 (npr. nakon ručnih ispravaka), skenovi se ne šalju ponovno.
    label      : upisani naziv (npr. studijski program) - dodaje se u ime serije

    Vraća rječnik s podacima o spremljenoj seriji.
    """
    if previous:
        xlsx = dataframe_to_xlsx_bytes(results_df)
        f = _replace_content(service, previous["results_file_id"], xlsx, XLSX_MIME)
        return {**previous, "results_link": f.get("webViewLink", previous.get("results_link")),
                "rows": len(results_df), "updated": True}

    batch = new_batch_name(label)
    results_folder = _ensure_path(service, ROOT_FOLDER, RESULTS_FOLDER, form_type)
    xlsx = dataframe_to_xlsx_bytes(results_df)
    results_file = _upload(service, f"{batch}.xlsx", xlsx, XLSX_MIME, results_folder["id"])

    info = {
        "batch": batch,
        "form_type": form_type,
        "rows": len(results_df),
        "results_file_id": results_file["id"],
        "results_link": results_file.get("webViewLink"),
        "scans_saved": 0,
        "scans_link": None,
        "updated": False,
    }

    if scans:
        scans_parent = _ensure_path(service, ROOT_FOLDER, SCANS_FOLDER, form_type)
        batch_folder = _ensure_folder(service, batch, scans_parent["id"])
        for name, data in scans:
            _upload(service, name, data, _guess_mime(name), batch_folder["id"])
        info["scans_saved"] = len(scans)
        info["scans_link"] = batch_folder.get("webViewLink")

    return info


def load_all_results(service, form_type):
    """Učitava sve spremljene serije jednog obrasca u jednu tablicu.
    Dodaje stupac "Serija" (iz imena Excel datoteke) na početak."""
    root = _find_folder(service, ROOT_FOLDER)
    if not root:
        return pd.DataFrame()
    results = _find_folder(service, RESULTS_FOLDER, root["id"])
    if not results:
        return pd.DataFrame()
    form_folder = _find_folder(service, form_type, results["id"])
    if not form_folder:
        return pd.DataFrame()

    files = _list(
        service,
        f"'{form_folder['id']}' in parents and mimeType = '{XLSX_MIME}' and trashed = false",
    )
    frames = []
    for f in files:
        data = service.files().get_media(fileId=f["id"]).execute()
        df = pd.read_excel(io.BytesIO(data))
        df.insert(0, "Serija", os.path.splitext(f["name"])[0])
        frames.append(df)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def root_folder_link(service):
    root = _find_folder(service, ROOT_FOLDER)
    return root.get("webViewLink") if root else None
