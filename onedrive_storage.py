"""
Spremanje rezultata, skenova i izvještaja u poslovni OneDrive (mvinscak@unizg.hr).

Struktura mapa (po fakultetu, zatim studijskom programu ili nastavniku):

    OMR Ankete/
        Rezultati/<vrsta obrasca>/<fakultet>/<studij ili nastavnik>/<serija>.xlsx
        Skenovi/<vrsta obrasca>/<fakultet>/<studij ili nastavnik>/<serija>/<izvorne slike>
        Izvještaji/<vrsta obrasca>/<fakultet>/<izvještaj>.pdf
        Za fakultete/<fakultet>/<ak. god.>/<vrsta>/...   (mapa koja se šalje fakultetu)

Prijava (jednom, zatim se sama obnavlja):
    Program prikaže kratki kod -> na https://microsoft.com/devicelogin upišete kod i prijavite se
    računom @unizg.hr. Lozinka nikad ne prolazi kroz program. Prijava se sprema u onedrive_token.json
    (NIKADA na GitHub) i vrijedi dok se program koristi barem jednom u 90 dana.

Aplikacija "OMR Ankete" registrirana je u Microsoft Entri Sveučilišta u Zagrebu (samo za taj
tenant). ID aplikacije i tenanta nisu tajni: bez prijave korisnika s njima se ne može ništa.
Ovlasti: Files.ReadWrite (vlastiti OneDrive), User.Read, offline_access.
"""

import base64
import io
import json
import os
import socket
import time
from datetime import datetime
from urllib.parse import quote

import pandas as pd

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.environ.get("OMR_DATA_DIR") or APP_DIR
TOKEN_FILE = os.path.join(DATA_DIR, "onedrive_token.json")
SERVER_MODE = os.environ.get("OMR_SERVER") == "1"

CLIENT_ID = os.environ.get("OMR_MS_CLIENT_ID", "81b53565-04b2-4d78-9d51-a9b30fca601f")
TENANT_ID = os.environ.get("OMR_MS_TENANT_ID", "df0ba543-57f4-4b1b-8a1c-31687c1d9f58")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0"
SCOPES = "Files.ReadWrite User.Read offline_access"
GRAPH = "https://graph.microsoft.com/v1.0"

SERVICE_NAME = "OneDrive"
ROOT_FOLDER = "OMR Ankete"
RESULTS_FOLDER = "Rezultati"
SCANS_FOLDER = "Skenovi"
REPORTS_FOLDER = "Izvještaji"
DELIVERY_FOLDER = "Za fakultete"
SCANS_SUBFOLDER = "Skenirane ankete"

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024
CHUNK = 320 * 1024 * 16  # 5 MB, mora biti višekratnik od 320 KiB


class DriveNotConfigured(Exception):
    """Nedostaje biblioteka requests."""


class LoginError(Exception):
    """Microsoft je odbio prijavu (poruka je za korisnika)."""


# ------------------------------------------------------------------------------
# Prijava (device code: kod se upiše na microsoft.com/devicelogin)
# ------------------------------------------------------------------------------

def libraries_installed():
    try:
        import requests  # noqa: F401
        return True
    except ImportError:
        return False


def is_configured():
    return libraries_installed()


def has_saved_login():
    return os.path.exists(TOKEN_FILE)


def _write_token(token):
    os.makedirs(DATA_DIR, exist_ok=True)
    tmp = TOKEN_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(token, f)
    try:
        os.chmod(tmp, 0o600)
    except OSError:
        pass
    os.replace(tmp, TOKEN_FILE)


def _read_token():
    try:
        with open(TOKEN_FILE, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return None


def _token_from_response(data, old_refresh=None):
    return {
        "access_token": data["access_token"],
        "refresh_token": data.get("refresh_token") or old_refresh,
        "expires_at": time.time() + int(data.get("expires_in", 3600)),
    }


def _ms_error(data):
    """Razumljiva poruka iz Microsoftove greške (AADSTS...)."""
    desc = (data.get("error_description") or data.get("error") or "nepoznata greška").split("\r\n")[0]
    if "AADSTS65001" in desc or "consent" in desc.lower():
        return ("Aplikaciju OMR Ankete treba odobriti administrator Sveučilišta (admin consent). " + desc)
    if "AADSTS50199" in desc or "AADSTS530" in desc:
        return "Pravila Sveučilišta ne dopuštaju ovaj način prijave. " + desc
    return desc


def start_device_login():
    """Pokreće prijavu: vraća rječnik s kodom (user_code) i adresom (verification_uri)."""
    import requests
    r = requests.post(f"{AUTHORITY}/devicecode", data={"client_id": CLIENT_ID, "scope": SCOPES}, timeout=30)
    data = r.json()
    if r.status_code != 200:
        raise LoginError(_ms_error(data))
    return {
        "device_code": data["device_code"],
        "user_code": data["user_code"],
        "verification_uri": data.get("verification_uri", "https://microsoft.com/devicelogin"),
        "interval": int(data.get("interval", 5)),
        "expires_at": time.time() + int(data.get("expires_in", 900)),
    }


def poll_device_login(flow, wait=60):
    """Čeka da korisnik upiše kod (najviše `wait` sekundi).
    Vraća servis kad je prijava gotova, None ako još nije. Baca LoginError ako je odbijena/istekla."""
    import requests
    deadline = min(time.time() + wait, flow["expires_at"])
    while True:
        r = requests.post(f"{AUTHORITY}/token", data={
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": CLIENT_ID,
            "device_code": flow["device_code"],
        }, timeout=30)
        data = r.json()
        if r.status_code == 200:
            _write_token(_token_from_response(data))
            return OneDriveService(_read_token())
        err = data.get("error")
        if err == "authorization_pending":
            pass
        elif err == "slow_down":
            flow["interval"] += 5
        elif err == "expired_token":
            raise LoginError("Kod je istekao. Pokrenite prijavu ponovno.")
        elif err == "authorization_declined":
            raise LoginError("Prijava je odbijena.")
        else:
            raise LoginError(_ms_error(data))
        if time.time() + flow["interval"] > deadline:
            return None
        time.sleep(flow["interval"])


def get_service(interactive=False):
    """Vraća OneDrive servis ako postoji spremljena (valjana) prijava, inače None.
    (interactive se ne koristi - prijava ide preko start_device_login/poll_device_login.)"""
    if not libraries_installed():
        raise DriveNotConfigured("Nedostaje biblioteka requests. Pokrenite: pip install -r requirements.txt")
    token = _read_token()
    if not token or not token.get("refresh_token"):
        return None
    service = OneDriveService(token)
    try:
        service.access_token()  # Obnovi ako je istekla; baca LoginError ako je prijava opozvana
    except LoginError:
        return None
    return service


def logout():
    if os.path.exists(TOKEN_FILE):
        os.remove(TOKEN_FILE)


_IDS_BY_URL = {}


class OneDriveService:
    """Mali klijent za Microsoft Graph: prijava se sama obnavlja, ponavlja kod zagušenja (429/503)."""

    def __init__(self, token):
        import requests
        self.token = token
        self.http = requests.Session()
        self._drive_id = None
        self.ids_by_url = _IDS_BY_URL  # webUrl -> id (za dijeljenje bez dodatnog upita)

    def access_token(self):
        if self.token.get("expires_at", 0) - 120 > time.time():
            return self.token["access_token"]
        # Druga sesija (drugo računalo) možda je već obnovila prijavu
        disk = _read_token()
        if disk and disk.get("expires_at", 0) - 120 > time.time():
            self.token = disk
            return disk["access_token"]
        r = self.http.post(f"{AUTHORITY}/token", data={
            "grant_type": "refresh_token",
            "client_id": CLIENT_ID,
            "refresh_token": (disk or self.token)["refresh_token"],
            "scope": SCOPES,
        }, timeout=30)
        data = r.json()
        if r.status_code != 200:
            raise LoginError(_ms_error(data))
        self.token = _token_from_response(data, (disk or self.token)["refresh_token"])
        _write_token(self.token)
        return self.token["access_token"]

    def call(self, method, url, ok=(200, 201, 202, 204), allow404=False, auth=True, **kw):
        if not url.startswith("http"):
            url = GRAPH + url
        for attempt in range(6):
            headers = dict(kw.pop("headers", None) or {})
            if auth:
                headers["Authorization"] = "Bearer " + self.access_token()
            r = self.http.request(method, url, headers=headers, timeout=120, **kw)
            kw["headers"] = {k: v for k, v in headers.items() if k != "Authorization"}
            if r.status_code in (429, 503, 504) and attempt < 5:
                time.sleep(min(int(r.headers.get("Retry-After", 2 ** attempt)), 60))
                continue
            if r.status_code == 401 and auth and attempt == 0:
                self.token["expires_at"] = 0  # Prisilno obnovi prijavu i pokušaj još jednom
                continue
            break
        if allow404 and r.status_code == 404:
            return None
        if r.status_code not in ok:
            try:
                err = r.json().get("error", {})
                msg = f"{err.get('code')}: {err.get('message')}"
            except ValueError:
                msg = r.text[:300]
            raise RuntimeError(f"OneDrive ({r.status_code}) {msg}")
        return r

    def json(self, method, url, **kw):
        r = self.call(method, url, **kw)
        if r is None:
            return None
        item = r.json() if r.content else {}
        if isinstance(item, dict) and item.get("webUrl") and item.get("id"):
            self.ids_by_url[item["webUrl"]] = item["id"]
        return item

    @property
    def drive_id(self):
        if not self._drive_id:
            self._drive_id = self.json("GET", "/me/drive?$select=id")["id"]
        return self._drive_id


def account_email(service):
    try:
        me = service.json("GET", "/me?$select=mail,userPrincipalName")
        return me.get("mail") or me.get("userPrincipalName")
    except Exception:
        return None


# ------------------------------------------------------------------------------
# Mape i datoteke
# ------------------------------------------------------------------------------

def _q(name):
    return quote(name, safe="")


def _child(service, parent_id, name):
    """Stavka imena `name` u mapi parent_id ("root" = vrh OneDrivea) ili None."""
    return service.json("GET", f"/me/drive/items/{parent_id}:/{_q(name)}", allow404=True)


def _find_folder(service, name, parent_id=None):
    item = _child(service, parent_id or "root", name)
    return item if item and "folder" in item else None


def _ensure_folder(service, name, parent_id=None):
    parent_id = parent_id or "root"
    item = _find_folder(service, name, parent_id)
    if item:
        return item
    r = service.call("POST", f"/me/drive/items/{parent_id}/children", ok=(200, 201, 409), json={
        "name": name, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"})
    if r.status_code == 409:  # Netko ju je upravo stvorio (drugo računalo)
        return _find_folder(service, name, parent_id)
    item = r.json()
    service.ids_by_url[item.get("webUrl")] = item["id"]
    return item


def _ensure_path(service, *names):
    parent = None
    for name in names:
        parent = _ensure_folder(service, name, parent["id"] if parent else None)
    return parent


def _find_path(service, *names):
    """Kao _ensure_path, ali ništa ne stvara: vraća mapu ili None."""
    parent = None
    for name in names:
        parent = _find_folder(service, name, parent["id"] if parent else None)
        if not parent:
            return None
    return parent


def _children(service, folder_id):
    url = (f"/me/drive/items/{folder_id}/children"
           "?$select=id,name,file,folder,webUrl,createdDateTime,size&$top=999")
    items = []
    while url:
        page = service.json("GET", url)
        items.extend(page.get("value", []))
        url = page.get("@odata.nextLink")
    for it in items:
        service.ids_by_url[it.get("webUrl")] = it["id"]
    return sorted(items, key=lambda it: it.get("createdDateTime", ""))


def _files_in(service, folder_id):
    return [{"id": f["id"], "name": f["name"]} for f in _children(service, folder_id) if "folder" not in f]


def _upload_session(service, session_url, data, conflict):
    """Velike datoteke (> 4 MB) šalju se u dijelovima."""
    upload_url = service.json("POST", session_url, json={
        "item": {"@microsoft.graph.conflictBehavior": conflict}})["uploadUrl"]
    total, start, last = len(data), 0, None
    while start < total:
        part = data[start:start + CHUNK]
        end = start + len(part) - 1
        last = service.call("PUT", upload_url, auth=False, data=part, headers={
            "Content-Length": str(len(part)), "Content-Range": f"bytes {start}-{end}/{total}"})
        start = end + 1
    item = last.json()
    service.ids_by_url[item.get("webUrl")] = item["id"]
    return item


def _upload(service, name, data, parent_id, conflict="rename"):
    if len(data) > SIMPLE_UPLOAD_LIMIT:
        return _upload_session(service, f"/me/drive/items/{parent_id}:/{_q(name)}:/createUploadSession",
                               data, conflict)
    return service.json("PUT", f"/me/drive/items/{parent_id}:/{_q(name)}:/content"
                               f"?@microsoft.graph.conflictBehavior={conflict}", data=data)


def _replace_content(service, file_id, data):
    if len(data) > SIMPLE_UPLOAD_LIMIT:
        return _upload_session(service, f"/me/drive/items/{file_id}/createUploadSession", data, "replace")
    return service.json("PUT", f"/me/drive/items/{file_id}/content", data=data)


def _download(service, file_id):
    return service.call("GET", f"/me/drive/items/{file_id}/content").content


def dataframe_to_xlsx_bytes(df):
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


# ------------------------------------------------------------------------------
# Spremanje i učitavanje serija
# ------------------------------------------------------------------------------

def _safe_name(text, max_len=60):
    """Uklanja znakove koji smetaju u imenima datoteka (Windows/OneDrive)."""
    cleaned = "".join("_" if c in '\\/:*?"<>|#%' else c for c in text)
    return " ".join(cleaned.split())[:max_len].strip().rstrip(".")


def new_batch_name(label="", now=None):
    now = now or datetime.now()
    computer = os.environ.get("COMPUTERNAME") or socket.gethostname() or "racunalo"
    name = f"{now:%Y-%m-%d_%H-%M-%S}_{_safe_name(computer)}"
    label = _safe_name(label)
    return f"{name}_{label}" if label else name


def folder_part(text, max_len=70):
    """Naziv mape (isto pravilo kao reports.safe_file_part, da se mape fakulteta podudaraju)."""
    cleaned = "".join("_" if c in '\\/:*?"<>|' else c for c in str(text))
    return " ".join(cleaned.split())[:max_len].strip().rstrip(".").strip()


def _text(value):
    t = "" if value is None else str(value).strip()
    return "" if t.lower() in ("nan", "none") else t


_PROGRAMS = None


def _programs():
    global _PROGRAMS
    if _PROGRAMS is None:
        import programi
        _PROGRAMS = programi.load()
    return _PROGRAMS


def batch_folders(form_type, values):
    """[fakultet, studijski program ili nastavnik] - podmape u Rezultati/ i Skenovi/.
    values: upisani podaci serije (Fakultet, Kod studijskog programa, Studijski program, Nastavnik)."""
    faculty = folder_part(_text(values.get("Fakultet"))) or "Bez fakulteta"
    if "nastavnik" in form_type.lower():
        group = folder_part(_text(values.get("Nastavnik"))) or "Bez nastavnika"
    else:
        import programi
        code = programi.normalize_code(values.get("Kod studijskog programa"))  # 190 / "190.0" -> "0190"
        entry = programi.lookup(_programs(), code, form_type) if code else None
        name = entry["program"] if entry else _text(values.get("Studijski program"))
        name = name.split(";")[0].strip()  # Mapa bez popisa smjerova: "0190 Inženjerstvo okoliša"
        group = folder_part(" ".join(p for p in [code, name] if p))
        group = group or "Bez studijskog programa"
    return [faculty, group]


def save_batch(service, form_type, results_df, scans=None, previous=None, label="", folders=()):
    """Sprema seriju anketa u OneDrive.
    folders: podmape ispod vrste obrasca, npr. batch_folders(...) = [fakultet, studij ili nastavnik]."""
    if previous:
        f = _replace_content(service, previous["results_file_id"], dataframe_to_xlsx_bytes(results_df))
        return {**previous, "results_link": f.get("webUrl", previous.get("results_link")),
                "rows": len(results_df), "updated": True}

    batch = new_batch_name(label)
    folders = [f for f in folders if f]
    results_folder = _ensure_path(service, ROOT_FOLDER, RESULTS_FOLDER, form_type, *folders)
    results_file = _upload(service, f"{batch}.xlsx", dataframe_to_xlsx_bytes(results_df), results_folder["id"])

    info = {
        "batch": batch,
        "form_type": form_type,
        "rows": len(results_df),
        "results_file_id": results_file["id"],
        "results_link": results_file.get("webUrl"),
        "results_folder_link": results_folder.get("webUrl"),
        "scans_saved": 0,
        "scans_link": None,
        "updated": False,
    }
    if scans:
        scans_parent = _ensure_path(service, ROOT_FOLDER, SCANS_FOLDER, form_type, *folders)
        batch_folder = _ensure_folder(service, batch, scans_parent["id"])
        for name, data in scans:
            _upload(service, _safe_name(name, 120) or "sken", data, batch_folder["id"])
        info["scans_saved"] = len(scans)
        info["scans_link"] = batch_folder.get("webUrl")
    return info


MAP_COLUMN = "Mapa"  # fakultet/studij - gdje je serija spremljena (za pronalazak skenova)


def _xlsx_files(service, folder_id, path=()):
    """Sve .xlsx datoteke u mapi i njezinim podmapama: (datoteka, putanja podmapa)."""
    out = []
    for item in _children(service, folder_id):
        if "folder" in item:
            out += _xlsx_files(service, item["id"], path + (item["name"],))
        elif item["name"].lower().endswith(".xlsx"):
            out.append((item, path))
    return out


def load_all_results(service, form_type):
    """Sve spremljene serije jednog obrasca u jednoj tablici; stupci "Serija" i "Mapa" na početku."""
    folder = _find_path(service, ROOT_FOLDER, RESULTS_FOLDER, form_type)
    if not folder:
        return pd.DataFrame()
    frames = []
    for f, path in _xlsx_files(service, folder["id"]):
        df = pd.read_excel(io.BytesIO(_download(service, f["id"])))
        df.insert(0, "Serija", os.path.splitext(f["name"])[0])
        df.insert(1, MAP_COLUMN, "/".join(path))
        frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def move_item(service, item_id, dest_id):
    """Premješta datoteku ili mapu (OneDrive to radi odmah). Vraća False ako u odredištu već postoji isto ime."""
    r = service.call("PATCH", f"/me/drive/items/{item_id}", ok=(200, 409), json={
        "parentReference": {"id": dest_id}, "@microsoft.graph.conflictBehavior": "fail"})
    return r.status_code == 200


def save_report(service, form_type, folder_name, file_name, data):
    """Sprema PDF u OMR Ankete/Izvještaji/<obrazac>/<folder_name>/ (isto ime -> zamjena sadržaja).
    Vraća (poveznica na datoteku, poveznica na mapu) - poveznice vide samo vi dok se ne podijele."""
    folder = _ensure_path(service, ROOT_FOLDER, REPORTS_FOLDER, form_type, folder_name)
    f = _upload(service, file_name, data, folder["id"], conflict="replace")
    return f.get("webUrl"), folder.get("webUrl")


def share_with_link(service, link):
    """Poveznica "Svatko s vezom može gledati" za datoteku ili mapu (link = njezin webUrl).
    Vraća poveznicu za dijeljenje - nju treba staviti u e-mail (razlikuje se od webUrl)."""
    item_id = service.ids_by_url.get(link)
    if not item_id:
        encoded = base64.urlsafe_b64encode(link.encode("utf-8")).decode("ascii").rstrip("=")
        item_id = service.json("GET", f"/shares/u!{encoded}/driveItem?$select=id")["id"]
    r = service.call("POST", f"/me/drive/items/{item_id}/createLink", ok=(200, 201, 403, 400),
                     json={"type": "view", "scope": "anonymous"})
    if r.status_code in (400, 403):
        raise RuntimeError("OneDrive ne dopušta poveznicu 'Svatko s vezom' za ovu stavku: "
                           + r.json().get("error", {}).get("message", ""))
    return r.json()["link"]["webUrl"]


# ------------------------------------------------------------------------------
# Mapa za fakultet: PDF izvještaji + skenirane ankete na jednom mjestu (bez Excela)
# ------------------------------------------------------------------------------

def list_reports(service, form_type, folder_name):
    """PDF izvještaji jednog fakulteta: Izvještaji/<obrazac>/<fakultet>/*.pdf"""
    folder = _find_path(service, ROOT_FOLDER, REPORTS_FOLDER, form_type, folder_name)
    if not folder:
        return []
    return [f for f in _files_in(service, folder["id"]) if f["name"].lower().endswith(".pdf")]


def scan_folder_id(service, form_type, batch, map_path=""):
    """Mapa sa skenovima serije: Skenovi/<obrazac>/<fakultet>/<studij>/<serija> (ili stari raspored bez podmapa)."""
    sub = [p for p in str(map_path or "").split("/") if p and p.lower() != "nan"]
    for path in ([sub, []] if sub else [[]]):
        folder = _find_path(service, ROOT_FOLDER, SCANS_FOLDER, form_type, *path, batch)
        if folder:
            return folder["id"]
    return None


def _start_copies(service, source_files, dest_id):
    """Pokreće kopiranje (OneDrive kopira sam, u pozadini) datoteka koje u odredištu još ne postoje.
    Vraća popis adresa za praćenje."""
    have = {f["name"] for f in _files_in(service, dest_id)}
    monitors = []
    for f in source_files:
        if f["name"] in have:
            continue
        r = service.call("POST", f"/me/drive/items/{f['id']}/copy?@microsoft.graph.conflictBehavior=fail",
                         json={"parentReference": {"driveId": service.drive_id, "id": dest_id}, "name": f["name"]})
        have.add(f["name"])
        if r.headers.get("Location"):
            monitors.append(r.headers["Location"])
    return monitors


def _wait_for_copies(service, monitors, progress=None, timeout=900):
    """Čeka da OneDrive dovrši kopiranje, da fakultet ne dobije poluprazne mape."""
    pending, failed = list(monitors), []
    total, deadline = len(pending) or 1, time.time() + timeout
    while pending and time.time() < deadline:
        still = []
        for url in pending:
            r = service.call("GET", url, auth=False, ok=(200, 202, 303), allow_redirects=False)
            body = r.json() if r.content and r.status_code != 303 else {}
            status = "completed" if r.status_code == 303 else body.get("status", "")
            if status == "failed":
                code = json.dumps(body.get("error") or body).lower()
                if "namealreadyexists" not in code:  # Već kopirano ranije -> u redu
                    failed.append(url)
            elif status not in ("completed",):
                still.append(url)
        pending = still
        if progress:
            progress(1 - len(pending) / total, f"OneDrive kopira datoteke: još {len(pending)}")
        if pending:
            time.sleep(3)
    return len(pending), len(failed)


def prepare_delivery(service, faculty_folder, year_folder, parts, progress=None):
    """Složi mapu za fakultet i podijeli je (svatko s poveznicom može gledati).

        OMR Ankete/Za fakultete/<fakultet>/<ak. god.>/<vrsta>/<izvještaji>.pdf
                                                     /<vrsta>/Skenirane ankete/<studij ili nastavnik>/<slike>

    parts : popis (naziv vrste, [PDF datoteke {id, name}], [(naziv podmape, id mape sa skenovima)])
    Vraća (poveznica za dijeljenje, broj kopiranih izvještaja, broj kopiranih skenova)."""
    dest = _ensure_path(service, ROOT_FOLDER, DELIVERY_FOLDER, faculty_folder, year_folder)
    monitors, n_reports, n_scans = [], 0, 0
    steps = sum(1 + len(scans) for _, _, scans in parts) or 1
    done = 0
    for kind_name, report_files, scans in parts:
        kind_folder = _ensure_folder(service, kind_name, dest["id"])
        started = _start_copies(service, report_files, kind_folder["id"])
        monitors += started
        n_reports += len(started)
        done += 1
        if progress:
            progress(0.5 * done / steps, f"{kind_name}: izvještaji")
        if scans:
            scans_root = _ensure_folder(service, SCANS_SUBFOLDER, kind_folder["id"])
            for label, source_id in scans:
                target = _ensure_folder(service, _safe_name(label, 120), scans_root["id"])
                started = _start_copies(service, _files_in(service, source_id), target["id"])
                monitors += started
                n_scans += len(started)
                done += 1
                if progress:
                    progress(0.5 * done / steps, f"{kind_name}: skenovi {label}")
    left, failed = _wait_for_copies(
        service, monitors, (lambda p, t: progress(0.5 + 0.5 * p, t)) if progress else None)
    if failed:
        raise RuntimeError(f"OneDrive nije kopirao {failed} datoteka. Pokrenite pripremu mape ponovno.")
    if left:
        raise RuntimeError(f"OneDrive još kopira {left} datoteka. Pričekajte minutu i pokrenite ponovno.")
    return share_with_link(service, dest["webUrl"]), n_reports, n_scans


def read_json(service, name):
    """JSON datoteka u mapi OMR Ankete (npr. poslano.json) ili None ako je još nema."""
    root = _find_folder(service, ROOT_FOLDER)
    item = _child(service, root["id"], name) if root else None
    if not item or "folder" in item:
        return None
    return json.loads(_download(service, item["id"]).decode("utf-8"))


def write_json(service, name, data):
    root = _ensure_folder(service, ROOT_FOLDER)
    body = json.dumps(data, ensure_ascii=False, indent=1).encode("utf-8")
    _upload(service, name, body, root["id"], conflict="replace")


def root_folder_link(service):
    root = _find_folder(service, ROOT_FOLDER)
    return root.get("webUrl") if root else None
