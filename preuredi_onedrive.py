"""
Preuređenje mapa u OneDriveu: svaka serija ide u podmapu po fakultetu, zatim studijskom programu
ili nastavniku.

    Rezultati/<obrazac>/<fakultet>/<studij ili nastavnik>/<serija>.xlsx
    Skenovi/<obrazac>/<fakultet>/<studij ili nastavnik>/<serija>/

Premješta i serije spremljene po datumu (prije 17.11.2026.) i serije u mapi sa zastarjelim nazivom
studija (npr. nakon ispravka u popisu kodova). Fakultet i studij čitaju se iz Excel tablice serije,
a naziv studija iz popisa kodova (programi.json). Ništa se ne briše osim mapa koje ostanu potpuno
prazne (idu u Koš za smeće OneDrivea). Sigurno ga je pokrenuti više puta.

Probni prikaz bez premještanja:   ... preuredi_onedrive.py --proba

Na poslužitelju:
    sudo -u omr OMR_DATA_DIR=/var/lib/omr-ankete OMR_SERVER=1 \\
        /opt/omr-ankete-venv/bin/python /opt/omr-ankete/preuredi_onedrive.py
Na računalu (u mapi programa):  python preuredi_onedrive.py
"""

import io
import os
import re
import sys

import pandas as pd

import onedrive_storage as od

BATCH_NAME = re.compile(r"^\d{4}-\d{2}-\d{2}_")
INFO_COLUMNS = ["Fakultet", "Kod studijskog programa", "Studijski program", "Nastavnik"]


def most_common(df, col):
    if col not in df.columns:
        return ""
    vals = df[col].dropna().astype(str).str.strip()
    vals = vals[(vals != "") & (vals.str.lower() != "nan")]
    return vals.mode().iloc[0] if len(vals) else ""


def scan_dirs(service, folder_id, path=()):
    """Mape serija sa skenovima (ime počinje datumom): {ime serije: (mapa, putanja podmapa)}."""
    out = {}
    for item in od._children(service, folder_id):
        if "folder" not in item:
            continue
        if BATCH_NAME.match(item["name"]):
            out[item["name"]] = (item, path)
        else:
            out.update(scan_dirs(service, item["id"], path + (item["name"],)))
    return out


def remove_empty(service, folder_id, dry):
    """Briše prazne podmape (nakon premještanja). Vraća True ako je i ova mapa ostala prazna."""
    empty = True
    for item in od._children(service, folder_id):
        if "folder" in item and remove_empty(service, item["id"], dry):
            if not dry:
                service.call("DELETE", f"/me/drive/items/{item['id']}", ok=(204, 404))
        else:
            empty = False
    return empty


def main():
    dry = "--proba" in sys.argv
    service = od.get_service()
    if service is None:
        sys.exit("Nema prijave u OneDrive. Najprije se prijavite u programu (bočna traka -> OneDrive).")
    print(f"OneDrive: {od.account_email(service)}" + ("   (PROBA - ništa se ne premješta)" if dry else ""))

    results_root = od._find_path(service, od.ROOT_FOLDER, od.RESULTS_FOLDER)
    if not results_root:
        sys.exit("U OneDriveu nema mape OMR Ankete/Rezultati.")
    scans_root = od._find_path(service, od.ROOT_FOLDER, od.SCANS_FOLDER)

    moved = in_place = problems = orphans = 0
    for form_folder in od._children(service, results_root["id"]):
        if "folder" not in form_folder or form_folder["name"] == od.ISVU_FOLDER:
            continue  # ISVU online tablice imaju vlastiti raspored (godina / semestar)
        form = form_folder["name"]
        scans_form = od._find_folder(service, form, scans_root["id"]) if scans_root else None
        scans = scan_dirs(service, scans_form["id"]) if scans_form else {}
        files = od._xlsx_files(service, form_folder["id"])
        print(f"\n== {form}: {len(files)} serija")
        for f, path in files:
            batch = os.path.splitext(f["name"])[0]
            scan = scans.pop(batch, None)
            try:
                df = pd.read_excel(io.BytesIO(od._download(service, f["id"])), dtype=str)
                folders = od.batch_folders(form, {c: most_common(df, c) for c in INFO_COLUMNS})
                scan_ok = scan is None or list(scan[1]) == folders
                if list(path) == folders and scan_ok:
                    in_place += 1
                    continue
                print(f"  {batch}  ->  {' / '.join(folders)}" + ("" if scan else "   (bez skenova)"))
                moved += 1
                if dry:
                    continue
                if list(path) != folders:
                    dest = od._ensure_path(service, od.ROOT_FOLDER, od.RESULTS_FOLDER, form, *folders)
                    if not od.move_item(service, f["id"], dest["id"]):
                        print("     ! rezultati: u odredištu već postoji datoteka istog imena - ostavljeno")
                        problems += 1
                if not scan_ok:
                    dest = od._ensure_path(service, od.ROOT_FOLDER, od.SCANS_FOLDER, form, *folders)
                    if not od.move_item(service, scan[0]["id"], dest["id"]):
                        print("     ! skenovi: u odredištu već postoji mapa istog imena - ostavljeno")
                        problems += 1
            except Exception as e:
                print(f"     ! nije premješteno: {e}")
                problems += 1
        for name in scans:
            print(f"  ? skenovi bez tablice rezultata (ostavljeno na mjestu): {name}")
            orphans += 1
        for root in [form_folder] + ([scans_form] if scans_form else []):
            remove_empty(service, root["id"], dry)

    print(f"\nGotovo. Premješteno: {moved}, već na pravom mjestu: {in_place}, upozorenja: {problems}, "
          f"skenova bez tablice: {orphans}.")
    link = od.root_folder_link(service)
    if link:
        print(f"Mapa u OneDriveu: {link}")


if __name__ == "__main__":
    main()
