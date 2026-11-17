"""
Jednokratno preuređenje mapa u OneDriveu: serije spremljene po datumu premješta u podmape
po fakultetu, zatim studijskom programu ili nastavniku.

    Rezultati/<obrazac>/<serija>.xlsx          ->  Rezultati/<obrazac>/<fakultet>/<studij ili nastavnik>/<serija>.xlsx
    Skenovi/<obrazac>/<serija>/                ->  Skenovi/<obrazac>/<fakultet>/<studij ili nastavnik>/<serija>/

Fakultet i studij (nastavnik) čitaju se iz same Excel tablice serije. Ništa se ne briše, samo premješta.
Sigurno ga je pokrenuti više puta (već preuređene serije se preskaču).

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

INFO_COLUMNS = ["Fakultet", "Kod studijskog programa", "Studijski program", "Nastavnik"]


def most_common(df, col):
    if col not in df.columns:
        return ""
    vals = df[col].dropna().astype(str).str.strip()
    vals = vals[(vals != "") & (vals.str.lower() != "nan")]
    return vals.mode().iloc[0] if len(vals) else ""


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

    moved = skipped = problems = 0
    for form_folder in od._children(service, results_root["id"]):
        if "folder" not in form_folder:
            continue
        form = form_folder["name"]
        scans_form = od._find_folder(service, form, scans_root["id"]) if scans_root else None
        scan_dirs = ({c["name"]: c for c in od._children(service, scans_form["id"]) if "folder" in c}
                     if scans_form else {})
        files = [f for f in od._children(service, form_folder["id"])
                 if "folder" not in f and f["name"].lower().endswith(".xlsx")]
        print(f"\n== {form}: {len(files)} serija za preuređenje")
        for f in files:
            batch = os.path.splitext(f["name"])[0]
            try:
                df = pd.read_excel(io.BytesIO(od._download(service, f["id"])), dtype=str)
                folders = od.batch_folders(form, {c: most_common(df, c) for c in INFO_COLUMNS})
                scan_dir = scan_dirs.pop(batch, None)
                print(f"  {batch}  ->  {' / '.join(folders)}" + ("" if scan_dir else "   (bez skenova)"))
                if dry:
                    moved += 1
                    continue
                dest = od._ensure_path(service, od.ROOT_FOLDER, od.RESULTS_FOLDER, form, *folders)
                if not od.move_item(service, f["id"], dest["id"]):
                    print("     ! rezultati: u odredištu već postoji datoteka istog imena - ostavljeno")
                    problems += 1
                    continue
                if scan_dir:
                    dest = od._ensure_path(service, od.ROOT_FOLDER, od.SCANS_FOLDER, form, *folders)
                    if not od.move_item(service, scan_dir["id"], dest["id"]):
                        print("     ! skenovi: u odredištu već postoji mapa istog imena - ostavljeno")
                        problems += 1
                moved += 1
            except Exception as e:
                print(f"     ! nije premješteno: {e}")
                problems += 1

        # Skenovi koji su ostali po datumu: smjeste se uz svoju tablicu rezultata (ako je već preuređena)
        leftovers = [n for n in scan_dirs if re.match(r"^\d{4}-\d{2}-\d{2}_", n)]
        if leftovers:
            placed = {os.path.splitext(x["name"])[0]: path for x, path in od._xlsx_files(service, form_folder["id"])}
            for name in leftovers:
                path = placed.get(name)
                if not path:
                    print(f"  ? skenovi bez tablice rezultata (ostavljeno na mjestu): {name}")
                    skipped += 1
                    continue
                print(f"  {name} (skenovi)  ->  {' / '.join(path)}")
                if not dry:
                    dest = od._ensure_path(service, od.ROOT_FOLDER, od.SCANS_FOLDER, form, *path)
                    od.move_item(service, scan_dirs[name]["id"], dest["id"])

    print(f"\nGotovo. Premješteno serija: {moved}, upozorenja: {problems}, skenova bez tablice: {skipped}.")
    link = od.root_folder_link(service)
    if link:
        print(f"Mapa u OneDriveu: {link}")


if __name__ == "__main__":
    main()
