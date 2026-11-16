"""
Jednokratni prijenos: Google Drive "OMR Ankete" -> OneDrive "OMR Ankete" (ista struktura mapa).

Preduvjeti: u programu je ranije bila prijava u Google Drive (google_token.json) i sada je
napravljena prijava u OneDrive (onedrive_token.json) - obje u istoj mapi podataka.

Na poslužitelju:
    sudo -u omr OMR_DATA_DIR=/var/lib/omr-ankete OMR_SERVER=1 \\
        /opt/omr-ankete-venv/bin/python /opt/omr-ankete/prenesi_iz_google_drivea.py

Sigurno ga je pokrenuti više puta: datoteke koje već postoje u OneDriveu se preskaču.
Google Drive ostaje netaknut (ništa se ne briše).
"""

import sys

import drive_storage as gdrive
import onedrive_storage as onedrive


def google_children(gservice, folder_id):
    return gdrive._list(gservice, f"'{folder_id}' in parents and trashed = false",
                        fields="id, name, mimeType, size")


def copy_tree(gservice, oservice, g_folder, o_folder, path, stats):
    existing = {c["name"] for c in onedrive._children(oservice, o_folder["id"])}
    seen = set()
    for item in google_children(gservice, g_folder["id"]):
        name = item["name"]
        here = f"{path}/{name}"
        if item["mimeType"] == gdrive.FOLDER_MIME:
            sub = onedrive._ensure_folder(oservice, name, o_folder["id"])
            copy_tree(gservice, oservice, item, sub, here, stats)
            continue
        if item["mimeType"].startswith("application/vnd.google-apps"):
            print(f"  preskočeno (Googleov dokument, ne datoteka): {here}")
            stats["skipped"] += 1
            continue
        if name in existing or name in seen:
            stats["existing"] += 1
            continue
        data = gservice.files().get_media(fileId=item["id"]).execute()
        onedrive._upload(oservice, name, data, o_folder["id"], conflict="fail")
        seen.add(name)
        stats["copied"] += 1
        stats["bytes"] += len(data)
        print(f"  ✓ {here}  ({len(data) // 1024} KB)")


def main():
    gservice = gdrive.get_service(interactive=False)
    if gservice is None:
        sys.exit("Nema prijave u Google Drive (google_token.json). Prijenos nije moguć.")
    oservice = onedrive.get_service()
    if oservice is None:
        sys.exit("Nema prijave u OneDrive. Najprije se prijavite u programu (bočna traka -> OneDrive).")

    print(f"Google:   {gdrive.account_email(gservice)}")
    print(f"OneDrive: {onedrive.account_email(oservice)}")
    g_root = gdrive._find_folder(gservice, gdrive.ROOT_FOLDER)
    if not g_root:
        sys.exit(f"U Google Driveu nema mape {gdrive.ROOT_FOLDER}.")
    o_root = onedrive._ensure_folder(oservice, onedrive.ROOT_FOLDER)

    stats = {"copied": 0, "existing": 0, "skipped": 0, "bytes": 0}
    copy_tree(gservice, oservice, g_root, o_root, gdrive.ROOT_FOLDER, stats)
    print()
    print(f"Gotovo. Preneseno {stats['copied']} datoteka ({stats['bytes'] / 1024 / 1024:.1f} MB), "
          f"već postojalo {stats['existing']}, preskočeno {stats['skipped']}.")
    print(f"Mapa u OneDriveu: {o_root.get('webUrl')}")


if __name__ == "__main__":
    main()
