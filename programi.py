"""
Kodovi studijskih programa (4 znamenke koje studenti zacrne na 1. stranici obrasca).

Popis se čita iz programi.json (izrađen iz tablice "Završne kodovi").
Kad dobijete novu tablicu kodova, spremite je kao .xlsx i pokrenite:

    python programi.py "Završne kodovi 2026-27.xlsx"

To prepisuje programi.json - aplikacija ga koristi nakon ponovnog pokretanja.
"""

import json
import os
import re
import sys

PROGRAMS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "programi.json")

PREDDIPLOMSKI, DIPLOMSKI, INTEGRIRANI = "Preddiplomski studij", "Diplomski studij", "Integrirani studij"

# Stari nazivi sastavnica u tablici -> današnji nazivi
RENAME = {
    "Šumarski fakultet": "Fakultet šumarstva i drvne tehnologije",
    "Hrvatski studiji": "Fakultet hrvatskih studija",
}
# Riječi koje opisuju vrstu studija - skidaju se s početka naziva da ostane ime programa
LEVEL_WORDS = {
    "integrirani", "preddiplomski", "prijediplomski", "prediplomski", "diplomski", "sveučilišni",
    "sveučiilšni", "stručni", "specijalistički", "izvanredni", "dislocirani", "združeni",
    "međunarodni", "i", "studij", "j",
}


# ------------------------------------------------------------------------------
# Korištenje u aplikaciji
# ------------------------------------------------------------------------------

def load(path=PROGRAMS_FILE):
    """Vraća {kod: [zapis, ...]}. Isti kod se u tablici rijetko pojavljuje dvaput (npr. 0370)."""
    if not os.path.exists(path):
        return {}
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def lookup(programs, code, form_type=None):
    """Zapis za kod; ako kod ima više zapisa, prednost ima onaj iste razine kao obrazac."""
    entries = programs.get(str(code).strip().zfill(4), []) if code else []
    for e in entries:
        if form_type and e["razina"] == form_type:
            return e
    return entries[0] if entries else None


def choices(programs, form_type=None):
    """Popis za odabir u aplikaciji: programi razine obrasca prvi, zatim ostali."""
    same, other = [], []
    for code, entries in programs.items():
        for e in entries:
            (same if e["razina"] == form_type else other).append((code, e))
    return same + other


def label(code, entry):
    return f"{code} – {entry['program']} ({entry['fakultet']})"


# ------------------------------------------------------------------------------
# Izrada programi.json iz tablice kodova
# ------------------------------------------------------------------------------

def _clean(text):
    return " ".join(str(text).replace("*", " ").split())


def _faculty_name(text):
    t = re.sub(r"\s*-\s*", "-", _clean(text))
    t = re.sub(r"^\d{1,2}\.\s*\d{1,2}\.\s*\d{4}\.\s*", "", t)  # "6.7.2023. Katolički bogoslovni fakultet"
    t = re.sub(r",\s*(Varaždin|Sisak)$", "", t)
    return RENAME.get(t, t)


def _level(full, section):
    f = full.lower()
    if "integrirani" in f:
        return INTEGRIRANI
    if re.search(r"\b(pred|prije|pre)diplomski", f):
        return PREDDIPLOMSKI
    if re.search(r"\bdiplomski", f):
        return DIPLOMSKI
    if f.startswith("stručni"):
        return PREDDIPLOMSKI  # "Stručni studij X" = prijediplomski stručni
    return section  # npr. "Čembalo - izvođački smjer" u diplomskom dijelu tablice


def _short_name(full):
    words = full.split(" ")
    i = 0
    while i < len(words) - 1 and words[i].lower().strip(",;:") in LEVEL_WORDS:
        i += 1
    rest = " ".join(words[i:]).strip()
    rest = re.sub(r"^za\s+", "", rest)
    return (rest[:1].upper() + rest[1:]) if rest else full


def _faculty_header(text):
    """Naziv sastavnice ako je redak naslov sastavnice, inače None."""
    t = _clean(text)
    m = re.search(r"dodani kodovi za (.+?fakultet)", t, flags=re.I)  # "...dodani kodovi za Ekonomski fakultet ..."
    if m:
        return _faculty_name(m.group(1))
    if not t or t.isupper() or "dodan" in t.lower() or t.lower().startswith(("kodovi", "nitko", "muza")):
        return None
    if re.match(r"^\d{1,2}\.\s*\d{1,2}\.\s*\d{4}\.\s*\D", t) or not re.match(r"^\d", t):
        if re.search(r"fakultet|akademija|studiji|HVU|Sveučilište u Zagrebu", t, flags=re.I):
            return _faculty_name(t)
    return None


def parse_codes_table(path):
    """Čita tablicu kodova (.xlsx, ili .xls ako je instaliran xlrd). Vraća ({kod: [zapisi]}, napomene)."""
    import pandas as pd
    df = pd.read_excel(path, header=None, dtype=str)
    faculty, section, out, notes = "", PREDDIPLOMSKI, {}, []
    for _, row in df.iterrows():
        text = row.iloc[1] if len(row) > 1 else None
        code = str(row.iloc[2]).strip() if len(row) > 2 and isinstance(row.iloc[2], str) else ""
        if not isinstance(text, str) or not text.strip():
            continue
        if re.fullmatch(r"\d{1,4}", code):
            code = code.zfill(4)
            full = _clean(text)
            entry = {"fakultet": faculty, "program": _short_name(full), "puni_naziv": full,
                     "razina": _level(full, section)}
            existing = out.setdefault(code, [])
            if any(e["puni_naziv"] == full for e in existing):
                continue
            if existing:
                notes.append(f"Kod {code} se pojavljuje više puta: " + " / ".join(
                    e["puni_naziv"] for e in existing + [entry]))
            existing.append(entry)
            continue
        if "PODRUČJE PRIRODNIH" in text.upper() and out:
            section = DIPLOMSKI  # Drugi dio tablice su diplomski studiji
        header = _faculty_header(text)
        if header:
            faculty = header
    return dict(sorted(out.items())), notes


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    data, notes = parse_codes_table(sys.argv[1])
    with open(PROGRAMS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    print(f"Spremljeno {len(data)} kodova u {PROGRAMS_FILE}")
    for n in notes:
        print(" -", n)
