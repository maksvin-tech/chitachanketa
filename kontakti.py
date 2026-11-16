"""
Kontakti fakulteta i e-mail uz gotov izvještaj.

Popis kontakata NE stoji u repozitoriju (repozitorij je javan, a to su tuđe adrese).
Učitava se jednom u aplikaciji (bočna traka -> "Kontakti fakulteta") iz tablice
"Kontakti fakulteta – studentske ankete" (preuzeta kao .xlsx ili .csv) i sprema se kao
kontakti.csv u mapu podataka (uz google_credentials.json).

Stupci tablice: Fakultet, Ime i prezime, E-mail, Ankete, Izvor, Napomena, Oslovljavanje
    Ankete        : Nastavnici / Studijski programi / Oboje
    Izvor         : Izravna komunikacija / Samo mailing lista / Adresa ne radi
    Oslovljavanje : npr. "Poštovana kolegice Borošić," (prazno = "Poštovani,")

E-mail se ne šalje iz programa: program složi poruku i otvori je u Outlooku
(poveznica mailto:), a pošiljatelj je račun prijavljen u Outlooku na tom računalu.
"""

import io
import os
import unicodedata
from urllib.parse import quote

import pandas as pd

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.environ.get("OMR_DATA_DIR") or APP_DIR
CONTACTS_FILE = os.path.join(DATA_DIR, "kontakti.csv")

COLUMNS = ["Fakultet", "Ime i prezime", "E-mail", "Ankete", "Izvor", "Napomena", "Oslovljavanje"]
REQUIRED = ["Fakultet", "E-mail"]

KIND_TEACHERS = "Nastavnici"
KIND_STUDIES = "Studijski programi"
KIND_BOTH = "Oboje"
DEAD = "Adresa ne radi"
DIRECT = "Izravna komunikacija"

DEFAULT_GREETING = "Poštovani,"

SIGNATURE = (
    "Srdačan pozdrav,\n"
    "Maks Vinšćak\n"
    "Ured za razvoj karijera Sveučilišta u Zagrebu\n"
    "Mob: 0915768572\n"
    "Email: mvinscak@unizg.hr"
)

# Riječi koje se ne broje kad se uspoređuju nazivi fakulteta
_STOP = {"fakultet", "sveuciliste", "u", "zagrebu", "i", "za", "akademija", "studij", "studiji", "the"}


# ------------------------------------------------------------------------------
# Učitavanje i spremanje
# ------------------------------------------------------------------------------

def _clean(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError("U tablici nedostaju stupci: " + ", ".join(missing))
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[COLUMNS].fillna("").astype(str)
    for col in COLUMNS:
        df[col] = df[col].str.strip()
    df = df[df["E-mail"].str.contains("@")]
    return df.reset_index(drop=True)


def parse_upload(name, data):
    """Čita tablicu kontakata iz .xlsx ili .csv (bajtovi). Vraća DataFrame ili baca ValueError."""
    if name.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(io.BytesIO(data), dtype=str)
    else:
        text = data.decode("utf-8-sig", errors="replace")
        sep = ";" if text.splitlines() and text.splitlines()[0].count(";") > text.splitlines()[0].count(",") else ","
        df = pd.read_csv(io.StringIO(text), sep=sep, dtype=str)
    df = _clean(df)
    if df.empty:
        raise ValueError("U tablici nema nijedne e-mail adrese.")
    return df


def save_contacts(df):
    os.makedirs(DATA_DIR, exist_ok=True)
    df[COLUMNS].to_csv(CONTACTS_FILE, index=False, encoding="utf-8")


def load_contacts():
    """Vraća spremljene kontakte (DataFrame) ili prazan DataFrame ako ih nema."""
    if not os.path.exists(CONTACTS_FILE):
        return pd.DataFrame(columns=COLUMNS)
    try:
        return _clean(pd.read_csv(CONTACTS_FILE, dtype=str))
    except Exception:
        return pd.DataFrame(columns=COLUMNS)


# ------------------------------------------------------------------------------
# Odabir kontakata za izvještaj
# ------------------------------------------------------------------------------

def _words(text):
    plain = unicodedata.normalize("NFKD", text.lower())
    plain = "".join(c for c in plain if not unicodedata.combining(c)).replace("đ", "d")
    for ch in "-–,.()\"'/":
        plain = plain.replace(ch, " ")
    return {w for w in plain.split() if w not in _STOP and len(w) > 2}


def faculty_score(a, b):
    """Koliko se dva naziva fakulteta poklapaju (0 = nimalo). Riječi se uspoređuju po
    početku (npr. 'medicinski' i 'medicina'), jer se nazivi pišu na više načina."""
    wa, wb = _words(a), _words(b)
    if not wa or not wb:
        return 0.0
    hits = sum(1 for x in wa if any(x[:6] == y[:6] for y in wb))
    return hits / max(len(wa), len(wb))


def suggestions(contacts, faculty, kind):
    """Kontakti za taj fakultet i tu vrstu ankete, najbolji prvi.
    Vraća (popis prikladnih, popis svih ostalih ispravnih adresa)."""
    if contacts.empty:
        return [], []
    rows = contacts[contacts["Izvor"] != DEAD]
    scored = []
    for _, r in rows.iterrows():
        score = faculty_score(faculty or "", r["Fakultet"])
        fits = r["Ankete"] in (kind, KIND_BOTH, "")
        scored.append((score, fits, r.to_dict()))
    best = max((s for s, _, _ in scored), default=0)
    good = [r for s, fits, r in scored if best > 0 and s == best and fits]
    if not good:  # Nitko s tog fakulteta nije označen za ovu vrstu ankete: ponudi sve s tog fakulteta
        good = [r for s, _fits, r in scored if best > 0 and s == best]
    # Oni s kojima se već dopisivalo idu prvi, zatim imenovane osobe, pa opće adrese
    good.sort(key=lambda r: (r["Izvor"] != DIRECT, not r["Ime i prezime"], r["Ime i prezime"]))
    ids = {r["E-mail"] for r in good}
    rest = sorted((r for _, _, r in scored if r["E-mail"] not in ids),
                  key=lambda r: (r["Fakultet"], r["Ime i prezime"]))
    return good, rest


def label(contact):
    name = contact.get("Ime i prezime") or ""
    text = f"{name} <{contact['E-mail']}>" if name else contact["E-mail"]
    return f"{text} · {contact.get('Fakultet', '')}"


def greeting(contact):
    return (contact.get("Oslovljavanje") or "").strip() or DEFAULT_GREETING


# ------------------------------------------------------------------------------
# Tekst poruke
# ------------------------------------------------------------------------------

def _join(parts):
    return ", ".join(p for p in parts if p)


def compose(kind, info, links, semester=None, year=None, greet=DEFAULT_GREETING, signature=True):
    """Složi naslov i tekst standardne poruke.

    kind  : KIND_TEACHERS ili KIND_STUDIES
    info  : upisani podaci serije (Fakultet, Kolegij, Nastavnik / Kod studijskog programa, Studijski program)
    links : popis (opis, poveznica), npr. [("Izvješće", url), ("Skenirane ankete", url)]
    """
    if kind == KIND_TEACHERS:
        who = info.get("Nastavnik", "")
        course = info.get("Kolegij", "")
        when = _join([f"{semester} semestar" if semester else "", f"ak. god. {year}" if year else ""])
        subject = "Rezultati ankete za procjenu nastavnika – " + _join([who, course])
        what = "izvješće o provedenoj anketi za procjenu nastavnika"
        detail = _join([who, course, when])
    else:
        code = info.get("Kod studijskog programa", "")
        program = info.get("Studijski program", "")
        name = " ".join(p for p in [code, program] if p)
        subject = "Rezultati ankete za vrednovanje studija – " + name
        what = "izvješće o provedenoj anketi za vrednovanje studija"
        detail = _join([name, f"ak. god. {year}" if year else ""])

    has_scans = any("sken" in d.lower() for d, _ in links)
    body = f"{greet}\n\nšaljem vam {what}"
    body += f" ({detail})" if detail else ""
    body += " i skenirane ankete." if has_scans else "."
    if links:
        body += "\n\n" + "\n".join(f"{d}: {u}" for d, u in links)
    body += "\n\n" + (SIGNATURE if signature else "Srdačan pozdrav,\nMaks Vinšćak")
    return subject.strip(" –"), body


def compose_bulk(kind, faculty, folder_link, count, year=None, greet=DEFAULT_GREETING, signature=True):
    """Poruka za više izvještaja jednog fakulteta odjednom (poveznica na mapu)."""
    survey = "procjenu nastavnika" if kind == KIND_TEACHERS else "vrednovanje studija"
    subject = f"Rezultati ankete za {survey} – {faculty}"
    body = (f"{greet}\n\nšaljem vam izvješća o provedenoj anketi za {survey}"
            f"{f', ak. god. {year}' if year else ''} ({count}).\n\n"
            f"Izvješća: {folder_link}\n\n" + (SIGNATURE if signature else "Srdačan pozdrav,\nMaks Vinšćak"))
    return subject, body


def _and(items):
    items = [i for i in items if i]
    return " i ".join([", ".join(items[:-1]), items[-1]]) if len(items) > 1 else (items[0] if items else "")


def compose_folder(kinds, faculty, folder_link, year=None, greet=DEFAULT_GREETING, signature=True):
    """Poruka s poveznicom na mapu fakulteta (izvještaji i skenirane ankete).
    kinds: nazivi vrsta u mapi, npr. ["Prijediplomski studij", "Diplomski studij"]"""
    studies = [k.split()[0].lower() for k in kinds if "studij" in k.lower()]
    surveys = []
    if any("nastavnik" in k.lower() for k in kinds):
        surveys.append("procjenu nastavnika")
    if studies:
        surveys.append(f"vrednovanje studija ({_and(studies)} studij)")
    survey_text = _and(surveys) or "studentske ankete"
    subject = f"Rezultati anketa – {faculty}"
    body = (f"{greet}\n\nšaljem vam poveznicu na mapu s izvješćima i skeniranim anketama o provedenoj anketi "
            f"za {survey_text}{f', ak. god. {year.rstrip(chr(46))}' if year else ''}.\n\n"
            f"Mapa: {folder_link}\n\n" + (SIGNATURE if signature else "Srdačan pozdrav,\nMaks Vinšćak"))
    return subject, body


def mailto(to, subject, body, cc=None):
    """Poveznica koja otvara novu poruku u zadanom programu za e-poštu (Outlook)."""
    params = [f"subject={quote(subject)}", f"body={quote(body.replace(chr(10), chr(13) + chr(10)))}"]
    if cc:
        params.insert(0, f"cc={quote(cc, safe='@,')}")
    return f"mailto:{quote(to, safe='@,')}?" + "&".join(params)
