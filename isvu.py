"""
ISVU online ankete (V4, procjena nastavnika): čitanje "sirovih podataka" iz ISVU-a.

ISVU Studiji i studenti -> Anketa za VU -> "Ispis sirovih podataka za VU (XML)".
Datoteka je u kodnoj stranici windows-1250 (cp1250). Iz nje se dobiva tablica s jednim retkom
po ispunjenom anketnom listiću (student x kolegij x nastavnik x grupa), u istim stupcima kao
papirnate ankete V4 iz OMR čitača, uz dodatne stupce o izvoru.

Kodovi u XML-u:
    A (pitanja o kolegiju, jednom po studentu i kolegiju): 309 spol (1=Ž, 2=M), 310 prisutnost 1-3,
      311 interes 1-3, 312 najčešća ocjena 2-5, 313 očekivana ocjena 2-5
    B (pitanja o nastavniku, po grupi): 315-323 = B1-B9 (1-5, 6 = ne mogu procijeniti), 324 = B10
U tablici: Spol 0=Ž / 1=M (kao na papiru), "x" = bez odgovora, "n" = ne mogu procijeniti.
"""

import io
import os
import re
import zipfile
import xml.etree.ElementTree as ET

import pandas as pd

SOURCE = "ISVU"
MISSING, CANT_SAY = "x", "n"

A_CODES = {"309": "Spol", "310": "Prisutnost", "311": "Interes kolegija", "312": "Indeks", "313": "Ocjena kolegija"}
B_CODES = {str(315 + i): f"B{i + 1}" for i in range(10)}  # 315..324 -> B1..B10
A_COLUMNS = list(A_CODES.values())
B_COLUMNS = list(B_CODES.values())
ANSWER_COLUMNS = A_COLUMNS + B_COLUMNS

INFO_COLUMNS = ["Izvor", "Fakultet", "Akademska godina", "Semestar", "Šifra predmeta", "Kolegij",
                "Oznaka nastavnika", "Nastavnik", "Grupa", "Vrsta nastave", "Pristupnik"]
COLUMNS = INFO_COLUMNS + ANSWER_COLUMNS

SEMESTERS = {"1": "Zimski", "2": "Ljetni"}

# Nazivi sastavnica u ISVU-u -> nazivi kakvi se koriste u ostatku programa (popis kodova, OneDrive mape)
FACULTY_RENAME = {
    "Šumarski fakultet": "Fakultet šumarstva i drvne tehnologije",
    "Hrvatski studiji": "Fakultet hrvatskih studija",
}


class ISVUError(ValueError):
    pass


def _decode(data):
    """XML kao tekst. ISVU piše windows-1250; ako je datoteka ipak UTF-8, i to se prihvaća."""
    head = data[:200].decode("ascii", "ignore").lower()
    for enc in (["utf-8", "cp1250"] if "utf-8" in head else ["cp1250", "utf-8"]):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("cp1250", "replace")


def _answer(value, b_item=False):
    v = (value or "").strip()
    if not v:
        return MISSING
    if b_item and v == "6":
        return CANT_SAY
    return int(v) if v.isdigit() else MISSING


def _question_codes(root, file_name=""):
    """Šifre pitanja -> stupci. Šifre ovise o anketi (npr. 309-324 za anketu 307, 9042-9056 za 9038),
    pa se određuju iz kataloga pitanja: pitanja o kolegiju redom su A1-A5, pitanja o nastavniku B1-B10."""
    order = [q.get("sifPitanje") for q in root.iterfind("katalogPitanje/pitanje")]
    a_seen = {q.get("sifPitanje") for q in root.iterfind("predmeti/predmet/pitanjaZaPredmet/pitanje")}
    b_seen = {q.get("sifPitanje") for q in
              root.iterfind("predmeti/predmet/nastavnici/nastavnik/nastGrupe/nastGrupa/pitanjaZaNastGrupa/pitanje")}
    if not a_seen and not b_seen:
        return dict(A_CODES), dict(B_CODES)  # Prazan izvoz
    key = lambda c: (order.index(c) if c in order else len(order), int(c) if str(c).isdigit() else 0)
    if len(order) >= 15:  # Katalog: prvih 5 pitanja o kolegiju, zatim 10 o nastavniku
        a_list, b_list = order[:5], order[5:15]
    else:
        a_list, b_list = sorted(a_seen, key=key), sorted(b_seen, key=key)
    if not (a_seen <= set(a_list) and b_seen <= set(b_list)) or len(a_list) != 5 or len(b_list) != 10:
        raise ISVUError(f"{file_name}: neočekivan popis pitanja ({len(a_list)} o kolegiju, {len(b_list)} o "
                        "nastavniku; očekuje se 5 i 10)")
    return dict(zip(a_list, A_COLUMNS)), dict(zip(b_list, B_COLUMNS))


def parse(data, file_name=""):
    """Bajtovi jedne XML datoteke -> (opći podaci, DataFrame s jednim retkom po anketnom listiću).

    Opći podaci: {"Fakultet", "Akademska godina", "Semestar", "Anketa", "Šifra ankete", "Šifra VU",
                  "Datum od", "Datum do", "Datoteka"}
    """
    text = _decode(data)
    text = re.sub(r"^<\?xml[^>]*\?>", "", text.lstrip("﻿ \r\n\t"))  # deklaracija kodiranja smeta ET-u
    try:
        root = ET.fromstring(text)
    except ET.ParseError as e:
        raise ISVUError(f"{file_name}: nije ispravan XML ({e})") from None
    if root.tag != "anketa" or root.find("opciPodaci") is None:
        raise ISVUError(f"{file_name}: nije ISVU \"Ispis sirovih podataka za VU (XML)\"")

    op = root.find("opciPodaci")
    g = lambda tag: (op.findtext(tag) or "").strip()
    year = int(g("akGodina")) if g("akGodina").isdigit() else None
    faculty = FACULTY_RENAME.get(g("nazVU"), g("nazVU"))
    meta = {
        "Fakultet": faculty,
        "Akademska godina": f"{year}./{year + 1}." if year else "",
        "Semestar": SEMESTERS.get(g("rbrElUstNast"), g("oznElUstNast")),
        "Anketa": g("nazAnketa"),
        "Šifra ankete": g("sifAnketa"),
        "Šifra VU": g("sifVU"),
        "Datum od": g("datumOd"),
        "Datum do": g("datumDo"),
        "Datoteka": os.path.basename(file_name),
    }
    if "V4" not in meta["Anketa"] and "procjen" not in meta["Anketa"].lower():
        raise ISVUError(f"{file_name}: anketa \"{meta['Anketa']}\" nije anketa za procjenu nastavnika (V4)")

    a_codes, b_codes = _question_codes(root, file_name)
    base = [SOURCE, faculty, meta["Akademska godina"], meta["Semestar"]]
    rows, pairs = [], set()
    for p in root.iterfind("predmeti/predmet"):
        course_id, course = p.get("sifPred", ""), (p.get("nazPred") or "").strip()
        a_answers = {}  # pristupnik -> {stupac: odgovor}
        for q in p.iterfind("pitanjaZaPredmet/pitanje"):
            col = a_codes.get(q.get("sifPitanje"))
            if not col:
                continue
            for o in q.iterfind("odgovori/odgovor"):
                a_answers.setdefault(o.get("sifPristupnik"), {})[col] = _answer(o.findtext("rbrOdgovor"))
        # Listić = student x nastavnik x grupa na kolegiju. Ako ISVU isti listić navodi više puta (npr. ista
        # grupa i za predavanja i za seminar), računa se jednom, kao i u ISVU-ovoj obradi rezultata.
        forms, vrsta = {}, {}
        for n in p.iterfind("nastavnici/nastavnik"):
            teacher_id = (n.get("oznNastavnik") or "").strip()
            teacher = " ".join(x for x in [(n.get("imeNastavnik") or "").strip(),
                                           (n.get("prezNastavnik") or "").strip()] if x)
            for grp in n.iterfind("nastGrupe/nastGrupa"):
                group = grp.get("oznGrupa", "")
                for q in grp.iterfind("pitanjaZaNastGrupa/pitanje"):
                    col = b_codes.get(q.get("sifPitanje"))
                    if not col:
                        continue
                    for o in q.iterfind("odgovori/odgovor"):
                        key = (teacher_id, teacher, group, o.get("sifPristupnik"))
                        forms.setdefault(key, {})[col] = _answer(o.findtext("rbrOdgovor"), True)
                        vrsta.setdefault(key, set()).add(grp.get("oznVrNast", ""))
        for (teacher_id, teacher, group, student), b in forms.items():
            a = a_answers.get(student, {})
            spol = a.get("Spol", MISSING)
            rows.append(base + [course_id, course, teacher_id, teacher, group,
                                "/".join(sorted(v for v in vrsta[(teacher_id, teacher, group, student)] if v)),
                                student]
                        + [{1: 0, 2: 1}.get(spol, MISSING)]
                        + [a.get(c, MISSING) for c in A_COLUMNS[1:]]
                        + [b.get(c, MISSING) for c in B_COLUMNS])
        # Student koji je odgovorio samo na pitanja o kolegiju (bez ijednog nastavnika): redak bez nastavnika,
        # da se broji u anketiranjima i u A dijelu izvještaja (nije anketni listić)
        with_form = {student for (_t, _n, _g, student) in forms}
        for student, a in a_answers.items():
            if student not in with_form:
                spol = a.get("Spol", MISSING)
                rows.append(base + [course_id, course, "", "", "", "", student]
                            + [{1: 0, 2: 1}.get(spol, MISSING)]
                            + [a.get(c, MISSING) for c in A_COLUMNS[1:]] + [MISSING] * len(B_COLUMNS))
        pairs.update((course_id, s) for s in a_answers)
    df = pd.DataFrame(rows, columns=COLUMNS)
    pairs.update(zip(df["Šifra predmeta"], df["Pristupnik"]))
    meta["_pairs"] = pairs
    _counts(meta, df)
    return meta, df


def _counts(meta, df):
    meta["Anketiranja"] = len(meta["_pairs"])
    meta["Anketni listići"] = int((df["Oznaka nastavnika"] != "").sum())
    meta["Nastavnika"] = df.loc[df["Oznaka nastavnika"] != "", "Oznaka nastavnika"].nunique()
    meta["Kolegija"] = df["Šifra predmeta"].nunique()


def read_uploads(files):
    """Učitane datoteke (XML ili ZIP s XML-ovima) -> [(ime, bajtovi)]."""
    out = []
    for name, data in files:
        if name.lower().endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                for info in z.infolist():
                    if not info.is_dir() and info.filename.lower().endswith(".xml"):
                        out.append((os.path.basename(info.filename), z.read(info)))
        elif name.lower().endswith(".xml"):
            out.append((name, data))
    return out


def unit_key(meta):
    """Ključ za spajanje datoteka iste sastavnice i semestra (npr. "dio 2" izvoza)."""
    return (meta["Fakultet"], meta["Akademska godina"], meta["Semestar"])


def combine(parsed):
    """[(meta, df)] -> {(fakultet, godina, semestar): (meta, df)}; dijelovi istog semestra se spajaju,
    a listić koji se pojavi u više dijelova broji se jednom."""
    units = {}
    for meta, df in parsed:
        key = unit_key(meta)
        if key in units:
            m0, d0 = units[key]
            df = pd.concat([d0, df], ignore_index=True).drop_duplicates(
                ["Šifra predmeta", "Oznaka nastavnika", "Grupa", "Pristupnik"])
            meta = {**m0, "Datoteka": f"{m0['Datoteka']}, {meta['Datoteka']}",
                    "_pairs": m0["_pairs"] | meta["_pairs"]}
        units[key] = (meta, df)
    for meta, df in units.values():
        _counts(meta, df)
    return units
