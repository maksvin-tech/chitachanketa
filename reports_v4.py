"""
Izvješća o studentskom vrednovanju nastave (anketa V4) za sastavnicu ili cijelo Sveučilište:
ISVU online, papir-olovka ili oboje zajedno; zimski, ljetni semestar ili cijela akademska godina.

Ulaz je tablica s jednim retkom po anketnom listiću u stupcima isvu.COLUMNS (ISVU izravno iz isvu.parse,
papirnate ankete iz OMR čitača preko paper_rows). Za svaku jedinicu nastaju:
    izvješće (tablice A, 1-11, slike 1-2; za Sveučilište i tablica 12 i zaključak),
    ključ za identifikaciju nastavnika, rang (najmanje 1 anketa), rang (najmanje 10 anketa).

Pravila (kao ISVU-ova obrada, provjereno na izvozima 2025./2026.):
  - anketiranje = jedan student na jednom predmetu (A dio: spol, prisutnost, interes, ocjene);
    kod papirnatih anketa svaki je listić jedno anketiranje
  - anketa (listić) = student x nastavnik x grupa na predmetu (B dio)
  - M i SD samo iz odgovora 1-5 ("ne mogu procijeniti" i bez odgovora ne ulaze); SD populacijska (kao ISVU)
  - D = raspon (max - min); NN = broj anketa bez valjanog odgovora
  - postoci od ukupnog broja anketiranja (A) odnosno anketa (B), uključujući bez odgovora
  - složeni pokazatelji: prosjek triju pitanja po anketi, samo ankete s odgovorom na sva tri
  - rang: prosjek svih odgovora B.1-B.9 za nastavnika na predmetu, zaokružen na 2 decimale; jednaki dijele mjesto
"""

import io
import re
from datetime import date

import numpy as np
import pandas as pd
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import Image, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, TableStyle

import isvu
import reports as base
from reports import _esc, _table, fmt, fmt_int, hbar_chart

SOURCES = {"isvu": "ISVU online", "papir": "papir-olovka", "oboje": "ISVU online i papir-olovka (objedinjeno)"}
SOURCE_OF_ROW = {"isvu": isvu.SOURCE, "papir": "papir"}
UNIVERSITY = "Sveučilište u Zagrebu"
MIN_RANK = 10

A_ITEMS = [
    ("Spol", "A.1 Spol", [(0, "Ženski"), (1, "Muški")]),
    ("Prisutnost", "A.2 Prisutnost na nastavi", [(1, "Rijetka (do 30%)"), (2, "Povremena (30-70%)"),
                                                (3, "Redovita (više od 70%)")]),
    ("Interes kolegija", "A.3 Početni interes za sadržaj predmeta", [(1, "Mali"), (2, "Srednji"), (3, "Veliki")]),
    ("Indeks", "A.4 Najčešća ocjena tijekom dosadašnjeg studija", [(v, str(v)) for v in (2, 3, 4, 5)]),
    ("Ocjena kolegija", "A.5 Očekivana ocjena iz predmeta", [(v, str(v)) for v in (2, 3, 4, 5)]),
]
B_TEXTS = [text for _c, text in base.V4_B_ITEMS] + [base.V4_B10[1]]
B9 = [f"B{i}" for i in range(1, 10)]
B_ALL = B9 + ["B10"]
COMPOSITES = [("Pokazatelj 1 (B.1-B.3): jasnoća i organizacija nastave", B9[0:3]),
              ("Pokazatelj 2 (B.4-B.6): metode i komunikacija", B9[3:6]),
              ("Pokazatelj 3 (B.7-B.9): odnos prema obvezama i studentima", B9[6:9]),
              ("Ukupno (B.1-B.9)", B9)]
SCALE_LABELS = ["1", "2", "3", "4", "5"]
BAR = base.BAR


# ------------------------------------------------------------------------------
# Priprema podataka
# ------------------------------------------------------------------------------

def _num(series):
    """Odgovor -> broj 1-5 (ili 0-1 za spol); 'x', 'n' i prazno -> NaN."""
    return pd.to_numeric(series.astype(str).str.strip().str.replace(",", ".", regex=False), errors="coerce")


def semester_from_batch(batch_name):
    """Papirnate serije bez upisanog semestra: iz datuma serije (studeni-veljača = zimski, ostalo ljetni)."""
    m = re.match(r"(\d{4})-(\d{2})-\d{2}", str(batch_name))
    if not m:
        return "", ""
    y, mo = int(m.group(1)), int(m.group(2))
    start = y if mo >= 10 else y - 1
    return f"{start}./{start + 1}.", ("Zimski" if mo >= 11 or mo <= 2 else "Ljetni")


def paper_rows(df):
    """Papirnate ankete V4 iz OMR čitača (spremljene serije) -> retci u stupcima isvu.COLUMNS."""
    if df is None or df.empty:
        return pd.DataFrame(columns=isvu.COLUMNS)
    out = pd.DataFrame(index=df.index)
    get = lambda c: df[c].fillna("").astype(str).str.strip() if c in df.columns else pd.Series("", index=df.index)
    inferred = get("Serija").map(semester_from_batch)
    year = get("Akademska godina").where(get("Akademska godina") != "", inferred.str[0])
    sem = get("Semestar").str.capitalize().where(get("Semestar") != "", inferred.str[1])
    key = get("Identifikacijski ključ").where(get("Identifikacijski ključ") != "", get("Nastavnik"))
    out["Izvor"] = "papir"
    out["Fakultet"] = get("Fakultet")
    out["Akademska godina"] = year
    out["Semestar"] = sem
    out["Šifra predmeta"] = get("Kolegij")
    out["Kolegij"] = get("Kolegij")
    out["Oznaka nastavnika"] = key.where(key != "", "?")
    out["Nastavnik"] = get("Nastavnik")
    out["Grupa"] = ""
    out["Vrsta nastave"] = ""
    out["Pristupnik"] = [f"P{i}" for i in range(len(df))]
    for c in isvu.ANSWER_COLUMNS:
        out[c] = get(c).replace("", isvu.MISSING)
    return out[isvu.COLUMNS].reset_index(drop=True)


def select(rows, academic_year, semester=None, faculty=None, source="isvu"):
    """semester: "Zimski", "Ljetni" ili None (cijela godina); faculty None = Sveučilište."""
    r = rows[rows["Akademska godina"] == academic_year]
    if semester:
        r = r[r["Semestar"] == semester]
    if faculty:
        r = r[r["Fakultet"] == faculty]
    if source in SOURCE_OF_ROW:
        r = r[r["Izvor"] == SOURCE_OF_ROW[source]]
    return r


# ------------------------------------------------------------------------------
# Statistika jedne jedinice
# ------------------------------------------------------------------------------

def _desc(values):
    v = values.dropna()
    n = len(v)
    return {"N": n, "M": v.mean() if n else np.nan, "SD": v.std(ddof=0) if n else np.nan,
            "D": (v.max() - v.min()) if n else np.nan, "Min": v.min() if n else np.nan,
            "Max": v.max() if n else np.nan}


def _anket_key(rows):
    return rows["Izvor"] + "|" + rows["Fakultet"] + "|" + rows["Semestar"] + "|" + rows["Šifra predmeta"] + "|" \
        + rows["Pristupnik"]


def unit_stats(rows):
    """Svi brojevi za izvješće jedne jedinice."""
    rows = rows.reset_index(drop=True)
    anket = rows.loc[~_anket_key(rows).duplicated()]
    forms = rows[rows["Oznaka nastavnika"] != ""].copy()
    nf = len(forms)
    s = {"anketiranja": len(anket), "ankete": nf,
         "nastavnika": forms["Oznaka nastavnika"].nunique() if nf else 0,
         "predmeta": forms["Šifra predmeta"].nunique() if nf else 0}
    # A
    s["A"] = {}
    for col, _title, cats in A_ITEMS:
        vals = _num(anket[col])
        s["A"][col] = {"total": len(anket), "missing": int(vals.isna().sum()),
                       "counts": [(lbl, int((vals == v).sum())) for v, lbl in cats]}
    # B
    nums = {c: _num(forms[c]).where(lambda x: x.between(1, 5)) for c in B_ALL}
    raw = {c: forms[c].astype(str).str.strip() for c in B_ALL}
    s["B"] = {}
    for c in B_ALL:
        d = _desc(nums[c])
        d["NN"] = nf - d["N"]
        d["cant"] = int((raw[c] == isvu.CANT_SAY).sum())
        d["blank"] = nf - d["N"] - d["cant"]
        d["counts"] = [int((nums[c] == v).sum()) for v in range(1, 6)]
        s["B"][c] = d
    pooled = pd.concat([nums[c] for c in B9], ignore_index=True)
    s["B9"] = {**_desc(pooled), "NN": nf * 9 - int(pooled.notna().sum())}
    s["composites"] = []
    for name, cols in COMPOSITES:
        m = pd.concat([nums[c] for c in cols], axis=1)
        per_form = m.mean(axis=1)[m.notna().all(axis=1)]
        s["composites"].append((name, _desc(per_form)))
    # Nastavnik x predmet
    if nf:
        forms["_m9_sum"] = sum(nums[c].fillna(0) for c in B9)
        forms["_m9_n"] = sum(nums[c].notna().astype(int) for c in B9)
        forms["_b10"] = nums["B10"]
        g = forms.groupby(["Izvor", "Oznaka nastavnika", "Šifra predmeta"], sort=False)
        keys = g.agg(Nastavnik=("Nastavnik", "first"), Kolegij=("Kolegij", "first"), N=("Pristupnik", "size"),
                     s9=("_m9_sum", "sum"), n9=("_m9_n", "sum"), M10=("_b10", "mean")).reset_index()
        keys["M9"] = keys["s9"] / keys["n9"].replace(0, np.nan)
        keys = keys.drop(columns=["s9", "n9"])
    else:
        keys = pd.DataFrame(columns=["Izvor", "Oznaka nastavnika", "Šifra predmeta", "Nastavnik", "Kolegij",
                                     "N", "M10", "M9"])
    s["keys"] = keys
    return s


def assign_keys(keys, prefix=""):
    """Identifikacijski ključ za svaku kombinaciju nastavnik / predmet (abecedno po nastavniku)."""
    k = keys.copy()
    k["_sort"] = k["Nastavnik"].fillna("").str.split().str[-1].fillna("") + " " + k["Nastavnik"].fillna("")
    k = k.sort_values(["_sort", "Kolegij", "Izvor"]).drop(columns="_sort").reset_index(drop=True)
    width = max(3, len(str(len(k))))
    k["Ključ"] = [f"{prefix}{i + 1:0{width}d}" + (" (P)" if src == "papir" and (k["Izvor"] != "papir").any() else "")
                  for i, src in enumerate(k["Izvor"])]
    return k


def ranking(keys, min_n):
    r = keys[(keys["N"] >= min_n) & keys["M9"].notna()].copy()
    r["_r"] = r["M9"].round(2)
    r = r.sort_values(["_r", "N"], ascending=[False, False]).reset_index(drop=True)
    r["Rang"] = r["_r"].rank(method="min", ascending=False).astype(int)
    return r.drop(columns="_r")


# ------------------------------------------------------------------------------
# PDF
# ------------------------------------------------------------------------------

def _styles():
    st = base._styles()
    st.update({
        "h1": ParagraphStyle("h1", fontName="Serif-Bold", fontSize=13, leading=16, spaceBefore=10, spaceAfter=8),
        "h2": ParagraphStyle("h2", fontName="Serif-Bold", fontSize=11.5, leading=14, spaceBefore=8, spaceAfter=6),
        "body": ParagraphStyle("body", fontName="Serif", fontSize=10.5, leading=14, alignment=TA_JUSTIFY,
                               spaceAfter=6),
        "small": ParagraphStyle("small", fontName="Serif-Italic", fontSize=9, leading=11, alignment=TA_LEFT,
                                spaceBefore=2, spaceAfter=8),
        "center": ParagraphStyle("center", fontName="Serif", fontSize=11, leading=14, alignment=TA_CENTER),
    })
    return st


def period_label(academic_year, semester):
    return f"{semester.lower()} semestar ak. god. {academic_year}" if semester else f"akademska godina {academic_year}"


def _cover(st, title, unit, academic_year, semester, source, report_date, lines=()):
    story = []
    if base.os.path.exists(base.LOGO):
        logo = Image(base.LOGO, width=9.5 * cm, height=9.5 * cm * 159 / 420)
        logo.hAlign = "CENTER"
        story.append(logo)
    story += [Spacer(1, 2.6 * cm), Paragraph(_esc(title), st["title"]), Spacer(1, 0.3 * cm),
              Paragraph("Anketa za procjenu nastavnika (V4)", st["cover_italic"]), Spacer(1, 1.6 * cm),
              Paragraph(_esc(unit), st["cover_bold"]), Spacer(1, 0.4 * cm),
              Paragraph(_esc(period_label(academic_year, semester).capitalize()), st["cover_bold"]),
              Spacer(1, 0.4 * cm), Paragraph(f"Način provedbe: {_esc(SOURCES[source])}", st["cover_italic"])]
    for line in lines:
        story.append(Paragraph(_esc(line), st["cover_italic"]))
    story += [Spacer(1, 3.0 * cm), Paragraph("Ured za upravljanje kvalitetom", st["cover_bold"]),
              Paragraph(UNIVERSITY, st["cover_bold"]), Spacer(1, 0.5 * cm),
              Paragraph(f"Zagreb, {report_date.day}. {report_date.month}. {report_date.year}.", st["cover_bold"]),
              PageBreak()]
    return story


def _pct(n, total):
    return fmt(n / total * 100, 1) if total else "–"


def _method_text(source, semester, university):
    parts = []
    if source in ("isvu", "oboje"):
        parts.append("elektronički, putem Informacijskog sustava visokih učilišta (ISVU)")
    if source in ("papir", "oboje"):
        parts.append("papirnatim anketnim listićima (papir-olovka) koji su očitani programom OMR Ankete")
    text = ("Anketa je provedena " + " i ".join(parts) + ", anketnim listom za procjenu nastavnika (V4). "
            "Svrha ankete nije kontrola rada nastavnika, niti anketa sama za sebe može poboljšati kvalitetu, ali je "
            "važan oblik dijagnostike dobrih i loših elemenata nastavnoga procesa iz perspektive studenata.")
    if university:
        text += " Rezultati su prikazani zbirno za sve anketirane sastavnice te po sastavnicama."
    else:
        text += (" Rezultati su obrađeni na razini anketirane jedinice i nastavnika na pojedinom predmetu. Uz izvješće "
                 "se dostavljaju ključ za identifikaciju nastavnika i rangiranje nastavnika (uz najmanje 1 i "
                 f"najmanje {MIN_RANK} ispunjenih anketa).")
    if not semester:
        text += " Rezultati za akademsku godinu dobiveni su objedinjavanjem anketa zimskog i ljetnog semestra."
    if source == "oboje":
        text += (" Rezultati elektroničke i papirnate ankete istog nastavnika na istom predmetu vode se pod zasebnim "
                 "ključevima; ključevi papirnate ankete označeni su s (P).")
    return text


def _w(widths):
    return [w * cm for w in widths]


def build_report(rows, unit, academic_year, semester=None, source="isvu", report_date=None, period_info=None,
                 unit_table=None, extra_units=None):
    """PDF izvješća jedne jedinice. rows: već odabrani retci (select).
    unit_table: za Sveučilište - [(sastavnica, stats)] za tablicu 12 i zaključak.
    extra_units: sastavnice bez pojedinačnih odgovora (npr. HVU iz zbirnog izvoza): [(naziv, anketiranja,
    ankete, M B1-B9, M B10)] - samo u tablici 12."""
    report_date = report_date or date.today()
    st = _styles()
    s = unit_stats(rows)
    university = unit_table is not None
    story = _cover(st, "Izvješće o rezultatima studentskog vrednovanja nastave", unit, academic_year, semester,
                   source, report_date)
    t = 0

    story += [Paragraph("1. Uvod", st["h1"]), Paragraph(_method_text(source, semester, university), st["body"]),
              Paragraph("Anketiranje je jedno pristupanje studenta anketi za pojedini predmet (na papirnatoj anketi "
                        "svaki je listić jedno anketiranje). Anketa je jedan ispunjeni anketni list kojim je student "
                        "procijenio rad jednog nastavnika na predmetu. Identifikacijski ključ zaseban je kod za svaku "
                        "kombinaciju predmet / nastavnik.", st["body"])]

    # --- Tablica A ---
    story.append(Paragraph("2. Rezultati ankete", st["h1"]))
    story.append(Paragraph("2.1 Informacije o anketiranoj jedinici", st["h2"]))
    info = [["Pokazatelj", "Vrijednost"], ["Anketirana jedinica", Paragraph(_esc(unit), st["cell"])],
            ["Razdoblje", period_label(academic_year, semester)]]
    if period_info:
        info.append(["Provedba ankete", period_info])
    info += [["Broj anketiranja", fmt_int(s["anketiranja"])], ["Broj anketa", fmt_int(s["ankete"])],
             ["Broj nastavnika", fmt_int(s["nastavnika"])], ["Broj predmeta", fmt_int(s["predmeta"])],
             ["Broj identifikacijskih ključeva", fmt_int(len(s["keys"]))]]
    if university:
        info.append(["Broj sastavnica", fmt_int(len(unit_table) + len(extra_units or []))])
    ta = _table(info, _w([6.5, 8]), total_row=False)
    ta.setStyle(TableStyle([("ALIGN", (1, 0), (1, -1), "LEFT")]))
    story += [Paragraph("Tablica A. Informacije o anketiranoj jedinici", st["caption"]), ta]

    # --- Tablice 1-5 (A dio) ---
    story.append(Paragraph("2.2 Opći podaci o anketiranim studentima", st["h2"]))
    for col, title, _cats in A_ITEMS:
        t += 1
        a = s["A"][col]
        data = [["Odgovor", "N", "%"]] + [[lbl, fmt_int(n), _pct(n, a["total"])] for lbl, n in a["counts"]]
        data += [["Bez odgovora", fmt_int(a["missing"]), _pct(a["missing"], a["total"])],
                 ["Ukupno", fmt_int(a["total"]), _pct(a["total"], a["total"])]]
        story.append(KeepTogether([Paragraph(f"Tablica {t}. {_esc(title)}", st["caption"]),
                                   _table(data, _w([6, 2, 2]))]))
    story.append(PageBreak())

    # --- Tablica 6: opisna statistika B1-B9 ---
    story.append(Paragraph("2.3 Procjena kvalitete rada nastavnika", st["h2"]))
    t += 1
    sem_means = {}
    if not semester and not rows.empty:
        for sem_name in ("Zimski", "Ljetni"):
            part = rows[rows["Semestar"] == sem_name]
            if not part.empty:
                sem_means[sem_name] = unit_stats(part)["B"]
    head = ["Pitanje", "N", "NN", "M", "SD", "D", "Min", "Max"] + [f"M {k[:1]}S" for k in sem_means]
    data = [head]
    for c, text in zip(B9, B_TEXTS):
        d = s["B"][c]
        data.append([Paragraph(_esc(text), st["cell"]), fmt_int(d["N"]), fmt_int(d["NN"]), fmt(d["M"], 2),
                     fmt(d["SD"], 2), fmt_int(d["D"]), fmt_int(d["Min"]), fmt_int(d["Max"])]
                    + [fmt(v[c]["M"], 2) for v in sem_means.values()])
    d = s["B9"]
    data.append([Paragraph("Ukupno (B.1-B.9)", st["cell_bold"]), fmt_int(d["N"]), fmt_int(d["NN"]), fmt(d["M"], 2),
                 fmt(d["SD"], 2), fmt_int(d["D"]), fmt_int(d["Min"]), fmt_int(d["Max"])]
                + ["" for _ in sem_means])
    widths = [7.0] + [1.15] * 7 + [1.15] * len(sem_means)
    story += [Paragraph(f"Tablica {t}. Procjena nastavnika (B.1-B.9) - opisna statistika", st["caption"]),
              _table(data, _w([widths[0] - 1.15 * len(sem_means)] + widths[1:]), font_size=9),
              Paragraph("N = broj valjanih odgovora (1-5), NN = bez odgovora ili \"ne mogu procijeniti\", "
                        "M = aritmetička sredina, SD = standardna devijacija, D = raspon (max - min)"
                        + (", M ZS / M LS = aritmetička sredina zimskog / ljetnog semestra" if sem_means else "")
                        + ".", st["small"])]

    # --- Tablica 7a/7b: distribucija B1-B9 ---
    t += 1
    labels = ["Bez odgovora", "Ne mogu procijeniti"] + SCALE_LABELS
    def col_counts(c):
        d = s["B"][c]
        return [d["blank"], d["cant"]] + d["counts"]
    matrix = np.array([col_counts(c) for c in B9]).T  # redci = odgovori
    head = ["Odgovor"] + [f"B.{i}" for i in range(1, 10)] + ["Ukupno"]
    n_rows = [head] + [[labels[r]] + [fmt_int(v) for v in matrix[r]] + [fmt_int(matrix[r].sum())]
                       for r in range(len(labels))]
    n_rows.append(["Ukupno"] + [fmt_int(v) for v in matrix.sum(axis=0)] + [fmt_int(matrix.sum())])
    nf = s["ankete"]
    p_rows = [head] + [[labels[r]] + [_pct(v, nf) for v in matrix[r]] + [_pct(matrix[r].sum(), nf * 9)]
                       for r in range(len(labels))]
    p_rows.append(["Ukupno"] + [_pct(nf, nf)] * 10)
    w = _w([3.3] + [1.18] * 9 + [1.5])
    story += [KeepTogether([Paragraph(f"Tablica {t}a. Distribucija odgovora B.1-B.9 (N)", st["caption"]),
                            _table(n_rows, w, font_size=8.5)]), Spacer(1, 0.3 * cm),
              KeepTogether([Paragraph(f"Tablica {t}b. Distribucija odgovora B.1-B.9 (%)", st["caption"]),
                            _table(p_rows, w, font_size=8.5)]), PageBreak()]

    # --- Tablica 8 i 9: B10 ---
    t += 1
    d = s["B"]["B10"]
    t8 = [["Pitanje", "N", "NN", "M", "SD", "D", "Min", "Max"],
          [Paragraph(_esc(B_TEXTS[9]), st["cell"]), fmt_int(d["N"]), fmt_int(d["NN"]), fmt(d["M"], 2),
           fmt(d["SD"], 2), fmt_int(d["D"]), fmt_int(d["Min"]), fmt_int(d["Max"])]]
    story += [Paragraph(f"Tablica {t}. Opća ocjena nastavnika (B.10) - opisna statistika", st["caption"]),
              _table(t8, _w([7.0] + [1.15] * 7), total_row=False, font_size=9)]
    t += 1
    b10_labels = ["Bez odgovora", "Ne mogu procijeniti"] + [lbl for _v, lbl in base.V4_B10_SCALE]
    vals = [d["blank"], d["cant"]] + d["counts"]
    t9 = [["Odgovor", "N", "%"]] + [[lbl, fmt_int(v), _pct(v, nf)] for lbl, v in zip(b10_labels, vals)]
    t9.append(["Ukupno", fmt_int(nf), _pct(nf, nf)])
    story += [Spacer(1, 0.3 * cm), Paragraph(f"Tablica {t}. Opća ocjena nastavnika (B.10) - distribucija odgovora",
                                             st["caption"]), _table(t9, _w([5.5, 2, 2]))]

    # --- Tablica 10: složeni pokazatelji ---
    t += 1
    t10 = [["Pokazatelj", "N", "M", "SD"]] + [[Paragraph(_esc(name), st["cell"]), fmt_int(d["N"]), fmt(d["M"], 2),
                                                fmt(d["SD"], 2)] for name, d in s["composites"]]
    story += [Spacer(1, 0.3 * cm), Paragraph(f"Tablica {t}. Složeni pokazatelji", st["caption"]),
              _table(t10, _w([9, 1.6, 1.6, 1.6]), total_row=True, font_size=9),
              Paragraph("Pokazatelj je prosjek odgovora na tri pitanja na jednoj anketi; računa se samo za ankete s "
                        "odgovorom na sva tri pitanja (za ukupno: na svih devet).", st["small"])]

    # --- Tablica 11: raspon prosjeka nastavnik / predmet ---
    t += 1
    keys = s["keys"]
    t11 = [["Prosjek po ključu", "Uvjet", "Ključeva", "M", "SD", "Min", "Q1", "Medijan", "Q3", "Max"]]
    for label, col in (("B.1-B.9", "M9"), ("B.10", "M10")):
        for min_n in (1, MIN_RANK):
            v = keys.loc[keys["N"] >= min_n, col].dropna()
            q = v.quantile([0.25, 0.5, 0.75]) if len(v) else pd.Series([np.nan] * 3, index=[0.25, 0.5, 0.75])
            t11.append([label, f"N ≥ {min_n}", fmt_int(len(v)), fmt(v.mean(), 2), fmt(v.std(ddof=0), 2),
                        fmt(v.min(), 2), fmt(q[0.25], 2), fmt(q[0.5], 2), fmt(q[0.75], 2), fmt(v.max(), 2)])
    story += [Spacer(1, 0.3 * cm),
              Paragraph(f"Tablica {t}. Raspon prosječnih ocjena na razini predmet / nastavnik", st["caption"]),
              _table(t11, _w([2.8, 1.5, 1.5, 1.2, 1.2, 1.2, 1.2, 1.5, 1.2, 1.2]), total_row=False, label_cols=2,
                     font_size=8.5), PageBreak()]

    # --- Slike ---
    means = [s["B"][c]["M"] for c in B_ALL]
    labels = [f"B.{i}" for i in range(1, 11)]
    story += [Paragraph("Slika 1. Prosječne ocjene po pitanjima (B.1-B.10)", st["caption"]),
              hbar_chart(labels[::-1], [(means[::-1], BAR, "")], x_max=5, decimals=2, height_cm=9)]
    total_valid = sum(sum(s["B"][c]["counts"]) for c in B9)
    dist = [sum(s["B"][c]["counts"][k] for c in B9) / total_valid * 100 if total_valid else 0 for k in range(5)]
    story += [Spacer(1, 0.5 * cm),
              Paragraph("Slika 2. Distribucija svih valjanih odgovora B.1-B.9 (%)", st["caption"]),
              hbar_chart([f"{lbl}" for _v, lbl in base.V4_B_SCALE], [(dist, BAR, "")], suffix=" %", height_cm=6.5)]

    # --- Sveučilište: rezultati po sastavnicama i zaključak ---
    if university:
        t += 1
        story.append(PageBreak())
        story.append(Paragraph("2.4 Rezultati po anketiranim sastavnicama", st["h2"]))
        per = [(name, us["anketiranja"], us["ankete"], us["B9"]["M"], us["B"]["B10"]["M"]) for name, us in unit_table]
        per += list(extra_units or [])
        per.sort(key=lambda r: r[0])
        t12 = [["Sastavnica", "Anketiranja", "Ankete", "M B.1-B.9", "M B.10"]]
        t12 += [[Paragraph(_esc(n), st["cell"]), fmt_int(a), fmt_int(f), fmt(m9, 2), fmt(m10, 2)]
                for n, a, f, m9, m10 in per]
        t12.append(["Ukupno", fmt_int(sum(r[1] for r in per)), fmt_int(sum(r[2] or 0 for r in per)),
                    fmt(s["B9"]["M"], 2), fmt(s["B"]["B10"]["M"], 2)])
        story += [Paragraph(f"Tablica {t}. Rezultati po sastavnicama", st["caption"]),
                  _table(t12, _w([8.5, 2, 2, 1.8, 1.6]), font_size=8.5)]
        if extra_units:
            story.append(Paragraph("Sastavnice za koje ISVU nije dao pojedinačne odgovore prikazane su prema "
                                   "zbirnom izvozu i nisu uključene u ostale tablice.", st["small"]))
        ranked = sorted([r for r in per if r[3] == r[3] and r[3] is not None], key=lambda r: r[3])
        story += [PageBreak(), Paragraph("Slika 3. Prosječna ocjena B.1-B.9 po sastavnicama", st["caption"]),
                  hbar_chart([r[0] for r in ranked], [([r[3] for r in ranked], BAR, "")], x_max=5, decimals=2,
                             width_cm=17, height_cm=max(8, 0.62 * len(ranked) + 2))]
        if ranked:
            mean = s["B9"]["M"]
            above = [r for r in ranked if r[3] > mean]
            below = [r for r in ranked if r[3] < mean]
            story += [Paragraph("3. Zaključak", st["h1"]), Paragraph(
                f"Prosječna ocjena svih odgovora B.1-B.9 na razini Sveučilišta je {fmt(mean, 2)}, a opća ocjena "
                f"nastavnika (B.10) {fmt(s['B']['B10']['M'], 2)}. Od {len(ranked)} sastavnica, {len(above)} ima "
                f"prosjek iznad sveučilišnog, a {len(below)} ispod njega. Najviši prosjek ima "
                f"{_esc(ranked[-1][0])} ({fmt(ranked[-1][3], 2)}), a najniži {_esc(ranked[0][0])} "
                f"({fmt(ranked[0][3], 2)}).", st["body"])]

    return _build(story, f"Izvješće V4 - {unit} - {period_label(academic_year, semester)}"), s


def _build(story, title, pagesize=A4):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=pagesize, leftMargin=2 * cm, rightMargin=2 * cm, topMargin=2 * cm,
                            bottomMargin=2 * cm, title=title, author="Ured za upravljanje kvalitetom, " + UNIVERSITY)
    doc.build(story, onFirstPage=lambda c, d: None, onLaterPages=base._page_number)
    return buf.getvalue()


def _list_header(st, title, unit, academic_year, semester, source, note):
    return [Paragraph(_esc(title), st["title"]), Spacer(1, 0.2 * cm),
            Paragraph(f"{_esc(unit)} - {_esc(period_label(academic_year, semester))}", st["center"]),
            Paragraph(f"Način provedbe: {_esc(SOURCES[source])}", st["center"]), Spacer(1, 0.3 * cm),
            Paragraph(_esc(note), st["small"])]


def build_key(keys, unit, academic_year, semester=None, source="isvu"):
    """Ključ za identifikaciju nastavnika (keys iz assign_keys)."""
    st = _styles()
    story = _list_header(st, "Ključ za identifikaciju nastavnika", unit, academic_year, semester, source,
                         "Dokument je povjerljiv: povezuje identifikacijske ključeve iz izvješća s imenima "
                         "nastavnika i predmeta.")
    data = [["Ključ", "Nastavnik", "Predmet", "Ankete"]]
    data += [[k, Paragraph(_esc(n), st["cell"]), Paragraph(_esc(c), st["cell"]), fmt_int(cnt)]
             for k, n, c, cnt in keys[["Ključ", "Nastavnik", "Kolegij", "N"]].itertuples(index=False)]
    tab = _table(data, _w([2.2, 5.5, 7.5, 1.6]), total_row=False, label_cols=1, font_size=8.5)
    tab.setStyle(TableStyle([("ALIGN", (1, 1), (2, -1), "LEFT")]))
    story.append(tab)
    return _build(story, f"Ključ - {unit}")


def build_ranking(keys, unit, academic_year, semester=None, source="isvu", min_n=1):
    st = _styles()
    r = ranking(keys, min_n)
    story = _list_header(st, f"Rang nastavnika (najmanje {min_n} {'anketa' if min_n == 1 else 'anketa'})", unit,
                         academic_year, semester, source,
                         "Poredak prema prosjeku svih odgovora B.1-B.9 za nastavnika na predmetu (zaokruženo na dvije "
                         "decimale; jednaki prosjeci dijele mjesto). Dokument je povjerljiv.")
    data = [["Rang", "Ključ", "Nastavnik", "Predmet", "N", "M B.1-B.9", "M B.10"]]
    data += [[str(rk), k, Paragraph(_esc(n), st["cell"]), Paragraph(_esc(c), st["cell"]), fmt_int(cnt),
              fmt(m9, 2), fmt(m10, 2)]
             for rk, k, n, c, cnt, m9, m10 in r[["Rang", "Ključ", "Nastavnik", "Kolegij", "N", "M9",
                                                 "M10"]].itertuples(index=False)]
    if len(data) == 1:
        data.append(["", "", "Nema nastavnika s dovoljno anketa", "", "", "", ""])
    tab = _table(data, _w([1.2, 2.0, 4.6, 6.0, 1.1, 1.8, 1.6]), total_row=False, label_cols=1, font_size=8.5)
    tab.setStyle(TableStyle([("ALIGN", (1, 1), (3, -1), "LEFT")]))
    story.append(tab)
    return _build(story, f"Rang min {min_n} - {unit}")


def unit_files(rows, unit, academic_year, semester=None, source="isvu", report_date=None, period_info=None):
    """Sve datoteke jedne sastavnice: {ime datoteke: pdf}, statistika."""
    pdf, s = build_report(rows, unit, academic_year, semester, source, report_date, period_info)
    keys = assign_keys(s["keys"])
    period = period_file_part(academic_year, semester)
    stem = f"{base.safe_file_part(unit)} - V4 {SHORT_SOURCE[source]} - {period}"
    files = {f"{stem} - izvješće.pdf": pdf,
             f"{stem} - ključ.pdf": build_key(keys, unit, academic_year, semester, source),
             f"{stem} - rang min 1.pdf": build_ranking(keys, unit, academic_year, semester, source, 1),
             f"{stem} - rang min {MIN_RANK}.pdf": build_ranking(keys, unit, academic_year, semester, source, MIN_RANK)}
    return files, s


SHORT_SOURCE = {"isvu": "ISVU", "papir": "papir", "oboje": "ISVU i papir"}


def period_file_part(academic_year, semester):
    y = "-".join(re.findall(r"\d{4}", academic_year)) or academic_year
    return f"{'ZS' if semester == 'Zimski' else 'LJS' if semester == 'Ljetni' else 'ak. god.'} {y}"


def university_report(rows, academic_year, semester=None, source="isvu", report_date=None, extra_units=None):
    """Izvješće za cijelo Sveučilište (bez ključa i ranga)."""
    table = [(fac, unit_stats(part)) for fac, part in rows.groupby("Fakultet") if len(part)]
    pdf, s = build_report(rows, UNIVERSITY, academic_year, semester, source, report_date,
                          unit_table=table, extra_units=extra_units)
    name = f"{UNIVERSITY} - V4 {SHORT_SOURCE[source]} - {period_file_part(academic_year, semester)} - izvješće.pdf"
    return name, pdf, s
