"""
Izrada PDF izvještaja iz očitanih anketa.

Pravila izračuna (dogovoreno):
  - Postoci se računaju od UKUPNOG broja anketa, uključujući "Nedostaje"
    (pa stupac % uvijek daje 100,0).
  - D = raspon = max - min.
  - Redak "Ukupno" u opisnoj statistici računa se iz SVIH odgovora zajedno
    (M i SD svih odgovora B1-B9, ne prosjek prosjeka).
  - SD je uzoračka standardna devijacija (kao Excelov STDEV).
  - 'x', prazno ili vrijednost izvan ljestvice = "Nedostaje".
"""

import io
import os
from datetime import date

import numpy as np
import pandas as pd

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib import font_manager  # noqa: E402

from reportlab.lib import colors  # noqa: E402
from reportlab.lib.enums import TA_CENTER, TA_LEFT  # noqa: E402
from reportlab.lib.pagesizes import A4  # noqa: E402
from reportlab.lib.styles import ParagraphStyle  # noqa: E402
from reportlab.lib.units import cm  # noqa: E402
from reportlab.pdfbase import pdfmetrics  # noqa: E402
from reportlab.pdfbase.ttfonts import TTFont  # noqa: E402
from reportlab.platypus import (  # noqa: E402
    Image, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)

APP_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(APP_DIR, "report_assets")
LOGO = os.path.join(ASSETS, "logo.jpg")

MISSING_LABEL = "Nedostaje"

# Boje kao u dosadašnjim izvještajima
GRAY = colors.HexColor("#D9D9D9")
BLUE = colors.HexColor("#9BC2E6")
BAR = "#2E75B6"

# ------------------------------------------------------------------------------
# Fontovi (Times New Roman na Windowsu, inače Liberation Serif iz report_assets)
# ------------------------------------------------------------------------------

_FONT_CANDIDATES = [
    {  # Windows
        "Serif": r"C:\Windows\Fonts\times.ttf",
        "Serif-Bold": r"C:\Windows\Fonts\timesbd.ttf",
        "Serif-Italic": r"C:\Windows\Fonts\timesi.ttf",
        "Serif-BoldItalic": r"C:\Windows\Fonts\timesbi.ttf",
    },
    {  # Uz program (radi i na Linuxu / VPS-u)
        "Serif": os.path.join(ASSETS, "fonts", "LiberationSerif-Regular.ttf"),
        "Serif-Bold": os.path.join(ASSETS, "fonts", "LiberationSerif-Bold.ttf"),
        "Serif-Italic": os.path.join(ASSETS, "fonts", "LiberationSerif-Italic.ttf"),
        "Serif-BoldItalic": os.path.join(ASSETS, "fonts", "LiberationSerif-BoldItalic.ttf"),
    },
]
_fonts_ready = False


def _register_fonts():
    global _fonts_ready
    if _fonts_ready:
        return
    for cand in _FONT_CANDIDATES:
        if all(os.path.exists(p) for p in cand.values()):
            for name, path in cand.items():
                pdfmetrics.registerFont(TTFont(name, path))
            pdfmetrics.registerFontFamily(
                "Serif", normal="Serif", bold="Serif-Bold", italic="Serif-Italic", boldItalic="Serif-BoldItalic"
            )
            _fonts_ready = True
            return
    raise FileNotFoundError("Nije pronađen font za izvještaj (report_assets/fonts).")


def _chart_font():
    """Grafikoni: Calibri kao u Excelu ako postoji, inače zadani font."""
    for path in (r"C:\Windows\Fonts\calibri.ttf",):
        if os.path.exists(path):
            font_manager.fontManager.addfont(path)
            return font_manager.FontProperties(fname=path).get_name()
    return "DejaVu Sans"


# ------------------------------------------------------------------------------
# Oblikovanje brojeva (hrvatski decimalni zarez)
# ------------------------------------------------------------------------------

def fmt(value, decimals=1):
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return "–"
    return f"{value:.{decimals}f}".replace(".", ",")


def fmt_int(value):
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return "–"
    return str(int(round(value)))


# ------------------------------------------------------------------------------
# Statistika
# ------------------------------------------------------------------------------

def clean_column(series, allowed_values):
    """Pretvara stupac u brojeve; sve što nije dopušteni odgovor postaje NaN (Nedostaje).
    Vraća (očišćeni stupac, broj vrijednosti izvan ljestvice)."""
    num = pd.to_numeric(series, errors="coerce")
    allowed = set(float(v) for v in allowed_values)
    invalid = num.notna() & ~num.isin(allowed)
    return num.where(~invalid), int(invalid.sum())


def distribution(clean, categories):
    """Redci: Nedostaje + kategorije. N i % od ukupnog broja anketa (s Nedostaje)."""
    total = len(clean)
    rows = [(MISSING_LABEL, int(clean.isna().sum()))]
    for value, label in categories:
        rows.append((label, int((clean == value).sum())))
    return [(label, n, (n / total * 100) if total else np.nan) for label, n in rows], total


def describe(values):
    """N (valjanih), NN (nedostaje), M, SD, D, Min, Max."""
    valid = values.dropna()
    n = len(valid)
    return {
        "N": n,
        "NN": int(values.isna().sum()),
        "M": valid.mean() if n else np.nan,
        "SD": valid.std(ddof=1) if n > 1 else np.nan,
        "D": (valid.max() - valid.min()) if n else np.nan,
        "Min": valid.min() if n else np.nan,
        "Max": valid.max() if n else np.nan,
    }


# ------------------------------------------------------------------------------
# Elementi izvještaja
# ------------------------------------------------------------------------------

def _styles():
    _register_fonts()
    return {
        "title": ParagraphStyle("title", fontName="Serif-Bold", fontSize=15, leading=19, alignment=TA_CENTER),
        "cover_bold": ParagraphStyle("cover_bold", fontName="Serif-Bold", fontSize=14, leading=18, alignment=TA_CENTER),
        "cover_italic": ParagraphStyle("cover_italic", fontName="Serif-Italic", fontSize=12, leading=16, alignment=TA_CENTER),
        "section": ParagraphStyle("section", fontName="Serif-Bold", fontSize=12, leading=15, alignment=TA_CENTER,
                                  spaceBefore=6, spaceAfter=14),
        "caption": ParagraphStyle("caption", fontName="Serif-Italic", fontSize=11, leading=14, alignment=TA_CENTER,
                                  spaceBefore=4, spaceAfter=8),
        "cell": ParagraphStyle("cell", fontName="Serif", fontSize=10, leading=12, alignment=TA_LEFT),
        "cell_bold": ParagraphStyle("cell_bold", fontName="Serif-Bold", fontSize=10, leading=12, alignment=TA_LEFT),
        "note": ParagraphStyle("note", fontName="Serif-Bold", fontSize=11, leading=15, alignment=TA_CENTER, spaceBefore=18),
    }


def _table(data, col_widths, header_rows=1, total_row=True, label_cols=1, font_size=10):
    """Tablica u stilu dosadašnjih izvještaja: sivo zaglavlje i oznake, plave vrijednosti."""
    t = Table(data, colWidths=col_widths, hAlign="CENTER")
    n_rows = len(data)
    style = [
        ("FONTNAME", (0, 0), (-1, -1), "Serif"),
        ("FONTSIZE", (0, 0), (-1, -1), font_size),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 0), (-1, header_rows - 1), GRAY),
        ("BACKGROUND", (0, header_rows), (label_cols - 1, -1), GRAY),
        ("BACKGROUND", (label_cols, header_rows), (-1, -1), BLUE),
        ("ALIGN", (label_cols, 0), (-1, -1), "RIGHT"),
        ("LINEBELOW", (0, 0), (-1, -1), 1, colors.white),
        ("LINEAFTER", (0, 0), (-1, -1), 2, colors.white),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]
    if total_row:
        style.append(("FONTNAME", (0, n_rows - 1), (-1, n_rows - 1), "Serif-Bold"))
    t.setStyle(TableStyle(style))
    return t


def _bar_chart(labels, percents, width_cm=16, height_cm=None, suffix=""):
    """Vodoravni stupčasti grafikon postotaka (prvi redak = dno grafikona, kao u Excelu)."""
    font = _chart_font()
    height_cm = height_cm or max(4.5, 1.15 * len(labels) + 1.5)
    fig, ax = plt.subplots(figsize=(width_cm / 2.54, height_cm / 2.54), dpi=200)
    vals = [0 if (p is None or np.isnan(p)) else p for p in percents]
    y = np.arange(len(labels))
    ax.barh(y, vals, height=0.45, color=BAR)
    ax.set_yticks(y, labels, fontfamily=font, fontsize=8, color="#595959")
    top = max(vals) if vals else 0
    x_max = 10 if top <= 0 else min(100, np.ceil(top * 1.12 / 10) * 10)
    if top > 0 and x_max < top + 3:
        x_max = top + 5
    ax.set_xlim(0, x_max)
    ax.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: fmt(v, 1) + suffix))
    ax.tick_params(axis="x", labelsize=8, colors="#595959", length=0)
    ax.tick_params(axis="y", length=0)
    ax.grid(axis="x", color="#D9D9D9", linewidth=0.7)
    ax.set_axisbelow(True)
    for side in ("top", "right", "bottom"):
        ax.spines[side].set_visible(False)
    ax.spines["left"].set_color("#BFBFBF")
    for yi, v in zip(y, vals):
        ax.text(v + x_max * 0.01, yi, fmt(v, 1) + suffix, va="center", fontsize=8, fontfamily=font, color="#404040")
    for t in ax.get_xticklabels():
        t.set_fontfamily(font)
    return figure_image(fig, width_cm, height_cm)


def figure_image(fig, width_cm, height_cm, dpi=150):
    """Pretvara matplotlib grafikon u sliku za PDF (sažeta paleta boja = mala datoteka)."""
    from PIL import Image as PILImage
    fig.patch.set_edgecolor("#D9D9D9")
    fig.patch.set_linewidth(1.2)
    fig.tight_layout(pad=0.8)
    raw = io.BytesIO()
    fig.savefig(raw, format="png", dpi=dpi, edgecolor=fig.get_edgecolor())
    plt.close(fig)
    raw.seek(0)
    small = io.BytesIO()
    PILImage.open(raw).convert("RGB").quantize(colors=48).save(small, format="PNG", optimize=True)
    small.seek(0)
    img = Image(small, width=width_cm * cm, height=height_cm * cm)
    img.hAlign = "CENTER"
    return img


def _cover(st, title_lines, middle_bold, middle_italic, n_surveys, report_date, place="Zagreb"):
    story = []
    if os.path.exists(LOGO):
        logo = Image(LOGO, width=9.5 * cm, height=9.5 * cm * 159 / 420)
        logo.hAlign = "CENTER"
        story += [logo]
    story += [Spacer(1, 3.2 * cm)]
    story += [Paragraph(line, st["title"]) for line in title_lines]
    story += [Spacer(1, 1.6 * cm)]
    for line in middle_bold:
        story += [Paragraph(line, st["cover_bold"]), Spacer(1, 0.45 * cm)]
    story += [Spacer(1, 1.2 * cm)]
    story += [Paragraph(line, st["cover_italic"]) for line in middle_italic]
    story += [Spacer(1, 0.6 * cm), Paragraph(f"Broj anketa: {n_surveys}", st["cover_italic"])]
    story += [Spacer(1, 1.8 * cm)]
    story += [Paragraph("Ured za upravljanje kvalitetom", st["cover_bold"]),
              Paragraph("Sveučilište u Zagrebu", st["cover_bold"]), Spacer(1, 0.5 * cm),
              Paragraph(f"{place}, {report_date.day}.{report_date.month}.{report_date.year}.", st["cover_bold"]),
              PageBreak()]
    return story


def _esc(text):
    return (str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


# ------------------------------------------------------------------------------
# Izvještaj: Anketa za procjenu nastavnika (V4)
# ------------------------------------------------------------------------------

V4_A_ITEMS = [
    ("Spol", "spol anketiranih studenata (A.1)", [(1, "Muški"), (0, "Ženski")]),
    ("Prisutnost", "prisutnost studenata na nastavi (A.2)",
     [(1, "Rijetka"), (2, "Povremena (30-70%)"), (3, "Redovita (više od 70%)")]),
    ("Interes kolegija", "početni interes za sadržaj predmeta (A.3)", [(1, "Mali"), (2, "Srednji"), (3, "Veliki")]),
    ("Indeks", "najčešću ocjenu u toku dosadašnjeg studiranja (A.4)", [(v, str(v)) for v in (2, 3, 4, 5)]),
    ("Ocjena kolegija", "očekivanu ocjenu iz predmeta (A.5)", [(v, str(v)) for v in (2, 3, 4, 5)]),
]

V4_B_ITEMS = [
    ("B1", "B.1 Jasno definira ishode učenja i ono što očekuje od studenata."),
    ("B2", "B.2 Nastava je dobro strukturirana, vrijeme racionalno iskorišteno."),
    ("B3", "B.3 Jasno i razumljivo izlaže nastavne sadržaje."),
    ("B4", "B.4 Uporabom različitih nastavnih materijala podiže kvalitetu."),
    ("B5", "B.5 Metode, primjeri i zadaci olakšavaju postizanje ishoda učenja."),
    ("B6", "B.6 Ima dobre komunikacijske vještine (ugodna radna atmosfera)."),
    ("B7", "B.7 Motiviran je za rad i savjesno izvršava svoje obaveze."),
    ("B8", "B.8 Prema studentima se odnosi korektno i s poštovanjem."),
    ("B9", "B.9 Nastavu održava redovito i na vrijeme."),
]
V4_B_SCALE = [(1, "potpuno nezadovoljan (1)"), (2, "nezadovoljan (2)"), (3, "zadovoljan (3)"),
              (4, "vrlo zadovoljan (4)"), (5, "potpuno zadovoljan (5)")]
V4_B10 = ("B10", "B.10 Koju biste ocjenu dali ovom nastavniku u cjelini?")
V4_B10_SCALE = [(1, "Nedovoljno (1)"), (2, "Dovoljno (2)"), (3, "Dobro (3)"), (4, "Vrlo dobro (4)"), (5, "Izvrsno (5)")]

PRELIMINARY_NOTE = ("*Ovo je prijevremeni izvještaj bez usporedbe prosječne ocjene i distribucije odgovora. "
                    "Cjeloviti izvještaj bit će dostupan nakon obrade svih anketnih upitnika.")


def _stats_row(label_par, s):
    return [label_par, fmt_int(s["N"]), fmt_int(s["NN"]), fmt(s["M"], 2), fmt(s["SD"], 2),
            fmt_int(s["D"]), fmt_int(s["Min"]), fmt_int(s["Max"])]


def build_v4_report(df, info, semester, academic_year, report_date=None):
    """Vraća (pdf_bajtovi, upozorenja).

    df   : ankete jednog nastavnika na jednom kolegiju (stupci kao u CONFIG-u)
    info : {"Fakultet", "Kolegij", "Nastavnik", "Nositelj predmeta", "Identifikacijski ključ"}
    """
    st = _styles()
    report_date = report_date or date.today()
    warnings = []
    n = len(df)
    if n == 0:
        raise ValueError("Nema anketa za izvještaj.")

    def col(name, allowed):
        if name not in df.columns:
            warnings.append(f"Nedostaje stupac {name} - računa se kao 'Nedostaje'.")
            return pd.Series([np.nan] * n)
        clean, bad = clean_column(df[name], allowed)
        if bad:
            warnings.append(f"{name}: {bad} vrijednost(i) izvan ljestvice računa se kao 'Nedostaje'.")
        return clean

    story = []

    # --- Naslovnica ---
    middle_italic = []
    if str(info.get("Identifikacijski ključ", "")).strip():
        middle_italic.append(f"Identifikacijski ključ (predmet / nastavnik): {_esc(info['Identifikacijski ključ'])}")
    if str(info.get("Nositelj predmeta", "")).strip():
        middle_italic.append(f"Nositelj predmeta: {_esc(info['Nositelj predmeta'])}")
    story += _cover(
        st,
        ["Izvješće o provedenoj anketi za procjenu nastavnika", "Studentsko vrjednovanje nastave",
         f"{_esc(semester)} semestar {_esc(academic_year)}"],
        [_esc(info.get("Fakultet", "")), _esc(info.get("Kolegij", "")), _esc(info.get("Nastavnik", ""))],
        middle_italic, n, report_date,
    )

    # --- A. Opći podaci ---
    table_no = chart_no = 0
    for i, (name, topic, cats) in enumerate(V4_A_ITEMS):
        clean = col(name, [v for v, _ in cats])
        rows, total = distribution(clean, cats)
        table_no += 1
        chart_no += 1
        block = []
        if i == 0:
            block.append(Paragraph("A. Opći podaci o studentici / studentu (A1-A5)", st["section"]))
        data = [["Odgovor", "N", "%"]] + [[lbl, fmt_int(c), fmt(p, 1)] for lbl, c, p in rows]
        data.append(["Ukupno", fmt_int(total), fmt(100.0 if total else np.nan, 1)])
        t = _table(data, [5.2 * cm, 1.6 * cm, 1.6 * cm])
        t.setStyle(TableStyle([("FONTNAME", (2, 1), (2, -1), "Serif-Bold")]))
        block += [Paragraph(f"Tablica {table_no}. Distribucija odgovora s obzirom na {topic}", st["caption"]), t,
                  Spacer(1, 0.9 * cm), _bar_chart([r[0] for r in rows], [r[2] for r in rows]),
                  Paragraph(f"Grafikon {chart_no}. Distribucija odgovora s obzirom na {topic}", st["caption"])]
        story += [KeepTogether(block), PageBreak()]

    # --- B. Procjena nastavnika: opisna statistika B1-B9 ---
    b_clean = {name: col(name, [v for v, _ in V4_B_SCALE]) for name, _ in V4_B_ITEMS}
    table_no += 1
    stats_data = [["Pitanje", "N", "NN", "M", "SD", "D", "Min", "Max"]]
    for name, text in V4_B_ITEMS:
        stats_data.append(_stats_row(Paragraph(_esc(text), st["cell"]), describe(b_clean[name])))
    pooled = pd.concat(list(b_clean.values()), ignore_index=True)
    stats_data.append(_stats_row(Paragraph("Ukupno", st["cell_bold"]), describe(pooled)))
    widths = [7.4 * cm] + [1.2 * cm] * 7
    story += [Paragraph("B. Procjena nastavnika / nastavnice na zadanom predmetu (B1-B10)", st["section"]),
              Paragraph(f"Tablica {table_no}. Prosječna ocjena odgovora B1-B9 - Opisna statistika", st["caption"]),
              _table(stats_data, widths), PageBreak()]

    # --- Distribucija B1-B9 (N i %) ---
    table_no += 1
    names = [nm for nm, _ in V4_B_ITEMS]
    dist = {nm: distribution(b_clean[nm], V4_B_SCALE)[0] for nm in names}
    labels = [r[0] for r in dist[names[0]]]
    counts = np.array([[dist[nm][r][1] for nm in names] for r in range(len(labels))])
    totals_row = counts.sum(axis=0)
    grand = counts.sum()

    head = [["Odgovor"] + [f"B.{k}" for k in range(1, 10)] + ["Ukupno"]]
    data_n = head + [[labels[r]] + [str(v) for v in counts[r]] + [str(counts[r].sum())] for r in range(len(labels))]
    data_n.append(["Ukupno"] + [str(v) for v in totals_row] + [str(grand)])
    data_p = head + [[labels[r]] + [fmt(v / n * 100, 1) for v in counts[r]] + [fmt(counts[r].sum() / grand * 100, 1)]
                     for r in range(len(labels))]
    data_p.append(["Ukupno"] + [fmt(100.0, 1)] * 9 + [fmt(100.0, 1)])
    w = [4.2 * cm] + [1.12 * cm] * 9 + [1.5 * cm]
    overall_pct = [counts[r].sum() / grand * 100 for r in range(len(labels))]
    story += [
        Paragraph(f"Tablica {table_no}a. Prosječna ocjena odgovora B1-B9 - Distribucija odgovora (N)", st["caption"]),
        _table(data_n, w, font_size=9), Spacer(1, 0.4 * cm),
        Paragraph(f"Tablica {table_no}b. Prosječna ocjena odgovora B1-B9 - Distribucija odgovora (%)", st["caption"]),
        _table(data_p, w, font_size=8.5), Spacer(1, 0.4 * cm),
    ]
    chart_no += 1
    story += [_bar_chart(labels, overall_pct, height_cm=7.2),
              Paragraph(f"Grafikon {chart_no}. Prosječna ocjena odgovora B1-B9 - Distribucija svih odgovora",
                        st["caption"]), PageBreak()]

    # --- B10 ---
    b10_name, b10_text = V4_B10
    b10 = col(b10_name, [v for v, _ in V4_B10_SCALE])
    table_no += 1
    s10 = [["Pitanje", "N", "NN", "M", "SD", "D", "Min", "Max"],
           _stats_row(Paragraph(_esc(b10_text), st["cell"]), describe(b10))]
    t10 = _table(s10, widths, total_row=False)
    rows10, total10 = distribution(b10, V4_B10_SCALE)
    d10 = [["B.10 Odgovor", "N", "%"]] + [[lbl, fmt_int(c), fmt(p, 1)] for lbl, c, p in rows10]
    d10.append(["Ukupno", fmt_int(total10), fmt(100.0, 1)])
    td10 = _table(d10, [5.2 * cm, 1.6 * cm, 1.6 * cm])
    td10.setStyle(TableStyle([("FONTNAME", (2, 1), (2, -1), "Serif-Bold")]))
    chart_no += 1
    story += [
        Paragraph(f"Tablica {table_no}. {_esc(b10_text[5:])} (B.10) - Opisna statistika", st["caption"]), t10,
        Spacer(1, 0.6 * cm),
        Paragraph(f"Tablica {table_no + 1}. {_esc(b10_text[5:])} (B.10) - Distribucija odgovora", st["caption"]), td10,
        Spacer(1, 0.6 * cm),
        _bar_chart([r[0] for r in rows10], [r[2] for r in rows10], height_cm=6.8),
        Paragraph(f"Grafikon {chart_no}. {_esc(b10_text[5:])} (B.10) – Distribucija odgovora", st["caption"]),
        Paragraph(PRELIMINARY_NOTE, st["note"]),
    ]

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4, leftMargin=2.2 * cm, rightMargin=2.2 * cm, topMargin=2.2 * cm, bottomMargin=2.0 * cm,
        title=f"Izvješće - {info.get('Nastavnik', '')} - {info.get('Kolegij', '')}",
        author="Ured za upravljanje kvalitetom, Sveučilište u Zagrebu",
    )
    doc.build(story, onFirstPage=lambda c, d: None, onLaterPages=_page_number)
    return buf.getvalue(), warnings


def _page_number(canvas, doc):
    canvas.saveState()
    canvas.setFont("Serif", 9)
    canvas.setFillColor(colors.HexColor("#595959"))
    canvas.drawCentredString(A4[0] / 2, 1.1 * cm, str(doc.page))
    canvas.restoreState()


# ------------------------------------------------------------------------------
# Popis dostupnih izvještaja i pomoćne funkcije za aplikaciju
# ------------------------------------------------------------------------------

REPORTS = {
    "Procjena nastavnika (V4)": {
        "build": build_v4_report,
        # Ankete se grupiraju po ovim podacima: jedan izvještaj = jedan nastavnik na jednom kolegiju
        "group_by": ["Fakultet", "Kolegij", "Nastavnik", "Nositelj predmeta", "Identifikacijski ključ"],
        "file_name": ["Nastavnik", "Kolegij"],
        "needs_semester": True,
        "year_label": "Akademska godina",
        "year_offset": 0,
        "min_surveys": 0,
    },
}


def has_report(form_type):
    return form_type in REPORTS


def safe_file_part(text, max_len=70):
    cleaned = "".join("_" if c in '\\/:*?"<>|' else c for c in str(text))
    return " ".join(cleaned.split())[:max_len].strip() or "bez_naziva"


def report_file_name(form_type, info):
    parts = [safe_file_part(info.get(k, "")) for k in REPORTS[form_type]["file_name"]]
    return " - ".join(p for p in parts if p) + ".pdf"


def build_report(form_type, df, info, semester, academic_year, report_date=None):
    return REPORTS[form_type]["build"](df, info, semester, academic_year, report_date)


def default_academic_year(today=None, offset=0):
    today = today or date.today()
    start = (today.year if today.month >= 9 else today.year - 1) + offset
    return f"{start}./{start + 1}."


def default_semester(today=None):
    today = today or date.today()
    return "zimski" if today.month >= 9 or today.month <= 2 else "ljetni"


def group_surveys(form_type, df_all):
    """Dijeli sve spremljene ankete na skupine za izvještaje.
    Vraća (lista (info, df), broj anketa bez potrebnih podataka)."""
    keys = REPORTS[form_type]["group_by"]
    df = df_all.copy()
    for k in keys:
        if k not in df.columns:
            df[k] = ""
        df[k] = df[k].fillna("").astype(str).str.strip()
    required = [k for k in keys if k != "Identifikacijski ključ"]
    if not keys:
        return [], 0
    incomplete = (df[required] == "").any(axis=1)
    groups = []
    for values, part in df[~incomplete].groupby(keys, sort=True):
        groups.append((dict(zip(keys, values)), part.reset_index(drop=True)))
    return groups, int(incomplete.sum())


# Izvještaji za studije (preddiplomski, diplomski, integrirani) - u zasebnoj datoteci
import reports_study  # noqa: E402

for _form in reports_study.LEVELS:
    REPORTS[_form] = {
        "build": reports_study.make_builder(_form),
        "group_by": ["Fakultet", "Studijski program"],  # Jedan izvještaj = jedan studijski program
        "file_name": ["Studijski program"],
        "needs_semester": False,
        "year_label": "Akademska godina završetka studija",
        "year_offset": -1,  # Izvještaj se radi za prošlu akademsku godinu
        "min_surveys": reports_study.MIN_SURVEYS,
    }
