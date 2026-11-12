import io
import os
import re
from collections import Counter
import cv2
import numpy as np
import pandas as pd
import streamlit as st
try:
    import pymupdf as fitz  # PyMuPDF (novije verzije)
except ImportError:
    import fitz  # PyMuPDF (starije verzije)
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LTChar

import drive_storage as drive
import programi
import reports

# ==============================================================================
# 1. DEFINICIJA STUPACA I PRAVILA KODIRANJA (IDENTIČNO VAŠIM EXCEL TABLICAMA)
# ==============================================================================

CONFIG = {
    "Procjena nastavnika (V4)": {
        "pdf": "Anketa za procjenu nastavnika V4.pdf",
        "scan_mode": "A4_SINGLE",
        # Podaci koji se upisuju ručno (umjesto kodova s obrasca) - idu na naslovnicu izvještaja
        "info_fields": [
            {"name": "Fakultet"},
            {"name": "Kolegij"},
            {"name": "Nastavnik"},
            {"name": "Nositelj predmeta", "choices": ["ne", "da"]},
            {"name": "Identifikacijski ključ", "optional": True},
        ],
        "batch_label": ["Nastavnik", "Kolegij"],  # Dodaje se u ime serije u Google Driveu
        "columns": [
            "Spol", "Prisutnost", "Interes kolegija", "Indeks", "Ocjena kolegija",
            "B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8", "B9", "B10"
        ],
        # Format: (stranica_0_indeks, indeks_retka_u_pdfu, lista_vrijednosti_za_kruzice)
        "questions": [
            (0, 1, [0, 1]),                  # 1. Spol: Ž=0, M=1
            (0, 2, [1, 2, 3]),               # 2. Prisutnost: 1, 2, 3
            (0, 3, [1, 2, 3]),               # 3. Interes: 1, 2, 3
            (0, 4, [2, 3, 4, 5]),            # 4. Indeks: 2, 3, 4, 5
            (0, 5, [2, 3, 4, 5]),            # 5. Ocjena kolegija: 2, 3, 4, 5
            # Row 6 je primjer u uputama (preskačemo)
            (0, 7, [1, 2, 3, 4, 5, 'x']),    # B1
            (0, 8, [1, 2, 3, 4, 5, 'x']),    # B2
            (0, 9, [1, 2, 3, 4, 5, 'x']),    # B3
            (0, 10, [1, 2, 3, 4, 5, 'x']),   # B4
            (0, 11, [1, 2, 3, 4, 5, 'x']),   # B5
            (0, 12, [1, 2, 3, 4, 5, 'x']),   # B6
            (0, 13, [1, 2, 3, 4, 5, 'x']),   # B7
            (0, 14, [1, 2, 3, 4, 5, 'x']),   # B8
            (0, 15, [1, 2, 3, 4, 5, 'x']),   # B9
            (0, 16, [1, 2, 3, 4, 5]),        # B10
        ]
    },

    "Preddiplomski studij": {
        "pdf": "Anketa preddiplomski.pdf",
        "scan_mode": "A3_TWO_IMAGES",  # 1. slika: desno str.1 | 2. slika: lijevo str.2, desno str.3
        "auto_pages": True,  # A3 (dvije stranice na slici) ili A4 ispis (stranica po slici) - prepoznaje se samo
        "info_fields": [{"name": "Kod studijskog programa", "code": True, "optional": True}, {"name": "Fakultet"}, {"name": "Studijski program"}],
        "batch_label": ["Kod studijskog programa", "Studijski program"],
        # Kod studija koji student zacrni na 1. stranici (desno je samo tiskani primjer "kód 2209")
        "code_grids": [{"column": "Očitani kod", "page": 0, "rows": [1, 2, 3, 4], "part": "left", "program": True}],
        "columns": [
            "Spol\n(muško = 1)\n(žensko = 0)", "Ocjena preddipl", "Vrsta studiranja",
            "Prvi izbor", "Nastavak studija",
            "B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8", "B9", "B10",
            "b11", "b12", "b13", "b14", "b15",
            "C16", "C17", "C18", "C19", "C20", "C21", "C22", "C23",
            "D24", "D25", "D26", "D27", "D28", "D29", "D30", "D31", "D32", "D33",
            "D34", "D35", "D36", "D37", "D38", "D39", "D40",
            "E41", "E42", "E43", "E44", "E45",
            "e46", "e47", "e48", "e49", "e50", "e51", "e52",
            "e53e", "e54e", "e55e", "e56e", "e57e", "e58e",
            "F59", "F60", "F61", "F62", "F63", "F64", "F65", "F66",
            "f67", "f68", "f69", "f70", "odgovornost za studij", "ukupna ocjena"
        ],
        "questions": [
            # Stranica 1: spajanje okomitih radio-kružića pomoću liste redaka
            (0, [5, 6], [0, 1]),                            # 3. Spol (Ž=0 gore, M=1 dolje)
            (0, 7, [2, 2.5, 3, 3.5, 4, 4.5, 5]),            # 4. Prosječna ocjena
            (0, [8, 9], [1, 2]),                            # 5. Vrsta studiranja (MZOŠ=1, osobne=2)
            (0, [10, 11, 12], [1, 2, 3]),                   # 6. Prvi izbor (da=1, ugl=2, ne=3)
            (0, [13, 14], [1, 2]),                          # 7. Nastavak studija (1, 2)
            # B1-B14 na 1. stranici (retci 16 do 29)
            *[(0, r, [1, 2, 3, 4, 5, 'x']) for r in range(16, 30)],
            # Stranica 2: B15, C16-C23, D24-D40, E41-E50 (svih 36 redaka na 2. stranici: 0 do 35)
            *[(1, r, [1, 2, 3, 4, 5, 'x']) for r in range(0, 36)],
            # Stranica 3: E51-E58 (retci 0 do 7)
            *[(2, r, [1, 2, 3, 4, 5, 'x']) for r in range(0, 8)],
            # Redak 8 je primjer ljestvice (preskačemo), slijede F59-F66 (retci 9 do 16)
            *[(2, r, [1, 2, 3, 4, 5, 'x']) for r in range(9, 17)],
            # f67-f70 DA/NE (retci 17 do 20)
            *[(2, r, [1, 2]) for r in range(17, 21)],
            # 71. Odgovornost za studij (3 okomita kružića: retci 21, 22, 23)
            (2, [21, 22, 23], [1, 2, 3]),
            # Ukupna ocjena (redak 24)
            (2, 24, [1, 2, 3, 4, 5, 'x'])
        ]
    },

    "Diplomski studij": {
        "pdf": "Anketa diplomski.pdf",
        "scan_mode": "A4_TWO_PAGES",  # Dvije A4 slike po anketi: 1. stranica pa 2. stranica
        "info_fields": [{"name": "Kod studijskog programa", "code": True, "optional": True}, {"name": "Fakultet"}, {"name": "Studijski program"}],
        "batch_label": ["Kod studijskog programa", "Studijski program"],
        # 3. Kód diplomskog studija (lijevo) i kód preddiplomskog studija (desno) na 1. stranici
        "code_grids": [
            {"column": "Očitani kod", "page": 0, "rows": [1, 2, 3, 4], "part": "left", "program": True},
            {"column": "Očitani kod preddipl.", "page": 0, "rows": [1, 2, 3, 4], "part": "right"},
        ],
        "columns": [
            "Spol\n(muško = 1)\n(žensko = 0)", "Ocjena Dipl", "Ocjena Preddipl",
            "Prvi izbor", "Nastavak studija",
            "B1", "B2", "B3", "B4", "B5", "B6", "B7",
            "C8", "C9", "C10", "C11", "C12", "C13", "C14", "C15", "C16", "C17",
            "C18", "C19", "C20", "C21", "C22", "C23", "C24",
            "D25", "D26", "D27", "D28", "D29", "D30",
            "F31", "F32", "F33", "F34", "F35", "F36", "F37",
            "F38", "F39", "F40", "F41", "F42"
        ],
        "questions": [
            (0, 5, [1, 0]),                                 # 4. Spol (M=1 lijevo, Ž=0 desno)
            (0, 6, [2, 2.5, 3, 3.5, 4, 4.5, 5]),            # 5. Ocjena Dipl
            (0, 7, [2, 2.5, 3, 3.5, 4, 4.5, 5]),            # 6. Ocjena Preddipl
            (0, 8, [1, 2, 3]),                              # 7. Prvi izbor
            (0, [9, 10], [1, 2]),                           # 8. Nastavak studija
            # B1-B7 i C8-C16 na 1. stranici (retci 11 do 26)
            *[(0, r, [1, 2, 3, 4, 5, 'x']) for r in range(11, 27)],
            # Stranica 2: C17-C24, D25-D30, F31-F37 (retci 0 do 20)
            *[(1, r, [1, 2, 3, 4, 5, 'x']) for r in range(0, 21)],
            # F38-F40 DA/NE (retci 21, 22, 23)
            *[(1, r, [1, 2]) for r in range(21, 24)],
            # F41 Odgovornost (redak 24)
            (1, 24, [1, 2, 3]),
            # F42 Opća ocjena (redak 25)
            (1, 25, [1, 2, 3, 4, 5])
        ]
    },

    "Integrirani studij": {
        "pdf": "Anketa integrirani.pdf",
        "scan_mode": "A3_TWO_IMAGES",
        "auto_pages": True,
        "info_fields": [{"name": "Kod studijskog programa", "code": True, "optional": True}, {"name": "Fakultet"}, {"name": "Studijski program"}],
        "batch_label": ["Kod studijskog programa", "Studijski program"],
        # Kod studija koji student zacrni na 1. stranici (desno je samo tiskani primjer "kód 2209")
        "code_grids": [{"column": "Očitani kod", "page": 0, "rows": [1, 2, 3, 4], "part": "left", "program": True}],
        "columns": [
            "Spol\n(muško = 1)\n(žensko = 0)", "Ocjena studij", "Prvi izbor",
            *[f"B{i}" for i in range(1, 16)],
            *[f"C{i}" for i in range(16, 24)],
            *[f"D{i}" for i in range(24, 41)],
            *[f"E{i}" for i in range(41, 58)],
            *[f"F{i}" for i in range(58, 69)],
            "odgovornost za studij", "ukupna ocjena"
        ],
        "questions": [
            (0, 5, [1, 0]),                                 # 3. Spol (M=1, Ž=0)
            (0, 6, [2, 2.5, 3, 3.5, 4, 4.5, 5]),            # 4. Prosječna ocjena
            (0, 7, [1, 2, 3]),                              # 5. Prvi izbor
            # B1-B18 na 1. stranici (retci 8 do 25)
            *[(0, r, [1, 2, 3, 4, 5, 'x']) for r in range(8, 26)],
            # Stranica 2: C19-E56 (retci 0 do 37)
            *[(1, r, [1, 2, 3, 4, 5, 'x']) for r in range(0, 38)],
            # Stranica 3: E57, F58-F65 (retci 0 do 8)
            *[(2, r, [1, 2, 3, 4, 5, 'x']) for r in range(0, 9)],
            # F66-F68 DA/NE (retci 9, 10, 11)
            *[(2, r, [1, 2]) for r in range(9, 12)],
            # 69. Odgovornost (redak 12)
            (2, 12, [1, 2, 3]),
            # 70. Opća ocjena (redak 13)
            (2, 13, [1, 2, 3, 4, 5])
        ]
    }
}

# ==============================================================================
# 2. ČITANJE VEKTORSKIH KOORDINATA KRUŽIĆA IZ VAŠIH PDF PREDLOŽAKA
# ==============================================================================

@st.cache_resource
def load_pdf_template(pdf_path, dpi=200):
    """Renderira PDF stranice u slike za poravnanje i izvlači točne koordinate kružića."""
    doc = fitz.open(pdf_path)
    ref_images = []
    scale = dpi / 72.0

    for page in doc:
        pix = page.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY)
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width)
        ref_images.append(img)

    pages_rows = []
    for page_layout in extract_pages(pdf_path):
        w, h = page_layout.bbox[2], page_layout.bbox[3]
        bubbles = []
        for elem in page_layout:
            if isinstance(elem, LTTextContainer):
                for line in elem:
                    if hasattr(line, '__iter__'):
                        for ch in line:
                            if isinstance(ch, LTChar):
                                code = ord(ch.get_text())
                                if ('OMR' in ch.fontname or 'Wingdings' in ch.fontname) and code not in (0xf020, 0x20):
                                    # Pretvorba iz PDF koordinata (ishodište dolje-lijevo) u piksele slike (ishodište gore-lijevo)
                                    cx = ((ch.bbox[0] + ch.bbox[2]) / 2.0) * scale
                                    cy = (h - (ch.bbox[1] + ch.bbox[3]) / 2.0) * scale
                                    bubbles.append((cx, cy))

        # Grupiranje kružića u retke po Y koordinati
        bubbles.sort(key=lambda b: (b[1], b[0]))
        rows = []
        for b in bubbles:
            if not rows or abs(rows[-1][0][1] - b[1]) > (3.5 * scale):
                rows.append([b])
            else:
                rows[-1].append(b)
        for r in rows:
            r.sort(key=lambda b: b[0])
        pages_rows.append(rows)

    return ref_images, pages_rows, scale

# ==============================================================================
# 3. AUTOMATSKO PORAVNANJE SLIKE (HOMOGRAFIJA) I OMR DETEKCIJA ZACRNJENOSTI
# ==============================================================================

def align_to_reference(scan_bgr, ref_gray):
    """Poravnava skeniranu stranicu prema izvornoj PDF stranici pomoću ORB značajki."""
    scan_gray = cv2.cvtColor(scan_bgr, cv2.COLOR_BGR2GRAY)
    th, tw = ref_gray.shape
    scan_resized = cv2.resize(scan_gray, (tw, th))
    scan_bgr_resized = cv2.resize(scan_bgr, (tw, th))

    orb = cv2.ORB_create(4000)
    kp1, des1 = orb.detectAndCompute(scan_resized, None)
    kp2, des2 = orb.detectAndCompute(ref_gray, None)

    if des1 is None or des2 is None:
        return scan_bgr_resized

    matcher = cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=True)
    matches = sorted(matcher.match(des1, des2), key=lambda x: x.distance)[:400]

    if len(matches) < 15:
        return scan_bgr_resized

    src_pts = np.float32([kp1[m.queryIdx].pt for m in matches]).reshape(-1, 1, 2)
    dst_pts = np.float32([kp2[m.trainIdx].pt for m in matches]).reshape(-1, 1, 2)

    M, _ = cv2.findHomography(src_pts, dst_pts, cv2.RANSAC, 5.0)
    if M is None:
        return scan_bgr_resized

    aligned_bgr = cv2.warpPerspective(scan_bgr_resized, M, (tw, th), borderValue=(255, 255, 255))
    return aligned_bgr

def bubble_radius(scale):
    return int(4.2 * scale)


def question_coords(pages_rows, page_idx, row_spec):
    """Koordinate kružića jednog pitanja (vodoravni redak, dio retka ili okomito složeni kružići)."""
    if isinstance(row_spec, tuple):  # (redak, od, do) - dio retka, npr. lijeva mreža za kod studija
        row, start, end = row_spec
        return pages_rows[page_idx][row][start:end]
    if isinstance(row_spec, list):
        return [pages_rows[page_idx][r][0] for r in row_spec]
    return pages_rows[page_idx][row_spec]


def code_questions(cfg):
    """Retci za znamenke koda studija (4 retka po kodu, kružići 0-9) - čitaju se ispred pitanja."""
    specs = []
    for grid in cfg.get("code_grids", []):
        start, end = (0, 10) if grid["part"] == "left" else (10, 20)
        specs += [(grid["page"], (r, start, end), list(range(10))) for r in grid["rows"]]
    return specs


def split_codes(cfg, answers):
    """Odvaja očitane kodove od odgovora. Kod je prazan ako neka znamenka nije jasno zacrnjena."""
    codes, i = [], 0
    for grid in cfg.get("code_grids", []):
        digits = answers[i:i + len(grid["rows"])]
        i += len(grid["rows"])
        codes.append("".join(str(d) for d in digits) if all(isinstance(d, int) for d in digits) else "")
    return codes, answers[i:]


def measure_ratios(aligned_pages, pages_rows, questions_spec, scale):
    """Mjeri zacrnjenost svakog kružića. Vraća listu (jedna lista omjera po pitanju).
    Ovo je jedini skupi korak - rezultat je malen pa se može čuvati za stotine anketa.

    Omjer = udio tamnih točaka + udio točaka u boji (plava kemijska). Tisak je crn, pa boja
    pouzdano odaje oznaku studenta i na tamnim fotokopijama gdje su prazni kružići "debeli"."""
    thresh_pages, ink_pages = [], []
    for p in aligned_pages:
        gray = cv2.cvtColor(p, cv2.COLOR_BGR2GRAY)
        # Adaptivni prag hvata i plavu i crnu kemijsku olovku
        th = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                   cv2.THRESH_BINARY_INV, 31, 10)
        thresh_pages.append(th)
        hsv = cv2.cvtColor(p, cv2.COLOR_BGR2HSV)
        ink = ((hsv[:, :, 1] > 70) & (hsv[:, :, 2] < 235)).astype(np.uint8) * 255
        # Papir u boji ili obojeni sken (većina stranice ima boju): boja tada ne govori ništa
        # pa se ne koristi i čitanje je isto kao kod crno-bijelog skena
        paper_tint = float(np.median(hsv[::8, ::8, 1]))
        ink_pages.append(ink if paper_tint < 30 and cv2.countNonZero(ink) < 0.15 * ink.size else None)

    radius = bubble_radius(scale)
    size = 2 * radius + 1
    disk = np.zeros((size, size), dtype=np.uint8)
    cv2.circle(disk, (radius, radius), radius, 255, -1)  # Ista maska kružića, samo jednom

    all_ratios = []
    for page_idx, row_spec, _values in questions_spec:
        th, ink = thresh_pages[page_idx], ink_pages[page_idx]
        h, w = th.shape
        ratios = []
        for (cx, cy) in question_coords(pages_rows, page_idx, row_spec):
            ix, iy = int(round(cx)), int(round(cy))
            # Mjerimo samo mali isječak oko kružića umjesto cijele stranice
            x0, y0 = max(ix - radius, 0), max(iy - radius, 0)
            x1, y1 = min(ix + radius + 1, w), min(iy + radius + 1, h)
            roi = th[y0:y1, x0:x1]
            m = disk[y0 - (iy - radius):y1 - (iy - radius), x0 - (ix - radius):x1 - (ix - radius)]
            total_pixels = cv2.countNonZero(m)
            marked_pixels = cv2.countNonZero(cv2.bitwise_and(roi, roi, mask=m))
            if ink is not None:
                ink_roi = ink[y0:y1, x0:x1]
                marked_pixels += cv2.countNonZero(cv2.bitwise_and(ink_roi, ink_roi, mask=m))
            ratios.append(marked_pixels / float(total_pixels) if total_pixels else 0.0)
        all_ratios.append(ratios)
    return all_ratios


def decide_marked(ratios, fill_threshold):
    """Koji su kružići u retku zacrnjeni? (brzo - ponavlja se kod promjene osjetljivosti)"""
    max_r = max(ratios)
    min_r = min(ratios)
    # Kružić mora biti osjetno tamniji od praznih kružića u istom retku
    return [
        i for i, r in enumerate(ratios)
        if r > (min_r + 0.12) and r >= max(fill_threshold, max_r * 0.78)
    ]


def answers_from_ratios(all_ratios, questions_spec, fill_threshold):
    """Pretvara izmjerene omjere u odgovore. 'x' = prazno ili dvostruko zacrnjeno."""
    answers = []
    for ratios, (_page, _rows, values) in zip(all_ratios, questions_spec):
        marked = decide_marked(ratios, fill_threshold)
        answers.append(values[marked[0]] if len(marked) == 1 else 'x')
    return answers


def draw_debug(aligned_pages, pages_rows, questions_spec, scale, all_ratios, fill_threshold):
    """Crta prepoznate kružiće za vizualnu kontrolu (samo za anketu koja se pregledava)."""
    debug_pages = [p.copy() for p in aligned_pages]
    radius = bubble_radius(scale)
    for ratios, (page_idx, row_spec, _values) in zip(all_ratios, questions_spec):
        marked = decide_marked(ratios, fill_threshold)
        for i, (cx, cy) in enumerate(question_coords(pages_rows, page_idx, row_spec)):
            ix, iy = int(round(cx)), int(round(cy))
            if len(marked) == 1 and i == marked[0]:
                cv2.circle(debug_pages[page_idx], (ix, iy), radius + 3, (0, 200, 0), 2)    # Zeleno = očitano
            elif i in marked:
                cv2.circle(debug_pages[page_idx], (ix, iy), radius + 3, (0, 140, 255), 2)  # Narančasto = više oznaka
            else:
                cv2.circle(debug_pages[page_idx], (ix, iy), radius, (200, 200, 200), 1)
    return debug_pages


def scan_units(scan_mode, files_sorted):
    """Grupira učitane slike u ankete: lista (oznaka, [datoteke])."""
    if scan_mode == "A4_SINGLE":
        return [(f.name, [f]) for f in files_sorted]
    if scan_mode in ("A3_TWO_IMAGES", "A4_TWO_PAGES"):
        return [(f"{a.name} + {b.name}", [a, b]) for a, b in zip(files_sorted[0::2], files_sorted[1::2])]
    return None  # Način skeniranja još nije podržan


PAGE_SIG_SIZE = (120, 170)
PAGE_ORB_WIDTH = 620


def page_signature(gray):
    """Mali, zamućeni otisak stranice za brzu usporedbu s predloškom (normalizirana korelacija)."""
    g = cv2.resize(gray, PAGE_SIG_SIZE, interpolation=cv2.INTER_AREA).astype(np.float32)
    g = cv2.GaussianBlur(g, (5, 5), 0)
    g -= g.mean()
    norm = np.linalg.norm(g)
    return g / norm if norm else g


def _orb_small(gray):
    small = cv2.resize(gray, (PAGE_ORB_WIDTH, int(gray.shape[0] * PAGE_ORB_WIDTH / gray.shape[1])),
                       interpolation=cv2.INTER_AREA)
    return cv2.ORB_create(1500).detectAndCompute(small, None)


def page_references(ref_images):
    """Otisci i ORB značajke predložaka - računaju se jednom po učitavanju."""
    return {"sigs": [page_signature(r) for r in ref_images], "orb": [_orb_small(r) for r in ref_images]}


def _orb_inliers(gray, refs):
    kp, des = _orb_small(gray)
    matcher = cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=True)
    counts = []
    for rkp, rdes in refs["orb"]:
        if des is None or rdes is None:
            counts.append(0)
            continue
        matches = matcher.match(des, rdes)
        if len(matches) < 10:
            counts.append(0)
            continue
        src = np.float32([kp[m.queryIdx].pt for m in matches]).reshape(-1, 1, 2)
        dst = np.float32([rkp[m.trainIdx].pt for m in matches]).reshape(-1, 1, 2)
        _H, mask = cv2.findHomography(src, dst, cv2.RANSAC, 5.0)
        counts.append(int(mask.sum()) if mask is not None else 0)
    return counts


def _classify_gray(small, medium_fn, refs):
    """Koja je stranica obrasca na (dijelu) slike: (indeks stranice ili None, okrenuta za 180°).

    Brza usporedba otisaka kad je rezultat jasan (ravno skenirano: ~0,85-0,95 prema < 0,35);
    inače (nakošeno, smanjeni ispis) usporedba ORB točaka, koja ne ovisi o nagibu ni okretanju."""
    if small is None or small.size == 0 or float(small.std()) < 6:
        return None, False  # Prazna stranica (npr. poleđina A3 arka): kontrast ~0, a stranice obrasca > 20
    scores = []  # (sličnost, stranica, okrenuta)
    for rotated, img in ((False, small), (True, cv2.rotate(small, cv2.ROTATE_180))):
        sig = page_signature(img)
        scores += [(float((sig * ref).sum()), page, rotated) for page, ref in enumerate(refs["sigs"])]
    scores.sort(reverse=True)
    best, second = scores[0], scores[1]
    if best[0] >= 0.6 and best[0] - second[0] >= 0.25:
        return best[1], best[2]
    counts = _orb_inliers(medium_fn(), refs)
    order = sorted(range(len(counts)), key=lambda p: counts[p], reverse=True)
    top = counts[order[0]]
    runner_up = counts[order[1]] if len(order) > 1 else 0
    if top >= 40 and top >= 3 * max(runner_up, 1):
        return order[0], False  # Poravnanje (ORB) samo ispravi nagib i okretanje
    return None, False


def half(img, part):
    """Lijeva ('L') ili desna ('R') polovica A3 slike; None = cijela slika."""
    if part is None:
        return img
    w = img.shape[1]
    return img[:, :w // 2] if part == "L" else img[:, w // 2:]


def classify_scan(f, refs):
    """Stranice obrasca na jednoj slici: lista (dio slike, stranica, okrenuta).

    Položena slika (šira nego viša) je A3 arak s dvije stranice - svaka polovica se prepoznaje
    zasebno; uspravna slika je jedna A4 stranica. Neprepoznate/prazne polovice se izostavljaju."""
    data = np.frombuffer(f.getvalue(), np.uint8)
    small = cv2.imdecode(data, cv2.IMREAD_REDUCED_GRAYSCALE_8)
    if small is None:
        return []
    medium = {}

    def medium_part(part):
        if "img" not in medium:
            medium["img"] = cv2.imdecode(data, cv2.IMREAD_REDUCED_GRAYSCALE_4)
        return half(medium["img"], part)

    parts = ("L", "R") if small.shape[1] > small.shape[0] else (None,)
    found = []
    for part in parts:
        page, rotated = _classify_gray(half(small, part), lambda p=part: medium_part(p), refs)
        if page is not None:
            found.append((part, page, rotated))
    return sorted(found, key=lambda item: item[1])  # Naopako skeniran A3: stranice ipak redom


def group_pages(files_sorted, classes, n_pages):
    """Slaže prepoznate stranice u ankete. Nova anketa počinje stranicom 1 (ili kad se stranica
    ponovi / dođe ranija stranica). Vraća (ankete, napomene); stranica koja nedostaje je None."""
    units, notes, current = [], [], None
    for f, found in zip(files_sorted, classes):
        if not found:
            notes.append(f"**{f.name}** nije prepoznata kao stranica obrasca - preskočena.")
            continue
        for part, page, rotated in found:
            if current is None or page == 0 or current[page] is not None or \
                    any(current[p] is not None for p in range(page + 1, n_pages)):
                current = [None] * n_pages
                units.append(current)
            current[page] = (f, part, rotated)
    result = []
    for pages in units:
        names = list(dict.fromkeys(entry[0].name for entry in pages if entry))
        missing = [str(p + 1) for p, entry in enumerate(pages) if entry is None]
        if missing:
            notes.append(f"Anketa **{' + '.join(names)}**: nedostaje stranica {', '.join(missing)} "
                         "(njezini odgovori ostaju prazni - upišite ih ručno ako imate papir).")
        result.append((" + ".join(names), pages))
    rotated = list(dict.fromkeys(entry[0].name for pages in units for entry in pages if entry and entry[2]))
    if rotated:
        notes.append(f"Okrenute naopako (program ih je okrenuo): {', '.join(rotated[:10])}"
                     + (" …" if len(rotated) > 10 else ""))
    return result, notes


def scan_formats(units):
    """Koliko je anketa na A3, A4 ili miješano - za kratku obavijest nakon učitavanja."""
    counts = Counter()
    for _label, pages in units:
        kinds = {"A3" if entry[1] else "A4" for entry in pages if entry}
        counts["miješano" if len(kinds) > 1 else next(iter(kinds))] += 1
    return counts


def decode_image(f):
    return cv2.imdecode(np.frombuffer(f.getvalue(), np.uint8), cv2.IMREAD_COLOR)


def align_unit(scan_mode, unit_files, ref_images):
    """Učitava slike jedne ankete i poravnava svaku stranicu prema PDF predlošku."""
    if scan_mode == "A4_SINGLE":
        return [align_to_reference(decode_image(unit_files[0]), ref_images[0])]
    if scan_mode == "A4_TWO_PAGES":
        # Dvije A4 slike po anketi: 1. slika = Str. 1, 2. slika = Str. 2
        return [align_to_reference(decode_image(f), ref) for f, ref in zip(unit_files, ref_images)]
    if scan_mode == "AUTO_PAGES":
        # (slika, dio slike, okrenuta) po stranici; stranica koja nedostaje = prazan predložak
        pages, decoded = [], {}
        for entry, ref in zip(unit_files, ref_images):
            if entry is None:
                pages.append(cv2.cvtColor(ref, cv2.COLOR_GRAY2BGR))
                continue
            f, part, rotated = entry
            if id(f) not in decoded:
                decoded[id(f)] = decode_image(f)  # A3 unutarnja slika nosi dvije stranice
            img = half(decoded[id(f)], part)
            pages.append(align_to_reference(cv2.rotate(img, cv2.ROTATE_180) if rotated else img, ref))
        return pages
    # A3_TWO_IMAGES: 1. slika (desna polovica = Str. 1), 2. slika (lijevo = Str. 2, desno = Str. 3)
    img1, img2 = decode_image(unit_files[0]), decode_image(unit_files[1])
    w1, w2 = img1.shape[1], img2.shape[1]
    return [
        align_to_reference(img1[:, w1 // 2:], ref_images[0]),
        align_to_reference(img2[:, :w2 // 2], ref_images[1]),
        align_to_reference(img2[:, w2 // 2:], ref_images[2]),
    ]


# --- Ručni ispravci (npr. student je zaokružio umjesto zacrnio kružiće) ---

MISSING = "x"


def answer_columns(cfg):
    """Za svaki stupac s odgovorima: (naziv, stranica obrasca, dopuštene vrijednosti ili None za kod)."""
    cols = [(g["column"], g["page"], None) for g in cfg.get("code_grids", [])]
    for name, (page, _rows, values) in zip(cfg["columns"], cfg["questions"]):
        allowed = [str(v) for v in values if v != MISSING] + [MISSING]
        cols.append((name, page, allowed))
    return cols


def clean_answer(value, allowed):
    """Ujednačuje upisani odgovor: '3,5' -> '3.5', prazno/X -> 'x', kod -> 4 znamenke."""
    text = "" if value is None or (isinstance(value, float) and np.isnan(value)) else str(value).strip()
    if allowed is None:  # Kod studija
        digits = "".join(ch for ch in text if ch.isdigit())
        return digits.zfill(4) if digits and len(digits) <= 4 else text
    text = text.replace(",", ".")
    if text.lower() in ("", "x", "nan", "none"):
        return MISSING
    try:
        number = float(text)
        return str(int(number)) if number.is_integer() else str(number)
    except ValueError:
        return text


def answer_pattern(allowed):
    """Regex za tablicu: dopuštene vrijednosti (3.5 ili 3,5), x/X ili prazno."""
    parts = sorted({re.escape(a).replace(r"\.", "[.,]") for a in allowed if a != MISSING}, key=len, reverse=True)
    return r"^(?:" + "|".join(parts + ["[xX]", ""]) + r")$"


def typed_answer(text):
    """Natrag u broj za Excel i izvještaje ('x' ostaje 'x')."""
    try:
        number = float(text)
        return int(number) if number.is_integer() else number
    except (TypeError, ValueError):
        return text


def short_question(name):
    return " ".join(str(name).split())


def shrink_for_preview(img_bgr, max_width=1200):
    h, w = img_bgr.shape[:2]
    if w > max_width:
        img_bgr = cv2.resize(img_bgr, (max_width, int(h * max_width / w)), interpolation=cv2.INTER_AREA)
    return cv2.cvtColor(img_bgr, cv2.COLOR_BGR2RGB)

# ==============================================================================
# 4. KORISNIČKO SUČELJE (STREAMLIT WEB/DESKTOP APLIKACIJA)
# ==============================================================================

st.set_page_config(page_title="UNIZG OMR Čitač Anketa", layout="wide")
st.title("📋 OMR Čitač Anketa – Sveučilište u Zagrebu")

form_type = st.sidebar.selectbox("1. Odaberite vrstu obrasca:", list(CONFIG.keys()))
sensitivity = st.sidebar.slider("Osjetljivost na kemijsku olovku:", 0.10, 0.35, 0.18, 0.01)

# --- Google Drive (spremanje u oblak) ---
st.sidebar.markdown("---")
st.sidebar.subheader("☁️ Google Drive")
drive_service = None
if not drive.libraries_installed():
    st.sidebar.warning("Google biblioteke nisu instalirane. Pokrenite: `pip install -r requirements.txt`")
elif not os.path.exists(drive.CREDENTIALS_FILE):
    st.sidebar.info("Google Drive još nije podešen. Učitajte ključ programa (*google_credentials.json*) "
                    "ili pogledajte upute u **PODESAVANJE_GOOGLE_DRIVE.md**.")
    key_file = st.sidebar.file_uploader("Ključ programa (google_credentials.json)", type=["json"], key="cred_upload")
    if key_file is not None:
        problem = drive.save_credentials(key_file.getvalue())
        if problem:
            st.sidebar.error(problem)
        else:
            st.sidebar.success("Ključ je spremljen.")
            st.rerun()
else:
    try:
        drive_service = drive.get_service(interactive=False)  # Tiha prijava ako je već spremljena
    except Exception as e:
        st.sidebar.error(f"Greška pri spajanju na Google Drive: {e}")

    if drive_service is None and drive.SERVER_MODE:
        # Na poslužitelju nema preglednika: prijava u pregledniku korisnika pa kopiranje adrese
        login = st.session_state.get("web_login")
        if login is None:
            st.sidebar.caption("Prijava se obnavlja otprilike jednom tjedno.")
            if st.sidebar.button("🔑 Prijavi se u Google Drive"):
                try:
                    url, flow = drive.start_web_login()
                    st.session_state.web_login = {"url": url, "flow": flow}
                    st.rerun()
                except Exception as e:
                    st.sidebar.error(f"Prijava se ne može pokrenuti: {e}")
        else:
            st.sidebar.markdown(
                f"**1.** [Otvorite Google prijavu]({login['url']}), odaberite račun i odobrite pristup.  \n"
                "**2.** Preglednik zatim prikaže stranicu koja se ne može otvoriti (*localhost*) - to je u redu. "
                "Kopirajte cijelu adresu iz trake preglednika i zalijepite je ovdje:"
            )
            pasted = st.sidebar.text_input("Adresa iz preglednika", key="web_login_paste",
                                           placeholder="http://localhost:8765/?state=...&code=...")
            c1, c2 = st.sidebar.columns(2)
            if c1.button("Završi prijavu", type="primary", disabled=not pasted):
                try:
                    drive.finish_web_login(login["flow"], pasted)
                    st.session_state.pop("web_login", None)
                    st.session_state.pop("web_login_paste", None)
                    st.rerun()
                except Exception as e:
                    st.sidebar.error(f"Prijava nije uspjela: {e}. Pokušajte ponovno od koraka 1.")
            if c2.button("Odustani"):
                st.session_state.pop("web_login", None)
                st.rerun()
    elif drive_service is None:
        st.sidebar.caption("Prva prijava otvara preglednik u kojem odobrite pristup.")
        if st.sidebar.button("🔑 Prijavi se u Google Drive"):
            logged_in = False
            try:
                drive.get_service(interactive=True)
                logged_in = True
            except Exception as e:
                st.sidebar.error(f"Prijava nije uspjela: {e}")
            if logged_in:
                st.rerun()
    else:
        if "drive_email" not in st.session_state:
            st.session_state.drive_email = drive.account_email(drive_service)
            st.session_state.drive_root_link = drive.root_folder_link(drive_service)
        st.sidebar.success(f"Povezano: {st.session_state.drive_email or 'Google račun'}")
        if st.session_state.drive_root_link:
            st.sidebar.markdown(f"[📂 Otvori mapu *{drive.ROOT_FOLDER}*]({st.session_state.drive_root_link})")
        if st.sidebar.button("Odjava s Google Drivea"):
            drive.logout()
            for k in ("drive_email", "drive_root_link"):
                st.session_state.pop(k, None)
            st.rerun()

# --- Postavke PDF izvještaja (vrijede za sve izvještaje u ovom krugu anketiranja) ---
if reports.has_report(form_type):
    st.sidebar.markdown("---")
    st.sidebar.subheader("📄 Postavke izvještaja")
    report_def = reports.REPORTS[form_type]
    report_semester = None
    if report_def["needs_semester"]:
        _semesters = ["zimski", "ljetni"]
        report_semester = st.sidebar.selectbox(
            "Semestar", _semesters, index=_semesters.index(reports.default_semester()), key="report_semester"
        )
    report_year = st.sidebar.text_input(
        report_def["year_label"], value=reports.default_academic_year(offset=report_def["year_offset"]),
        key=f"report_year|{report_def['year_offset']}",
    ).strip()
    report_date = st.sidebar.date_input("Datum izvještaja", format="DD.MM.YYYY", key="report_date")

cfg = CONFIG[form_type]
if not os.path.exists(cfg["pdf"]):
    st.error(f"U mapi programa nedostaje izvorna PDF datoteka: **{cfg['pdf']}**")
    st.stop()

ref_images, pages_rows, scale = load_pdf_template(cfg["pdf"])

scan_spec = code_questions(cfg) + cfg["questions"]  # Znamenke koda studija pa pitanja
code_grids = cfg.get("code_grids", [])
code_columns = [g["column"] for g in code_grids]
CODE_FIELD = "Kod studijskog programa"


@st.cache_data
def load_programs():
    return programi.load()


PROGRAMS = load_programs()

scan_mode = cfg["scan_mode"]
if cfg.get("auto_pages"):
    # Obrazac je A3, ali ga neke sastavnice ispisuju na A4 - format se prepoznaje za svaku sliku
    scan_mode = "AUTO_PAGES"
UPLOAD_HINTS = {
    "A3_TWO_IMAGES": "A3 obrazac: parovi slika strogo redom, naslovna pa unutarnja",
    "A4_TWO_PAGES": "parovi slika redom, 1. pa 2. stranica",
    "A4_SINGLE": "jedna slika po anketi",
    "AUTO_PAGES": "A3 i A4 ispis mogu biti zajedno - sve slike redom kako su skenirane",
}
uploaded_files = st.file_uploader(
    f"2. Učitajte skenirane JPEG slike ({UPLOAD_HINTS.get(scan_mode, '')}):",
    type=["jpg", "jpeg", "png"], accept_multiple_files=True
)
if cfg.get("auto_pages"):
    if st.toggle("Bez prepoznavanja stranica (samo A3, parovi slika strogo po redu)", key=f"fixed|{form_type}",
                 help="Program inače sam prepoznaje je li slika A3 arak ili A4 stranica i koja je stranica "
                      "obrasca. Uključite samo ako prepoznavanje na nekim skenovima ne radi dobro."):
        scan_mode = cfg["scan_mode"]

files_sorted, units, measurements, all_rows, proc_key = [], [], [], [], None
if uploaded_files:
    files_sorted = sorted(uploaded_files, key=lambda f: f.name)
    file_ids = tuple(getattr(f, "file_id", None) or f"{f.name}:{f.size}" for f in files_sorted)
    if scan_mode == "AUTO_PAGES":
        # Prepoznavanje stranica radi se jednom po učitanim datotekama
        page_cache = st.session_state.get("omr_page_classes")
        if page_cache is None or page_cache["key"] != (form_type, file_ids):
            page_refs = page_references(ref_images)
            progress = st.progress(0.0, text="Prepoznajem stranice...")
            classes = []
            for i, f in enumerate(files_sorted, start=1):
                classes.append(classify_scan(f, page_refs))
                progress.progress(i / len(files_sorted), text=f"Prepoznajem stranice... {i}/{len(files_sorted)}")
            progress.empty()
            page_cache = {"key": (form_type, file_ids), "classes": classes}
            st.session_state.omr_page_classes = page_cache
        units, page_notes = group_pages(files_sorted, page_cache["classes"], len(ref_images))
        formats = scan_formats(units)
        if units:
            st.caption(f"Prepoznato **{len(units)}** anketa: " + ", ".join(
                f"{n} × {kind}" for kind, n in sorted(formats.items())) + ".")
        if page_notes:
            with st.expander(f"⚠️ Napomene o stranicama ({len(page_notes)})", expanded=len(page_notes) <= 5):
                for note in page_notes:
                    st.markdown("- " + note)
    else:
        units = scan_units(scan_mode, files_sorted)
        if units is None:
            st.error(f"Način skeniranja **{scan_mode}** za ovaj obrazac još nije podržan.")
            units = []
        elif scan_mode in ("A3_TWO_IMAGES", "A4_TWO_PAGES") and len(files_sorted) % 2:
            st.warning(f"Učitan je neparan broj slika ({len(files_sorted)}) - zadnja slika "
                       f"**{files_sorted[-1].name}** nema par i nije očitana. Provjerite nedostaje li neka stranica.")

    # Skupi korak (poravnanje + mjerenje kružića) radi se samo JEDNOM po učitanim datotekama.
    # Ispravci u tablici, pregled i promjena osjetljivosti koriste spremljena mjerenja,
    # a slike se ne čuvaju u memoriji.
    proc_key = (form_type, scan_mode, file_ids)
    cache = st.session_state.get("omr_cache")
    if cache is None or cache["key"] != proc_key:
        measurements = []
        if units:
            progress = st.progress(0.0, text="Očitavam ankete...")
            for i, (_label, unit_files) in enumerate(units, start=1):
                pages = align_unit(scan_mode, unit_files, ref_images)
                measurements.append(measure_ratios(pages, pages_rows, scan_spec, scale))
                progress.progress(i / len(units), text=f"Očitavam ankete... {i}/{len(units)}")
            progress.empty()
        cache = {"key": proc_key, "measurements": measurements}
        st.session_state.omr_cache = cache
        st.session_state.pop("omr_preview", None)
    measurements = cache["measurements"]

    for (label, _files), m in zip(units, measurements):
        codes, answers = split_codes(cfg, answers_from_ratios(m, scan_spec, sensitivity))
        all_rows.append([label] + codes + answers)

# --- 3. Podaci o anketiranju (idu na naslovnicu izvještaja i u ime serije) ---
st.subheader("3. Podaci o anketiranju")
info_fields = cfg["info_fields"]
has_code_field = any(f.get("code") for f in info_fields)
program_col = next((1 + i for i, g in enumerate(code_grids) if g.get("program")), None)  # Stupac u all_rows
read_codes = [row[program_col] for row in all_rows] if program_col else []


def info_key(name):
    return f"info|{form_type}|{name}"


def fill_program(code, entry):
    """Upisuje kod, fakultet i naziv programa iz popisa kodova (polja se i dalje mogu mijenjati)."""
    st.session_state[info_key(CODE_FIELD)] = code
    st.session_state[info_key("Fakultet")] = entry["fakultet"]
    st.session_state[info_key("Studijski program")] = entry["program"]


def on_code_typed():
    code = st.session_state.get(info_key(CODE_FIELD), "").strip()
    entry = programi.lookup(PROGRAMS, code, form_type) if re.fullmatch(r"\d{1,4}", code) else None
    if entry:
        fill_program(code.zfill(4), entry)


if has_code_field:
    found = Counter(c for c in read_codes if c)
    # Novi skenovi: polja se popunjavaju iz koda koji je zacrnjen na većini anketa
    if proc_key and st.session_state.get(f"autofill|{form_type}") != proc_key:
        st.session_state[f"autofill|{form_type}"] = proc_key
        if found:
            top = found.most_common(1)[0][0]
            entry = programi.lookup(PROGRAMS, top, form_type)
            if entry:
                fill_program(top, entry)
            else:
                st.session_state[info_key(CODE_FIELD)] = top
    if read_codes:
        unread = sum(1 for c in read_codes if not c)
        msg = "Kod studija očitan s anketa: " + (", ".join(f"**{c}** ({n})" for c, n in found.most_common(5)) or "nijedan")
        if unread:
            msg += f"; nije očitan na {unread} anketa (prazno ili nejasno)"
        st.caption(msg + ".")

    if PROGRAMS:
        search_options = programi.choices(PROGRAMS, form_type)
        search_key = f"search|{form_type}"

        def on_search():
            idx = st.session_state.get(search_key)
            if idx is not None:
                fill_program(*search_options[idx])
                st.session_state.pop(search_key, None)  # Tražilica se isprazni za sljedeće traženje

        st.selectbox(
            "🔎 Pronađite studijski program po kodu ili nazivu", range(len(search_options)), index=None,
            format_func=lambda i: programi.label(*search_options[i]),
            placeholder="Upišite kod ili dio naziva...", key=search_key, on_change=on_search,
        )

info_values = {}
for start in range(0, len(info_fields), 3):
    for col, field in zip(st.columns(3), info_fields[start:start + 3]):
        key = info_key(field["name"])
        label = field["name"] + (" (neobavezno)" if field.get("optional") else "")
        if field.get("choices"):
            info_values[field["name"]] = col.selectbox(label, field["choices"], key=key)
        elif field.get("code"):
            info_values[field["name"]] = col.text_input(
                field["name"], key=key, max_chars=4, on_change=on_code_typed,
                help="4 znamenke s 1. stranice obrasca. Fakultet i studijski program popunjavaju se iz popisa kodova.",
            ).strip()
        else:
            info_values[field["name"]] = col.text_input(label, key=key).strip()
missing_info = [f["name"] for f in info_fields if not f.get("optional") and not info_values[f["name"]]]

if has_code_field and info_values.get(CODE_FIELD) and PROGRAMS:
    chosen_code = info_values[CODE_FIELD]
    entry = programi.lookup(PROGRAMS, chosen_code, form_type)
    if entry is None:
        st.warning(f"Kod **{chosen_code}** nije u popisu kodova - provjerite ga ili upišite fakultet i program ručno.")
    elif entry["razina"] != form_type:
        st.warning(f"Kod **{chosen_code}** je u popisu kodova *{entry['razina'].lower()}* ({entry['program']}), "
                   f"a odabran je obrazac **{form_type}**.")

if uploaded_files:
    # --- Ispravci se čuvaju zasebno od očitanja (vrijede i kad se promijeni osjetljivost) ---
    edits = st.session_state.get("omr_edits")
    if edits is None or edits["key"] != proc_key:
        edits = {"key": proc_key, "cells": {}, "added": [], "deleted": set(), "table_ver": 0, "manual_ver": 0}
        st.session_state.omr_edits = edits
    cells = edits["cells"]
    ans_cols = answer_columns(cfg)
    allowed_of = {name: allowed for name, _page, allowed in ans_cols}
    col_names = [name for name, _page, _allowed in ans_cols]

    def build_table():
        """Očitanja + ručni ispravci, sve kao tekst ('3.5', 'x', '0061')."""
        rows, labels = [], []
        for i, row in enumerate(all_rows):
            if i in edits["deleted"]:
                continue
            rows.append([row[0]] + [clean_answer(v, allowed_of[c]) for c, v in zip(col_names, row[1:])])
            labels.append(i)
        for k, lab in enumerate(edits["added"]):
            if lab not in edits["deleted"]:
                rows.append([f"Ručno dodana anketa {k + 1}"] + [MISSING if allowed_of[c] else "" for c in col_names])
                labels.append(lab)
        df = pd.DataFrame(rows, columns=["Datoteka"] + col_names, index=labels)
        for (lab, col), value in cells.items():
            if lab in df.index and col in df.columns:
                df.at[lab, col] = value
        return df

    table_df = build_table()
    read_df = pd.DataFrame([[r[0]] + [clean_answer(v, allowed_of[c]) for c, v in zip(col_names, r[1:])] for r in all_rows],
                           columns=["Datoteka"] + col_names)  # Izvorno očitano (za usporedbu)
    question_cols = [name for name, _p, allowed in ans_cols if allowed]
    missing_counts = (table_df[question_cols] == MISSING).sum(axis=1) if question_cols else pd.Series(0, index=table_df.index)
    many_missing = missing_counts[missing_counts >= max(5, 0.3 * len(question_cols))]

    st.subheader("📊 Očitani podaci")
    st.caption("Ispravke možete upisati izravno u tablicu i spremiti gumbom **💾 Spremi izmjene u tablici**; "
               "za ankete s puno praznih odgovora koristite **✏️ Ručni unos** ispod tablice - uz sliku obrasca. "
               "*x* = nedostaje.")
    table_key = f"omr_table|{edits['table_ver']}"
    shown_labels = list(table_df.index)

    def on_table_edit():
        change = st.session_state.get(table_key, {})
        for pos, values in change.get("edited_rows", {}).items():
            lab = shown_labels[int(pos)]
            for col, value in values.items():
                if col in allowed_of:
                    cells[(lab, col)] = clean_answer(value, allowed_of[col])
        structural = bool(change.get("added_rows") or change.get("deleted_rows"))
        for new_row in change.get("added_rows", []):
            lab = len(all_rows) + len(edits["added"])
            edits["added"].append(lab)
            for col, value in new_row.items():
                if col in allowed_of:
                    cells[(lab, col)] = clean_answer(value, allowed_of[col])
        for pos in change.get("deleted_rows", []):
            edits["deleted"].add(shown_labels[int(pos)])
        edits["manual_ver"] += 1  # Ručni unos prikazuje nove vrijednosti
        edits["table_ver"] += 1  # Tablica se iscrtava sa spremljenim vrijednostima
        n_changes = sum(len(v) for v in change.get("edited_rows", {}).values())
        st.session_state.omr_saved_msg = ("table", f"Spremljeno: {n_changes} izmjena u tablici"
                                          + (", dodani/obrisani redci" if structural else "") + ".")

    column_config = {"Datoteka": st.column_config.TextColumn("Datoteka", disabled=True)}
    for name, _page, allowed in ans_cols:
        # Svi stupci kao tekst (kao stupac koda) - tablica prihvaća samo dopuštene odgovore
        column_config[name] = st.column_config.TextColumn(
            short_question(name), max_chars=4,
            validate=answer_pattern(allowed) if allowed else r"^\d{0,4}$",
            help=("Dopušteno: " + ", ".join(allowed)) if allowed else "Kod studija (4 znamenke)",
        )
    # Obrazac: dok upisujete, ništa se ne osvježava; sve se sprema tek gumbom
    with st.form(f"table_form|{edits['table_ver']}", border=False):
        st.data_editor(table_df.reset_index(drop=True).astype(object), num_rows="dynamic",
                       use_container_width=True, column_config=column_config, key=table_key)
        st.form_submit_button("💾 Spremi izmjene u tablici", on_click=on_table_edit)
    saved_msg = st.session_state.get("omr_saved_msg")
    if saved_msg and saved_msg[0] == "table":
        st.success(saved_msg[1])
        st.session_state.pop("omr_saved_msg", None)

    if len(many_missing):
        st.warning(
            f"**{len(many_missing)}** anketa ima puno praznih odgovora (možda su kružići zaokruženi, a ne "
            "zacrnjeni): " + ", ".join(f"{table_df.at[lab, 'Datoteka']} ({n})" for lab, n in many_missing.head(10).items())
            + (" …" if len(many_missing) > 10 else "") + ". Upišite ih u **✏️ Ručni unos** ispod."
        )

    # Upisane vrijednosti izvan ljestvice (npr. 7 kod pitanja 1-5)
    invalid = [(table_df.at[lab, "Datoteka"], short_question(col), table_df.at[lab, col])
               for lab in table_df.index for col in question_cols
               if table_df.at[lab, col] not in allowed_of[col]]
    if invalid:
        st.error("Neispravne vrijednosti (računaju se kao *nedostaje*): " + "; ".join(
            f"{f} – {q}: '{v}'" for f, q, v in invalid[:10]) + (" …" if len(invalid) > 10 else ""))

    # Ankete na kojima je zacrnjen drugi kod - možda su zalutale u ovu seriju
    if program_col and info_values.get(CODE_FIELD) and len(table_df):
        code_col = code_columns[program_col - 1]
        codes_now = table_df[code_col].fillna("").astype(str).str.strip()
        other = table_df[(codes_now != "") & (codes_now != info_values[CODE_FIELD])]
        if len(other):
            examples = "; ".join(f"{r['Datoteka']} → {r[code_col]}" for _, r in other.head(8).iterrows())
            st.warning(
                f"**{len(other)}** anketa ima drugačiji kod od {info_values[CODE_FIELD]}: {examples}"
                f"{' …' if len(other) > 8 else ''}. Provjerite pripadaju li ovoj seriji ili ispravite kod u tablici."
            )

    # --- ✏️ Ručni unos uz sliku obrasca (crta se samo odabrana anketa) ---
    if units:
        st.subheader("✏️ Ručni unos i vizualna provjera")
        flagged = set(many_missing.index)
        # Odabir (filtar, anketa, stranica) pamti se između osvježavanja - nakon upisa odgovora
        # ostajete na istoj anketi i stranici
        if st.session_state.get("manual_for") != proc_key:
            st.session_state.manual_for = proc_key
            st.session_state.manual_only_flagged = bool(flagged)
            st.session_state.pop("manual_sel", None)
            st.session_state.manual_snap = {}
        only_flagged = st.checkbox(f"Prikaži samo ankete s puno praznih odgovora ({len(flagged)})",
                                   key="manual_only_flagged")
        alive = [i for i in range(len(units)) if i not in edits["deleted"]]
        choices = [i for i in alive if not only_flagged or i in flagged]
        current = st.session_state.get("manual_sel")
        if current in alive and current not in choices:
            choices = sorted(choices + [current])  # Ispravljena anketa ostaje odabrana i kad više nije označena
        if not choices:
            choices = alive
        if current not in choices:
            st.session_state.pop("manual_sel", None)

        def unit_label(i):
            n = int(missing_counts.get(i, 0))
            return f"{'⚠️ ' if i in flagged else ''}{units[i][0]}  ({n} praznih)"

        sel = st.selectbox("Odaberite anketu:", choices, format_func=unit_label, key="manual_sel") if choices else None
        if sel is not None:
            preview_key = (proc_key, sel, sensitivity)
            preview = st.session_state.get("omr_preview")
            if preview is None or preview["key"] != preview_key:
                with st.spinner("Pripremam pregled..."):
                    pages = align_unit(scan_mode, units[sel][1], ref_images)
                    dbg = draw_debug(pages, pages_rows, scan_spec, scale, measurements[sel], sensitivity)
                    preview = {"key": preview_key, "images": [shrink_for_preview(d) for d in dbg]}
                st.session_state.omr_preview = preview
            st.caption("Zeleno = očitano, narančasto = više oznaka u retku. Pregledajte sve stranice i upišite "
                       "odgovore (1-5, ocjenu, 0/1 ...) ili *x* ako nema odgovora. **Ništa se ne sprema dok ne "
                       "kliknete 💾 Spremi** - ako prije toga odaberete drugu anketu, upisi se gube.")
            saved_msg = st.session_state.get("omr_saved_msg")
            if saved_msg and saved_msg[0] == "manual":
                st.success(saved_msg[1])
                st.session_state.pop("omr_saved_msg", None)

            form_id = f"{edits['manual_ver']}|{sensitivity}|{sel}"
            snaps = st.session_state.setdefault("manual_snap", {})
            page_editors = []  # (ključ tablice, nazivi pitanja)
            later = [i for i in choices if i > sel]
            next_sel = later[0] if later else None

            def save_manual(go_next=False, editors=page_editors, lab=sel, nxt=next_sel):
                n_changes = 0
                for key, names in editors:
                    for pos, values in st.session_state.get(key, {}).get("edited_rows", {}).items():
                        if "Odgovor" in values:
                            col = names[int(pos)]
                            cells[(lab, col)] = clean_answer(values["Odgovor"], allowed_of[col])
                            n_changes += 1
                edits["manual_ver"] += 1  # Ručni unos se iscrtava sa spremljenim vrijednostima
                edits["table_ver"] += 1   # Glavna tablica također
                msg = f"Spremljeno: {n_changes} odgovora za {units[lab][0]}."
                if go_next and nxt is not None:
                    st.session_state.manual_sel = nxt
                    msg += " Prikazana je sljedeća anketa."
                st.session_state.omr_saved_msg = ("manual", msg)

            def save_buttons(labels):
                c1, c2, _ = st.columns([1, 1.4, 2])
                c1.form_submit_button(labels[0], on_click=save_manual, type="primary")
                c2.form_submit_button(labels[1], on_click=save_manual, kwargs={"go_next": True},
                                      disabled=next_sel is None)

            # Sve stranice odabrane ankete u jednom obrascu: upisi se ne spremaju dok ne kliknete Spremi
            with st.form(f"manual_form|{form_id}", border=True):
                save_buttons(("💾 Spremi", "💾 Spremi i sljedeća anketa ▶"))
                for page, image in enumerate(preview["images"]):
                    st.markdown(f"**Stranica {page + 1}**")
                    img_col, edit_col = st.columns([3, 2])
                    img_col.image(image, use_container_width=True)
                    page_cols = [(name, allowed) for name, p, allowed in ans_cols if p == page]
                    if not page_cols:
                        continue
                    manual_key = f"manual|{form_id}|{page}"
                    if manual_key not in snaps:
                        snaps[manual_key] = pd.DataFrame({
                            "Pitanje": [short_question(n) for n, _a in page_cols],
                            "Očitano": [read_df.at[sel, n] for n, _a in page_cols],
                            "Odgovor": [table_df.at[sel, n] for n, _a in page_cols],
                            "Dopušteno": [", ".join(a) if a else "4 znamenke" for _n, a in page_cols],
                        }).astype(object)
                    page_editors.append((manual_key, [n for n, _a in page_cols]))
                    edit_col.data_editor(
                        snaps[manual_key], hide_index=True, use_container_width=True, key=manual_key,
                        height=min(36 * (len(page_cols) + 1) + 4, 1400),
                        column_config={
                            "Pitanje": st.column_config.TextColumn(disabled=True),
                            "Očitano": st.column_config.TextColumn(disabled=True, width="small"),
                            "Odgovor": st.column_config.TextColumn(width="small"),
                            "Dopušteno": st.column_config.TextColumn(disabled=True),
                        },
                    )
                save_buttons(("💾 Spremi ispravke", "💾 Spremi ispravke i idi na sljedeću ▶"))

    # Ručni ispravci ulaze u Excel, izvještaj i Google Drive (brojevi kao brojevi, 'x' = nedostaje)
    edited_df = table_df.copy()
    for col in question_cols:
        edited_df[col] = edited_df[col].map(typed_answer)
    edited_df = edited_df.reset_index(drop=True)

    # Upisani nazivi dodaju se kao prvi stupci svake ankete u ovoj seriji
    final_df = edited_df.copy()
    for pos, (field, value) in enumerate(info_values.items()):
        final_df.insert(pos, field, value)

    # Gumb za spremanje u Excel (.xlsx) - izrađuje se u memoriji, bez datoteke na disku
    st.download_button(
        "💾 Preuzmi gotovu Excel tablicu (.xlsx)",
        data=drive.dataframe_to_xlsx_bytes(final_df.drop(columns=["Datoteka"])),
        file_name=f"{form_type}_rezultati.xlsx",
        mime=drive.XLSX_MIME
    )

    # PDF izvještaj za ovu seriju
    if reports.has_report(form_type) and all_rows:
        report_key = (form_type, final_df.to_csv(index=False), report_semester, report_year, str(report_date))
        stored = st.session_state.get("batch_report")
        if missing_info:
            st.caption(f"Za PDF izvještaj upišite: {', '.join(missing_info)}.")
        if st.button("📄 Izradi PDF izvještaj", disabled=bool(missing_info) or not report_year):
            try:
                with st.spinner("Izrađujem izvještaj..."):
                    pdf, report_warnings = reports.build_report(
                        form_type, final_df, info_values, report_semester, report_year, report_date
                    )
                stored = {"key": report_key, "pdf": pdf, "warnings": report_warnings,
                          "name": reports.report_file_name(form_type, info_values)}
                st.session_state.batch_report = stored
            except Exception as e:
                st.error(f"Izrada izvještaja nije uspjela: {e}")
        if stored and stored["key"] == report_key:
            for w in stored["warnings"]:
                st.warning(w)
            st.download_button("⬇️ Preuzmi PDF izvještaj", data=stored["pdf"], file_name=stored["name"],
                               mime="application/pdf")

    # Spremanje serije u Google Drive
    if drive_service is not None and all_rows:
        st.subheader("☁️ Spremanje u Google Drive")
        batch_key = form_type + "|" + "|".join(f"{f.name}:{f.size}" for f in files_sorted)
        saved_batches = st.session_state.setdefault("saved_batches", {})
        previous = saved_batches.get(batch_key)

        save_scans = st.checkbox("Spremi i izvorne skenove", value=True, disabled=previous is not None)
        button_label = "☁️ Spremi izmjene u Google Drive" if previous else "☁️ Spremi u Google Drive"
        if missing_info:
            st.warning(f"Prije spremanja upišite: **{', '.join(missing_info)}** (gore, pod točkom 2).")
        if st.button(button_label, type="primary", disabled=bool(missing_info)):
            try:
                with st.spinner("Spremam u Google Drive..."):
                    scans = None
                    if save_scans and previous is None:
                        scans = [(f.name, f.getvalue()) for f in files_sorted]
                    saved_batches[batch_key] = drive.save_batch(
                        drive_service, form_type, final_df, scans, previous=previous,
                        label=" - ".join(info_values[k] for k in cfg.get("batch_label", []) if info_values.get(k)),
                    )
                    st.session_state.drive_root_link = drive.root_folder_link(drive_service)
            except Exception as e:
                st.error(f"Spremanje nije uspjelo: {e}")

        info = saved_batches.get(batch_key)
        if info:
            msg = f"✅ Spremljeno kao serija **{info['batch']}** ({info['rows']} anketa"
            msg += f", {info['scans_saved']} skenova)." if info["scans_saved"] else ")."
            if info.get("updated"):
                msg += " Izmjene su zamijenile raniji Excel ove serije."
            st.success(msg)
            links = []
            if info.get("results_link"):
                links.append(f"[Excel rezultati]({info['results_link']})")
            if info.get("scans_link"):
                links.append(f"[Mapa sa skenovima]({info['scans_link']})")
            if links:
                st.markdown(" · ".join(links))


# ==============================================================================
# 5. SVI SPREMLJENI REZULTATI IZ GOOGLE DRIVEA (osnova za izvještaje)
# ==============================================================================

if drive_service is not None:
    st.markdown("---")
    st.subheader("📚 Svi spremljeni rezultati iz Google Drivea")
    if st.button(f"Učitaj sve spremljene ankete: {form_type}"):
        try:
            with st.spinner("Učitavam iz Google Drivea..."):
                st.session_state.all_results = (form_type, drive.load_all_results(drive_service, form_type))
        except Exception as e:
            st.error(f"Učitavanje nije uspjelo: {e}")

    loaded = st.session_state.get("all_results")
    if loaded and loaded[0] == form_type:
        df_all = loaded[1]
        if df_all.empty:
            st.info("Za ovaj obrazac još nema spremljenih anketa.")
        else:
            st.write(f"Ukupno **{len(df_all)}** anketa iz **{df_all['Serija'].nunique()}** serija.")
            if "Datoteka" in df_all.columns:
                dupes = df_all[df_all["Datoteka"].duplicated(keep=False)]
                if not dupes.empty:
                    st.warning(
                        f"Pazite: {dupes['Datoteka'].nunique()} skenova pojavljuje se u više serija "
                        "(možda su spremljeni dvaput). Provjerite stupce *Serija* i *Datoteka*."
                    )
            st.dataframe(df_all, use_container_width=True)
            st.download_button(
                "💾 Preuzmi sve rezultate (.xlsx)",
                data=drive.dataframe_to_xlsx_bytes(df_all),
                file_name=f"{form_type}_svi_rezultati.xlsx",
                mime=drive.XLSX_MIME,
            )

            # --- Svi PDF izvještaji odjednom (jedan izvještaj = jedan nastavnik na jednom kolegiju) ---
            if reports.has_report(form_type):
                st.markdown("#### 📄 PDF izvještaji")
                groups, incomplete = reports.group_surveys(form_type, df_all)
                if incomplete:
                    needed = [k for k in reports.REPORTS[form_type]["group_by"] if k != "Identifikacijski ključ"]
                    st.warning(f"{incomplete} anketa nema upisane sve podatke ({', '.join(needed).lower()}) "
                               "pa ne ulaze u izvještaje.")
                min_n = reports.REPORTS[form_type]["min_surveys"]
                too_small = [g for g in groups if len(g[1]) < min_n]
                groups = [g for g in groups if len(g[1]) >= min_n]
                if too_small:
                    st.info(f"Kriterij za ispis izvješća je najmanje {min_n} anketa. Preskače se: " + "; ".join(
                        f"{reports.report_file_name(form_type, gi)[:-4]} ({len(gd)})" for gi, gd in too_small))
                faculties = sorted({g_info["Fakultet"] for g_info, _ in groups})
                if not groups:
                    st.info("Nema anketa s upisanim podacima za izvještaje.")
                else:
                    chosen = st.selectbox("Fakultet", ["Svi fakulteti"] + faculties, key="bulk_faculty")
                    selected = [g for g in groups if chosen == "Svi fakulteti" or g[0]["Fakultet"] == chosen]
                    st.caption(f"Bit će izrađeno **{len(selected)}** izvještaja "
                               f"({sum(len(d) for _, d in selected)} anketa).")
                    to_drive = st.checkbox("Spremi izvještaje i u Google Drive (mapa *Izvještaji*, po fakultetima)",
                                           value=True, key="bulk_to_drive")
                    if st.button("📄 Izradi izvještaje", type="primary", disabled=not report_year):
                        import zipfile
                        zip_buf = io.BytesIO()
                        problems, used_names, drive_folders = [], set(), {}
                        progress = st.progress(0.0, text="Izrađujem izvještaje...")
                        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                            for k, (g_info, g_df) in enumerate(selected, start=1):
                                name = reports.report_file_name(form_type, g_info)
                                folder = reports.safe_file_part(g_info["Fakultet"])
                                base, n_dup = name[:-4], 2
                                while f"{folder}/{name}" in used_names:  # Isti naziv u istom fakultetu
                                    extra = g_info.get("Identifikacijski ključ") or str(n_dup)
                                    name, n_dup = f"{base} ({extra}).pdf", n_dup + 1
                                used_names.add(f"{folder}/{name}")
                                try:
                                    pdf, w = reports.build_report(form_type, g_df, g_info, report_semester,
                                                                  report_year, report_date)
                                    zf.writestr(f"{folder}/{name}", pdf)
                                    problems += [f"{name}: {x}" for x in w]
                                    if to_drive:
                                        _, folder_link = drive.save_report(drive_service, form_type, folder, name, pdf)
                                        drive_folders[folder] = folder_link
                                except Exception as e:
                                    problems.append(f"{name}: izvještaj nije izrađen ({e})")
                                progress.progress(k / len(selected), text=f"Izrađujem izvještaje... {k}/{len(selected)}")
                        progress.empty()
                        st.session_state.bulk_reports = {
                            "form": form_type, "zip": zip_buf.getvalue(), "count": len(selected),
                            "problems": problems, "folders": drive_folders,
                            "name": f"Izvjestaji_{reports.safe_file_part(chosen)}.zip",
                        }

                    done = st.session_state.get("bulk_reports")
                    if done and done["form"] == form_type:
                        st.success(f"✅ Izrađeno {done['count']} izvještaja.")
                        if done["folders"]:
                            st.markdown("Spremljeno u Google Drive: " + " · ".join(
                                f"[{f}]({link})" for f, link in sorted(done["folders"].items()) if link))
                        if done["problems"]:
                            with st.expander(f"⚠️ Napomene ({len(done['problems'])})"):
                                for msg in done["problems"]:
                                    st.write("- " + msg)
                        st.download_button("⬇️ Preuzmi sve izvještaje (.zip)", data=done["zip"],
                                           file_name=done["name"], mime="application/zip")
