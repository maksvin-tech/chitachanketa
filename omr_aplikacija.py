import os
import cv2
import numpy as np
import pandas as pd
import streamlit as st
import fitz  # PyMuPDF
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LTChar

# ==============================================================================
# 1. DEFINICIJA STUPACA I PRAVILA KODIRANJA (IDENTIČNO VAŠIM EXCEL TABLICAMA)
# ==============================================================================

CONFIG = {
    "Procjena nastavnika (V4)": {
        "pdf": "Anketa za procjenu nastavnika V4.pdf",
        "scan_mode": "A4_SINGLE",
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
        "scan_mode": "A4_TWO_PAGES",  # ili A3_SINGLE
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

def evaluate_bubbles(aligned_pages, pages_rows, questions_spec, scale, fill_threshold=0.18):
    """Mjeri zacrnjenost svakog kružića i vraća listu odgovora + vizualizaciju."""
    debug_pages = [p.copy() for p in aligned_pages]
    thresh_pages = []

    for p in aligned_pages:
        gray = cv2.cvtColor(p, cv2.COLOR_BGR2GRAY)
        # Adaptivni prag hvata i plavu i crnu kemijsku olovku
        th = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                   cv2.THRESH_BINARY_INV, 31, 10)
        thresh_pages.append(th)

    radius = int(4.2 * scale)
    answers = []

    for page_idx, row_spec, values in questions_spec:
        if isinstance(row_spec, list):
            coords = [pages_rows[page_idx][r][0] for r in row_spec]
        else:
            coords = pages_rows[page_idx][row_spec]

        ratios = []
        for (cx, cy) in coords:
            ix, iy = int(round(cx)), int(round(cy))
            mask = np.zeros(thresh_pages[page_idx].shape, dtype=np.uint8)
            cv2.circle(mask, (ix, iy), radius, 255, -1)
            dark_pixels = cv2.countNonZero(cv2.bitwise_and(thresh_pages[page_idx], thresh_pages[page_idx], mask=mask))
            total_pixels = cv2.countNonZero(mask)
            ratios.append(dark_pixels / float(total_pixels))

        # Određivanje koji je kružić zacrnjen
        max_r = max(ratios)
        min_r = min(ratios)
        # Kružić mora biti osjetno tamniji od praznih kružića u istom retku
        marked_indices = [
            i for i, r in enumerate(ratios)
            if r > (min_r + 0.12) and r >= max(fill_threshold, max_r * 0.78)
        ]

        if len(marked_indices) == 1:
            idx = marked_indices[0]
            ans = values[idx]
            color = (0, 200, 0)  # Zeleno = uredno očitan odgovor
        else:
            ans = 'x'
            color = (0, 0, 230)  # Crveno = prazno ili dvostruko zacrnjeno ('x')

        answers.append(ans)

        # Crtanje kružića za vizualnu kontrolu
        for i, (cx, cy) in enumerate(coords):
            ix, iy = int(round(cx)), int(round(cy))
            if len(marked_indices) == 1 and i == marked_indices[0]:
                cv2.circle(debug_pages[page_idx], (ix, iy), radius + 3, (0, 200, 0), 2)
            elif i in marked_indices:
                cv2.circle(debug_pages[page_idx], (ix, iy), radius + 3, (0, 140, 255), 2)
            else:
                cv2.circle(debug_pages[page_idx], (ix, iy), radius, (200, 200, 200), 1)

    return answers, debug_pages

# ==============================================================================
# 4. KORISNIČKO SUČELJE (STREAMLIT WEB/DESKTOP APLIKACIJA)
# ==============================================================================

st.set_page_config(page_title="UNIZG OMR Čitač Anketa", layout="wide")
st.title("📋 OMR Čitač Anketa – Sveučilište u Zagrebu")

form_type = st.sidebar.selectbox("1. Odaberite vrstu obrasca:", list(CONFIG.keys()))
sensitivity = st.sidebar.slider("Osjetljivost na kemijsku olovku:", 0.10, 0.35, 0.18, 0.01)

cfg = CONFIG[form_type]
if not os.path.exists(cfg["pdf"]):
    st.error(f"U mapi programa nedostaje izvorna PDF datoteka: **{cfg['pdf']}**")
    st.stop()

ref_images, pages_rows, scale = load_pdf_template(cfg["pdf"])

uploaded_files = st.file_uploader(
    "2. Učitajte skenirane JPEG slike (za A3 obrasce učitajte parove slika redom: naslovna pa unutarnja):",
    type=["jpg", "jpeg", "png"], accept_multiple_files=True
)

if uploaded_files:
    files_sorted = sorted(uploaded_files, key=lambda f: f.name)
    all_rows = []
    previews = []

    if cfg["scan_mode"] == "A4_SINGLE":
        for f in files_sorted:
            img = cv2.imdecode(np.frombuffer(f.read(), np.uint8), cv2.IMREAD_COLOR)
            aligned = align_to_reference(img, ref_images[0])
            ans, dbg = evaluate_bubbles([aligned], pages_rows, cfg["questions"], scale, sensitivity)
            all_rows.append([f.name] + ans)
            previews.append((f.name, dbg))

    elif cfg["scan_mode"] == "A3_TWO_IMAGES":
        # Spajanje po 2 slike: 1. slika (desna polovica = Str. 1), 2. slika (lijevo = Str. 2, desno = Str. 3)
        for i in range(0, len(files_sorted) - 1, 2):
            f_front, f_inner = files_sorted[i], files_sorted[i + 1]
            img1 = cv2.imdecode(np.frombuffer(f_front.read(), np.uint8), cv2.IMREAD_COLOR)
            img2 = cv2.imdecode(np.frombuffer(f_inner.read(), np.uint8), cv2.IMREAD_COLOR)

            h1, w1, _ = img1.shape
            h2, w2, _ = img2.shape
            page1_raw = img1[:, w1 // 2:]       # Desna polovica 1. slike = Naslovnica
            page2_raw = img2[:, :w2 // 2]       # Lijeva polovica 2. slike = Stranica 2
            page3_raw = img2[:, w2 // 2:]       # Desna polovica 2. slike = Stranica 3

            p1 = align_to_reference(page1_raw, ref_images[0])
            p2 = align_to_reference(page2_raw, ref_images[1])
            p3 = align_to_reference(page3_raw, ref_images[2])

            ans, dbg = evaluate_bubbles([p1, p2, p3], pages_rows, cfg["questions"], scale, sensitivity)
            label = f"{f_front.name} + {f_inner.name}"
            all_rows.append([label] + ans)
            previews.append((label, dbg))

    # Prikaz tablice i preuzimanje u Excelu
    df_out = pd.DataFrame(all_rows, columns=["Datoteka"] + cfg["columns"])
    st.subheader("📊 Očitani podaci (možete ih i ručno dopuniti prije izvoza)")
    edited_df = st.data_editor(df_out, num_rows="dynamic", use_container_width=True)

    # Gumb za spremanje u Excel (.xlsx)
    excel_file = "rezultati_anketa.xlsx"
    edited_df.drop(columns=["Datoteka"]).to_excel(excel_file, index=False)
    with open(excel_file, "rb") as f_xlsx:
        st.download_button(
            "💾 Preuzmi gotovu Excel tablicu (.xlsx)",
            data=f_xlsx,
            file_name=f"{form_type}_rezultati.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Vizualna provjera zelenih kružića
    st.subheader("🔍 Vizualna provjera prepoznatih kružića")
    selected = st.selectbox("Odaberite anketu za pregled:", [p[0] for p in previews])
    for name, dbg_imgs in previews:
        if name == selected:
            cols = st.columns(len(dbg_imgs))
            for idx, d_img in enumerate(dbg_imgs):
                cols[idx].image(cv2.cvtColor(d_img, cv2.COLOR_BGR2RGB), caption=f"Stranica {idx+1}", use_container_width=True)