import os
import cv2
import numpy as np
import pandas as pd
import streamlit as st
import fitz  # PyMuPDF
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LTChar

import drive_storage as drive

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

def bubble_radius(scale):
    return int(4.2 * scale)


def question_coords(pages_rows, page_idx, row_spec):
    """Koordinate kružića jednog pitanja (vodoravni redak ili okomito složeni kružići)."""
    if isinstance(row_spec, list):
        return [pages_rows[page_idx][r][0] for r in row_spec]
    return pages_rows[page_idx][row_spec]


def measure_ratios(aligned_pages, pages_rows, questions_spec, scale):
    """Mjeri zacrnjenost svakog kružića. Vraća listu (jedna lista omjera po pitanju).
    Ovo je jedini skupi korak - rezultat je malen pa se može čuvati za stotine anketa."""
    thresh_pages = []
    for p in aligned_pages:
        gray = cv2.cvtColor(p, cv2.COLOR_BGR2GRAY)
        # Adaptivni prag hvata i plavu i crnu kemijsku olovku
        th = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                   cv2.THRESH_BINARY_INV, 31, 10)
        thresh_pages.append(th)

    radius = bubble_radius(scale)
    size = 2 * radius + 1
    disk = np.zeros((size, size), dtype=np.uint8)
    cv2.circle(disk, (radius, radius), radius, 255, -1)  # Ista maska kružića, samo jednom

    all_ratios = []
    for page_idx, row_spec, _values in questions_spec:
        th = thresh_pages[page_idx]
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
            dark_pixels = cv2.countNonZero(cv2.bitwise_and(roi, roi, mask=m))
            ratios.append(dark_pixels / float(total_pixels) if total_pixels else 0.0)
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
    if scan_mode == "A3_TWO_IMAGES":
        return [(f"{a.name} + {b.name}", [a, b]) for a, b in zip(files_sorted[0::2], files_sorted[1::2])]
    return None  # Način skeniranja još nije podržan


def decode_image(f):
    return cv2.imdecode(np.frombuffer(f.getvalue(), np.uint8), cv2.IMREAD_COLOR)


def align_unit(scan_mode, unit_files, ref_images):
    """Učitava slike jedne ankete i poravnava svaku stranicu prema PDF predlošku."""
    if scan_mode == "A4_SINGLE":
        return [align_to_reference(decode_image(unit_files[0]), ref_images[0])]
    # A3_TWO_IMAGES: 1. slika (desna polovica = Str. 1), 2. slika (lijevo = Str. 2, desno = Str. 3)
    img1, img2 = decode_image(unit_files[0]), decode_image(unit_files[1])
    w1, w2 = img1.shape[1], img2.shape[1]
    return [
        align_to_reference(img1[:, w1 // 2:], ref_images[0]),
        align_to_reference(img2[:, :w2 // 2], ref_images[1]),
        align_to_reference(img2[:, w2 // 2:], ref_images[2]),
    ]


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
    st.sidebar.info("Google Drive još nije podešen. Upute su u datoteci **PODESAVANJE_GOOGLE_DRIVE.md**.")
else:
    try:
        drive_service = drive.get_service(interactive=False)  # Tiha prijava ako je već spremljena
    except Exception as e:
        st.sidebar.error(f"Greška pri spajanju na Google Drive: {e}")

    if drive_service is None:
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
    units = scan_units(cfg["scan_mode"], files_sorted)
    if units is None:
        st.error(f"Način skeniranja **{cfg['scan_mode']}** za ovaj obrazac još nije podržan.")
        units = []

    # Skupi korak (poravnanje + mjerenje kružića) radi se samo JEDNOM po učitanim datotekama.
    # Ispravci u tablici, pregled i promjena osjetljivosti koriste spremljena mjerenja,
    # a slike se ne čuvaju u memoriji.
    proc_key = (form_type, tuple(getattr(f, "file_id", None) or f"{f.name}:{f.size}" for f in files_sorted))
    cache = st.session_state.get("omr_cache")
    if cache is None or cache["key"] != proc_key:
        measurements = []
        if units:
            progress = st.progress(0.0, text="Očitavam ankete...")
            for i, (_label, unit_files) in enumerate(units, start=1):
                pages = align_unit(cfg["scan_mode"], unit_files, ref_images)
                measurements.append(measure_ratios(pages, pages_rows, cfg["questions"], scale))
                progress.progress(i / len(units), text=f"Očitavam ankete... {i}/{len(units)}")
            progress.empty()
        cache = {"key": proc_key, "measurements": measurements}
        st.session_state.omr_cache = cache
        st.session_state.pop("omr_preview", None)
    measurements = cache["measurements"]

    all_rows = [
        [label] + answers_from_ratios(m, cfg["questions"], sensitivity)
        for (label, _files), m in zip(units, measurements)
    ]

    # Prikaz tablice i preuzimanje u Excelu
    df_out = pd.DataFrame(all_rows, columns=["Datoteka"] + cfg["columns"])
    st.subheader("📊 Očitani podaci (možete ih i ručno dopuniti prije izvoza)")
    edited_df = st.data_editor(df_out, num_rows="dynamic", use_container_width=True)

    # Gumb za spremanje u Excel (.xlsx) - izrađuje se u memoriji, bez datoteke na disku
    st.download_button(
        "💾 Preuzmi gotovu Excel tablicu (.xlsx)",
        data=drive.dataframe_to_xlsx_bytes(edited_df.drop(columns=["Datoteka"])),
        file_name=f"{form_type}_rezultati.xlsx",
        mime=drive.XLSX_MIME
    )

    # Spremanje serije u Google Drive
    if drive_service is not None and all_rows:
        st.subheader("☁️ Spremanje u Google Drive")
        batch_key = form_type + "|" + "|".join(f"{f.name}:{f.size}" for f in files_sorted)
        saved_batches = st.session_state.setdefault("saved_batches", {})
        previous = saved_batches.get(batch_key)

        save_scans = st.checkbox("Spremi i izvorne skenove", value=True, disabled=previous is not None)
        button_label = "☁️ Spremi izmjene u Google Drive" if previous else "☁️ Spremi u Google Drive"
        if st.button(button_label, type="primary"):
            try:
                with st.spinner("Spremam u Google Drive..."):
                    scans = None
                    if save_scans and previous is None:
                        scans = [(f.name, f.getvalue()) for f in files_sorted]
                    saved_batches[batch_key] = drive.save_batch(
                        drive_service, form_type, edited_df, scans, previous=previous
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

    # Vizualna provjera zelenih kružića - crta se samo odabrana anketa
    if units:
        st.subheader("🔍 Vizualna provjera prepoznatih kružića")
        sel = st.selectbox("Odaberite anketu za pregled:", range(len(units)), format_func=lambda i: units[i][0])
        preview_key = (proc_key, sel, sensitivity)
        preview = st.session_state.get("omr_preview")
        if preview is None or preview["key"] != preview_key:
            with st.spinner("Pripremam pregled..."):
                pages = align_unit(cfg["scan_mode"], units[sel][1], ref_images)
                dbg = draw_debug(pages, pages_rows, cfg["questions"], scale, measurements[sel], sensitivity)
                preview = {"key": preview_key, "images": [shrink_for_preview(d) for d in dbg]}
            st.session_state.omr_preview = preview
        cols = st.columns(len(preview["images"]))
        for idx, img in enumerate(preview["images"]):
            cols[idx].image(img, caption=f"Stranica {idx+1}", use_container_width=True)

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
