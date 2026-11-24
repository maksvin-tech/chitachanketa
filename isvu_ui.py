"""
Zaslon "ISVU online i izvješća V4" u OMR aplikaciji:
  1. učitavanje ISVU izvoza (XML ili ZIP) -> provjera -> spremanje u OneDrive
  2. izvješća V4 za sastavnice i Sveučilište: ISVU, papir-olovka ili oboje; ZS, LJS ili cijela godina
"""

import io
import zipfile
from datetime import date

import pandas as pd
import streamlit as st

import isvu
import reports
import reports_v4

V4_FORM = "Procjena nastavnika (V4)"
PERIODS = {"Zimski semestar": "Zimski", "Ljetni semestar": "Ljetni", "Cijela akademska godina": None}


def _parse_uploads(uploaded):
    files = isvu.read_uploads([(f.name, f.getvalue()) for f in uploaded])
    parsed, errors, raw = [], [], {}
    progress = st.progress(0.0, text="Čitam ISVU datoteke…")
    for i, (name, data) in enumerate(files):
        try:
            meta, df = isvu.parse(data, name)
            parsed.append((meta, df))
            raw.setdefault(isvu.unit_key(meta), []).append((name, data))
        except Exception as e:
            errors.append(f"{name}: {e}")
        progress.progress((i + 1) / max(len(files), 1), text=f"Čitam ISVU datoteke… {i + 1}/{len(files)}")
    progress.empty()
    units = isvu.combine(parsed)
    return {key: (meta, df, raw.get(key, [])) for key, (meta, df) in units.items()}, errors, len(files)


def _check_table(units):
    years = pd.Series([k[1] for k, (m, d, r) in units.items() if m["Anketiranja"]])
    main_year = years.mode().iloc[0] if len(years) else ""
    rows = []
    for (fac, year, sem), (m, df, raw) in sorted(units.items()):
        note = []
        if not m["Anketiranja"]:
            note.append("prazan izvoz - ponovite izvoz u ISVU-u")
        if main_year and year != main_year:
            note.append(f"druga ak. godina ({year})")
        if len(raw) > 1:
            note.append(f"spojeno {len(raw)} datoteke")
        rows.append({"Sastavnica": fac, "Ak. godina": year, "Semestar": sem, "Anketiranja": m["Anketiranja"],
                     "Ankete": m["Anketni listići"], "Nastavnika": m["Nastavnika"], "Predmeta": m["Kolegija"],
                     "Provedba": f"{m['Datum od']} - {m['Datum do']}", "Napomena": "; ".join(note),
                     "Datoteke": m["Datoteka"]})
    return pd.DataFrame(rows)


def _upload_section(drive, drive_service):
    st.markdown("### 🌐 1. Učitaj ISVU izvoz")
    st.caption("U ISVU-u (Anketa za VU) odaberite **Ispis sirovih podataka za VU (XML)**. Ovdje povucite sve "
               "datoteke jednog semestra odjednom ili jedan ZIP s njima. Sastavnica, semestar i godina prepoznaju se "
               "iz sadržaja, pa imena datoteka nisu važna; dijelovi istog izvoza (\"dio 2\") spajaju se sami.")
    uploaded = st.file_uploader("ISVU XML datoteke ili ZIP", type=["xml", "zip"], accept_multiple_files=True,
                                key="isvu_upload")
    if uploaded and st.button("🔍 Pročitaj datoteke", type="primary", key="isvu_parse"):
        units, errors, n_files = _parse_uploads(uploaded)
        st.session_state.isvu_units = units
        st.session_state.isvu_errors = errors
        st.session_state.isvu_n_files = n_files
        st.session_state.pop("isvu_saved", None)

    units = st.session_state.get("isvu_units")
    if not units:
        return
    for e in st.session_state.get("isvu_errors", []):
        st.error(e)
    table = _check_table(units)
    non_empty = table[table["Anketiranja"] > 0]
    st.caption(f"Pročitano {st.session_state.get('isvu_n_files', 0)} datoteka: {len(table)} sastavnica × semestara. "
               f"Ukupno **{int(non_empty['Anketiranja'].sum()):,} anketiranja** i "
               f"**{int(non_empty['Ankete'].sum()):,} anketa**.".replace(",", "."))
    for (year, sem), part in non_empty.groupby(["Ak. godina", "Semestar"]):
        st.caption(f"{year} {sem.lower()} semestar: {len(part)} sastavnica, "
                   f"{int(part['Anketiranja'].sum()):,} anketiranja".replace(",", "."))
    st.dataframe(table, hide_index=True, use_container_width=True)
    if (table["Napomena"].str.contains("prazan")).any():
        st.warning("Neke datoteke su prazne (bez ijednog odgovora) i neće se spremiti. Ako sastavnica jest provodila "
                   "anketu, ponovite izvoz u ISVU-u.")

    if drive_service is None:
        st.info("Za spremanje se prijavite u OneDrive (bočna traka). Izvješća možete izraditi i bez spremanja.")
        return
    if st.session_state.get("isvu_saved"):
        st.success(st.session_state.isvu_saved)
        return
    st.caption("Spremanje: OneDrive › OMR Ankete › Rezultati › ISVU online › godina › semestar. Za svaku sastavnicu "
               "tablica (jedan redak po anketi) i izvorne XML datoteke. Ako je ista sastavnica i semestar već spremljen, "
               "zamijeni se (vrijedi novi izvoz).")
    if st.button("💾 Spremi u OneDrive", key="isvu_save"):
        todo = [(k, v) for k, v in sorted(units.items()) if v[0]["Anketiranja"]]
        progress = st.progress(0.0, text="Spremam…")
        failed = []
        for i, (key, (meta, df, raw)) in enumerate(todo):
            try:
                drive.save_isvu_unit(drive_service, meta, df, raw)
            except Exception as e:
                failed.append(f"{key[0]} ({key[2]}): {e}")
            progress.progress((i + 1) / len(todo), text=f"Spremam… {i + 1}/{len(todo)} ({key[0]})")
        progress.empty()
        for f in failed:
            st.error(f"Nije spremljeno: {f}")
        if not failed:
            st.session_state.isvu_saved = f"✅ Spremljeno {len(todo)} sastavnica × semestara u OneDrive."
            st.session_state.pop("v4_rows", None)  # Izvješća neka učitaju nove podatke
            st.rerun()


def _load_rows(drive, drive_service):
    """ISVU (OneDrive ili upravo učitano) + papirnate ankete V4 iz OneDrivea."""
    frames, notes = [], []
    if drive_service is not None:
        with st.spinner("Učitavam ISVU tablice iz OneDrivea…"):
            items = drive.list_isvu(drive_service)
            if items:
                frames.append(drive.load_isvu(drive_service, items))
                notes.append(f"ISVU: {len(items)} tablica iz OneDrivea")
        with st.spinner("Učitavam papirnate ankete V4 iz OneDrivea…"):
            paper = drive.load_all_results(drive_service, V4_FORM)
            if not paper.empty:
                rows = reports_v4.paper_rows(paper)
                frames.append(rows)
                notes.append(f"papir: {len(rows)} anketa")
    units = st.session_state.get("isvu_units") or {}
    loaded_isvu = any(not f.empty and (f["Izvor"] == isvu.SOURCE).any() for f in frames)
    if units and not loaded_isvu:
        frames.append(pd.concat([df for _m, df, _r in units.values()], ignore_index=True))
        notes.append("ISVU: upravo učitane datoteke (nisu iz OneDrivea)")
    rows = pd.concat(frames, ignore_index=True).fillna("") if frames else pd.DataFrame(columns=isvu.COLUMNS)
    periods = {}
    for (fac, year, sem), (m, _df, _r) in units.items():
        if m["Datum od"]:
            periods[(fac, year, sem)] = f"{m['Datum od']} - {m['Datum do']}"
    return rows.astype(str), notes, periods


def _reports_section(drive, drive_service):
    st.markdown("### 📊 2. Izvješća V4 (studentsko vrednovanje nastave)")
    st.caption("Za svaku sastavnicu: izvješće (tablice 1-11), ključ za identifikaciju nastavnika i rang nastavnika "
               f"(najmanje 1 i najmanje {reports_v4.MIN_RANK} anketa). Za Sveučilište: zbirno izvješće s rezultatima po "
               "sastavnicama (bez ključa i ranga).")
    if st.button("📥 Učitaj podatke (ISVU i papir)", key="v4_load"):
        st.session_state.v4_rows, st.session_state.v4_notes, st.session_state.v4_periods = _load_rows(drive,
                                                                                                    drive_service)
    rows = st.session_state.get("v4_rows")
    if rows is None:
        if st.session_state.get("isvu_units"):
            st.caption("Kliknite gumb da se uključe i upravo učitane ISVU datoteke.")
        return
    if rows.empty:
        st.warning("Nema podataka: spremite ISVU izvoz (korak 1) ili papirnate ankete V4 (skenirane ankete).")
        return
    st.caption(" · ".join(st.session_state.get("v4_notes", [])))

    counts = rows.loc[rows["Akademska godina"] != "", "Akademska godina"].value_counts()
    years = sorted(counts.index, reverse=True)
    c1, c2, c3 = st.columns(3)
    year = c1.selectbox("Akademska godina", years, index=years.index(counts.idxmax()), key="v4_year",
                        format_func=lambda y: f"{y} ({counts[y]:,} anketa)".replace(",", "."))
    periods = c2.multiselect("Razdoblje", list(PERIODS), default=list(PERIODS), key="v4_periods")
    sources_here = set(rows.loc[rows["Akademska godina"] == year, "Izvor"])
    source_opts = [k for k in reports_v4.SOURCES
                   if k == "oboje" and len(sources_here) > 1 or reports_v4.SOURCE_OF_ROW.get(k) in sources_here]
    source = c3.selectbox("Izvor podataka", source_opts, format_func=lambda k: reports_v4.SOURCES[k], key="v4_source")

    base_rows = reports_v4.select(rows, year, None, None, source)
    faculties = sorted({f for f in base_rows["Fakultet"] if f})
    chosen = st.multiselect(f"Sastavnice ({len(faculties)})", faculties, default=faculties, key=f"v4_fac|{year}|{source}")
    uni = st.checkbox("I zbirno izvješće za Sveučilište", value=len(faculties) > 1, key="v4_uni")
    report_date = st.date_input("Datum izvješća", value=date.today(), format="DD.MM.YYYY", key="v4_date")
    save = st.checkbox("Spremi u OneDrive (Izvještaji › Procjena nastavnika (V4) › sastavnica › godina)",
                       value=drive_service is not None, disabled=drive_service is None, key="v4_save")
    n_sets = len(chosen) * len(periods)
    st.caption(f"Izradit će se {n_sets * 4 + (len(periods) if uni else 0)} PDF datoteka.")

    if st.button("📄 Izradi izvješća", type="primary", disabled=not periods or not (chosen or uni), key="v4_build"):
        out, problems, saved = {}, [], 0
        progress = st.progress(0.0, text="Izrađujem izvješća…")
        total = n_sets + (len(periods) if uni else 0)
        done = 0
        year_dir = drive.year_folder(year)
        for period in periods:
            sem = PERIODS[period]
            for fac in chosen:
                sel = reports_v4.select(rows, year, sem, fac, source)
                done += 1
                progress.progress(done / max(total, 1), text=f"{fac} - {period.lower()} ({done}/{total})")
                if sel.empty or not (sel["Oznaka nastavnika"] != "").any():
                    continue
                try:
                    info = st.session_state.get("v4_periods", {}).get((fac, year, sem)) if sem else None
                    files, _s = reports_v4.unit_files(sel, fac, year, sem, source, report_date, info)
                except Exception as e:
                    problems.append(f"{fac} ({period}): {e}")
                    continue
                for name, pdf in files.items():
                    out[f"{reports.safe_file_part(fac)}/{name}"] = pdf
                    if save:
                        try:
                            drive.save_report(drive_service, V4_FORM, reports.safe_file_part(fac), name, pdf,
                                              subfolder=year_dir)
                            saved += 1
                        except Exception as e:
                            problems.append(f"OneDrive, {name}: {e}")
            if uni:
                done += 1
                progress.progress(done / max(total, 1), text=f"Sveučilište - {period.lower()}")
                sel = reports_v4.select(rows, year, sem, None, source)
                if not sel.empty:
                    try:
                        name, pdf, _s = reports_v4.university_report(sel, year, sem, source, report_date)
                        out[f"0 {reports_v4.UNIVERSITY}/{name}"] = pdf
                        if save:
                            drive.save_report(drive_service, V4_FORM, f"0 {reports_v4.UNIVERSITY}/{year_dir}",
                                              name, pdf)
                            saved += 1
                    except Exception as e:
                        problems.append(f"Sveučilište ({period}): {e}")
        progress.empty()
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for name, pdf in out.items():
                z.writestr(name, pdf)
        st.session_state.v4_zip = (f"Izvješća V4 {reports_v4.SHORT_SOURCE[source]} {year_dir}.zip",
                                   buf.getvalue(), len(out), saved, problems)

    if st.session_state.get("v4_zip"):
        name, data, n, saved, problems = st.session_state.v4_zip
        for p in problems:
            st.error(p)
        st.success(f"Izrađeno {n} PDF datoteka" + (f", u OneDrive spremljeno {saved}." if saved else "."))
        st.download_button("⬇️ Preuzmi sve (ZIP)", data, file_name=name, mime="application/zip", key="v4_dl")
        st.caption("Ključ i rang sadrže imena nastavnika - šaljite ih samo ovlaštenim osobama (npr. prodekanu za "
                   "nastavu i povjerenstvu za kvalitetu).")


def render(drive, drive_service):
    st.markdown("## 🌐 ISVU online i izvješća V4")
    _upload_section(drive, drive_service)
    st.markdown("---")
    _reports_section(drive, drive_service)
