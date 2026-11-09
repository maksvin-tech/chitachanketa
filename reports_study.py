"""
PDF izvještaji za ankete studenata koji su završili studij:
preddiplomski, diplomski i integrirani.

Pravila izračuna su ista kao za anketu o nastavnicima (vidi reports.py):
postoci od UKUPNOG broja (uključujući "Nedostaje"), D = max - min,
"Ukupno za skup pitanja" iz svih odgovora zajedno, SD uzoračka.

Svi fiksni tekstovi izvještaja nalaze se u odjeljku "TEKSTOVI I PITANJA"
na početku datoteke, pa ih se može lako ispraviti.
"""

import io
from datetime import date

import numpy as np
import pandas as pd

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    Image, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)

import reports as base
from reports import fmt, fmt_int, describe, clean_column, MISSING_LABEL, _esc

GENDER_COL = "Spol\n(muško = 1)\n(žensko = 0)"
MIN_SURVEYS = 10  # Kriterij za ispis izvješća

# ==============================================================================
# TEKSTOVI I PITANJA
# ==============================================================================

SATISFACTION = [(1, "potpuno nezadovoljan (1)"), (2, "djelomično nezadovoljan (2)"), (3, "prosječno zadovoljan (3)"),
                (4, "uglavnom zadovoljan (4)"), (5, "potpuno zadovoljan (5)")]
SATISFACTION_TEXT = ["potpuno nezadovoljnih (1)", "djelomično nezadovoljnih (2)", "prosječno zadovoljnih (3)",
                     "uglavnom zadovoljnih (4)", "potpuno zadovoljnih (5)"]
EXTENT = [(1, "nimalo (1)"), (2, "2"), (3, "3"), (4, "4"), (5, "u potpunosti (5)")]
EXTENT_TEXT = ["procjena 1 (nimalo)", "procjena 2", "procjena 3", "procjena 4", "procjena 5 (u potpunosti)"]

GRADES = [(v, fmt(v, 1).replace(",0", "")) for v in (2, 2.5, 3, 3.5, 4, 4.5, 5)]
FIRST_CHOICE = [(1, "Da"), (2, "Uglavnom da"), (3, "Ne")]
YES_NO = [(1, "Da"), (2, "Ne")]
OVERALL = [(1, "Nedovoljan (1)"), (2, "Dovoljan (2)"), (3, "Dobar (3)"), (4, "Vrlo dobar (4)"), (5, "Odličan (5)")]
OVERALL_IZVRSTAN = OVERALL[:4] + [(5, "Izvrstan (5)")]
RESPONSIBILITY_PREPARED = "Osjećaju se pripremljeni za stručni rad"
RESPONSIBILITY = [(1, "Uglavnom odgovornost studija"), (2, "Dijelom moja, a dijelom studija"),
                  (3, "Uglavnom moja odgovornost")]

AUTHORS = ("Autori anketnog upitnika su prof. dr. sc. Damir Ljubotina, prof. dr. sc. Vlasta Vizek Vidović i "
           "prof. dr. sc. Vesna Vlahović Štetić. Za koordinaciju istraživanja odgovoran je Ured za upravljanje "
           "kvalitetom, a za analizu podataka odgovoran je mag. psych. Maks Vinšćak.")

METHOD_TEXT = (
    "Različiti aspekti studija procijenjeni su na ljestvici od 1 do 5. Odgovori „ne mogu procijeniti”, "
    "neispunjeni i pogrešno ispunjeni odgovori označeni su kao „Nedostaje” (oznaka NN u tablici opisne statistike) "
    "i nisu uključeni u izračun prosjeka i standardne devijacije, ali su prikazani u distribuciji odgovora. "
    "Postoci se računaju od ukupnog broja odgovora, uključujući one koji nedostaju. Za svako pitanje i za cijeli "
    "skup pitanja zajedno prikazana je opisna statistika (tablica a: N = broj procjena, NN = broj odgovora koji "
    "nedostaju, M = aritmetička sredina, SD = standardna devijacija, D = raspon, Min i Max = najmanja i najveća "
    "procjena) te distribucija odgovora po pitanjima i za cijeli skup (tablica b). Na grafikonima je prikazana "
    "distribucija svih odgovora u skupu, ukupno (grafikon a) i po spolu (grafikon b)."
)


def _a(col, code, text, cats, kind, **extra):
    return dict(col=col, code=code, text=text, cats=cats, kind=kind, **extra)


PARENTS = {"B": "Rad službi i opći uvjeti studiranja na visokom učilištu",
           "E": "Odnos prema studentima i podrška u studiranju"}


def _g(code, title, subject, scale, items):
    """Skup pitanja. Kodovi B1/B2 i E1/E2/E3 prikazuju se kao podskupine cjelina B i E."""
    parent = code[0] if len(code) == 2 and code[0] in PARENTS else None
    return dict(code=code, title=title, subject=subject, scale=scale, items=items, parent=parent)


def _items(prefix_code, cols, numbers, texts):
    return [(c, f"{prefix_code}.{n}", t) for c, n, t in zip(cols, numbers, texts)]


# --- Zajednička pitanja (preddiplomski i integrirani imaju iste skupove B-F) ---

B1_TEXTS = [
    "Rad studentske referade fakulteta / akademije",
    "Rad administrativne službe (tajništva) Vašeg odjela/odsjeka/zavoda",
    "Rad informatičke službe za studente (dobivanje vlastite e-mail adrese i sl.)",
    "Korisnost informacijskog sustava visokih učilišta (ISVU)",
    "Mogućnost pristupa sustavu ISVU za studente (broj studomata i sl.)",
    "Korisnost WEB stranica Vašeg studija",
    "Opremljenost i ponuda literature u knjižnici",
    "Organizacija rada knjižnice",
    "Rad uprave fakulteta / akademije iz studentske perspektive",
    "Rad vodstva Vašeg studija (odjela, zavoda i sl.)",
]
B2_TEXTS = [
    "Opća kvaliteta i uređenost prostora (zgrade, pristup fakultetu i sl.)",
    "Opremljenost nastavnih dvorana",
    "Opremljenost prostora u kojima se izvode vježbe (laboratoriji, seminari i sl.)",
    "Primjerenost prostora s obzirom na broj studenata",
    "Mogućnost pristupa računalima i Internetu u prostoru fakulteta",
]
PROGRAM_TEXTS_UG = [  # preddiplomski / integrirani C16-C23
    "Sadržaj i kvaliteta obveznih kolegija",
    "Sadržaj i kvaliteta izbornih kolegija",
    "Ponuda izbornih kolegija",
    "Mogućnost pohađanja kolegija koji nisu u sastavu Vašeg studija",
    "Prilagođenost zahtjeva i težine kolegija predznanjima studenata",
    "Stupanj u kojemu kolegiji na prvoj godini olakšavaju prilagodbu na studij",
    "Povezanost i slijed sadržaja na različitim kolegijima i godinama studija",
    "Stupanj u kojemu je sadržaj studijskog programa zadovoljio Vaša očekivanja",
]
TEACHING_TEXTS = [  # izvedba nastave (diplomski C8-C24, ostali D24-D40)
    "Korisnost predavanja za razumijevanje zadanih sadržaja",
    "Praktični rad na vježbama (seminarima) i mogućnost praktične provjere stečenih znanja i vještina",
    "Korisnost i kvaliteta izvora informacija za učenje (literatura, skripte, Internet i dr.)",
    "Raspored kolegija i obveza studenata tijekom tjedna",
    "Organizacija prakse izvan fakulteta te suradnja sa stručnjacima koji rade u praksi",
    "Sudjelovanje u terenskoj nastavi (uključujući ljetne škole)",
    "Uključenost studenata u znanstveno-istraživačke projekte nastavnog osoblja",
    "Mogućnosti za međunarodnu suradnju (razmjena studenata, odlasci u inozemstvo i dr.)",
    "Jasno definiranje ishoda kolegija, tj. koja znanja i vještine studenti trebaju usvojiti nakon pojedinog kolegija",
    "Jasno definirani kriteriji za vrednovanje znanja (je li jasno što student treba naučiti za prolaz "
    "odnosno za pojedinu ocjenu)",
    "Ujednačenost kriterija na različitim kolegijima pri vrednovanju znanja",
    "Kontinuirano provjeravanje znanja na kolegijima tijekom semestra",
    "Redovite i jasne povratne informacije o uspješnosti u učenju i na ispitima",
    "Raspored ispitnih rokova i drugih oblika provjere znanja (kolokviji i sl.)",
    "Broj ispitnih rokova",
    "Način provjere znanja i vještina (izvedba ispita)",
    "Opće zadovoljstvo izvedbom studijskog programa",
]
E1_TEXTS = [
    "Poticanje studenata na slobodno izražavanje vlastitog mišljenja tijekom nastave",
    "Pokazivanje razumijevanja za studentove probleme i obaveze izvan studija",
    "Pokazivanje povjerenja u studentove mogućnosti i poticanje studentovog samopouzdanja",
    "Pristupačnost i susretljivost nastavnog osoblja prema studentima",
    "Nepristranost i pravednost u odnosu prema studentima",
]
E2_TEXTS = [
    "Dostupnost osobnih mentora",
    "Konzultacije s nastavnim osobljem i pomoć izvan vremena nastave",
    "Točnost i pravovremenost obavijesti o promjenama u organizaciji kolegija/nastave",
    "Savjetovanje / poučavanje studenata o strategijama efikasnijeg učenja",
    "Pokazivanje razumijevanja za studentske probleme vezane uz nastavu i studij",
    "Savjetovanje studenata pri izboru kolegija i mogućnostima nastavka studija",
    "Savjetovanje studenata o budućoj karijeri (zapošljavanju)",
]
E3_TEXTS_UG = [
    "Organizacija i kvaliteta studentske prehrane",
    "Usluge Studentskog centra i Student servisa (zapošljavanje studenata)",
    "Kvaliteta zdravstvenih usluga",
    "Mogućnosti organiziranja i djelovanja kroz studentske udruge",
    "Cijena studija u odnosu na dobitak od studija (ukoliko ste plaćali studij)",
    "Mogućnosti sudjelovanja u aktivnostima koje razvijaju osjećaj pripadnosti Vašem fakultetu / akademiji "
    "(izleti, proslave, natjecanja i sl.)",
]
E3_TEXTS_INT = [t for t in E3_TEXTS_UG if t != "Kvaliteta zdravstvenih usluga"]


def _yes_no(col, code, text, key):
    return dict(col=col, code=code, text=text, key=key)


YES_NO_SENTENCES = {
    "again": ("odgovara da bi ponovno upisali isti studij", "bi izabralo neki drugi studij"),
    "recommend": ("odgovara da bi drugima preporučili svoj studij", "to ne bi učinili"),
    "postgrad": ("namjerava u budućnosti upisati poslijediplomski studij", "za sada nema takve namjere"),
    "3+2": ("smatra da je organizacija studija prema načelu 3+2 korisna za njihovu struku", "smatra da nije"),
}

LEVELS = {
    # ------------------------------------------------------------------ DIPLOMSKI
    "Diplomski studij": dict(
        level="diplomski", level_gen="diplomskog", level_pl_gen="diplomskih",
        title="Vrjednovanje diplomskih studija od strane studenata koji su tijekom akademske godine {year} "
              "završili studij",
        intro=[
            "Za studije na Sveučilištu u Zagrebu, organizirane prema načelu 3+2 i 4+1, završio je ciklus diplomskog "
            "studija provedenog prema načelima Bolonjskog procesa. Ured za upravljanje kvalitetom organizirao je i "
            "proveo anketu s ciljem prikupljanja informacija o iskustvima vezanima uz izvedbu diplomskih studija. "
            "U ovome izvještaju prikazani su rezultati procjena studenata prikupljenih tijekom akademske godine "
            "{year}.",
        ],
        overall_labels=OVERALL,
        aim=("U sklopu vrjednovanja diplomskog studija pripremljena je anketa za studente koji su završili "
             "jednogodišnji ili dvogodišnji diplomski studij, a čiji je cilj vrjednovanje različitih vidika "
             "diplomskog studija u cjelini."),
        sections=["A. Podaci o studentici/studentu i završenom diplomskom studiju", "B. Studijski program",
                  "C. Izvedba nastave i vrjednovanje znanja", "D. Odnos prema studentima i podrška u studiranju",
                  "F. Opća procjena ishoda", "G. Prijedlozi i komentari"],
        n_questions=45,
        population=("Ciljna populacija su studenti koji su tijekom akademske godine {year} završili neki od "
                    "diplomskih studija na Sveučilištu u Zagrebu. Sve procjene odnosile su se na njihov diplomski "
                    "studij. Za svaku anketiranu jedinicu tiskan je potreban broj anketnih upitnika u skladu s brojem "
                    "studenata koji su bili upisani u posljednju godinu svojeg diplomskog studija, a svakom je "
                    "studijskom programu dodijeljen zasebni kod. Studenti su anketu ispunili u studentskoj službi "
                    "(referadi) pri obavljanju administrativnih poslova vezanih za završetak diplomskog studija."),
        a_items=[
            _a("Ocjena Dipl", "A.5", "Prosječna ocjena tijekom diplomskog studija navedena od strane studenata",
               GRADES, "grade", what="diplomskog studija"),
            _a("Ocjena Preddipl", "A.6", "Prosječna ocjena tijekom prijediplomskog studija navedena od strane studenata",
               GRADES, "grade", what="prijediplomskog studija"),
            _a("Prvi izbor", "A.7", "Je li studij koji ste završili bio Vaš prvi izbor?", FIRST_CHOICE, "first_choice"),
            _a("Nastavak studija", "A.8", "Studij koji ste završili predstavlja",
               [(1, "Nastavak prijediplomskog studija"), (2, "Različit diplomski studij")], "choice",
               lead=("Studenti su nakon prijediplomskog studija imali mogućnost nastaviti studij na diplomskoj razini "
                     "na istom fakultetu ili upisati diplomski studij koji ne predstavlja nastavak njihovog "
                     "prijediplomskog studija (na drugom fakultetu ili na istom fakultetu, ali uz promjenu osnovnog "
                     "područja)."),
               phrases=["odgovara da diplomski studij predstavlja nastavak njihovog prijediplomskog studija",
                        "navodi da je njihov diplomski studij različit od prijediplomskog"]),
        ],
        groups=[
            _g("B", "Studijski program", "studijski program", SATISFACTION, _items("B", [f"B{i}" for i in range(1, 8)], range(1, 8), [
                "Sadržaj i kvaliteta obveznih kolegija", "Sadržaj i kvaliteta izbornih kolegija",
                "Ponuda izbornih kolegija", "Mogućnost pohađanja kolegija koji nisu u sastavu Vašeg studija",
                "Prilagođenost zahtjeva i težine kolegija predznanjima studenata",
                "Povezanost i slijed sadržaja na različitim kolegijima i godinama studija",
                "Stupanj u kojemu je sadržaj studijskog programa zadovoljio Vaša očekivanja"])),
            _g("C", "Izvedba nastave i vrjednovanje znanja", "izvedbu nastave i vrjednovanje znanja", SATISFACTION,
               _items("C", [f"C{i}" for i in range(8, 25)], range(8, 25),
                      [t.replace(" (razmjena studenata, odlasci u inozemstvo i dr.)", "") for t in TEACHING_TEXTS])),
            _g("D", "Odnos prema studentima i podrška u studiranju",
               "odnos prema studentima i podršku u studiranju", SATISFACTION,
               _items("D", [f"D{i}" for i in range(25, 31)], range(25, 31), [
                   "Pokazivanje razumijevanja za studentove probleme i obaveze izvan studija",
                   "Pokazivanje povjerenja u studentove mogućnosti i poticanje samopouzdanja",
                   "Pristupačnost i susretljivost nastavnog osoblja prema studentima",
                   "Dostupnost osobnih mentora",
                   "Konzultacije s nastavnim osobljem i pomoć izvan vremena nastave",
                   "Savjetovanje studenata o budućoj karijeri (zapošljavanju)"])),
            _g("F", "Opća procjena ishoda", "opću procjenu ishoda studija", EXTENT,
               _items("F", [f"F{i}" for i in range(31, 38)], range(31, 38), [
                   "U kojoj mjeri je studij ispunio Vaša početna očekivanja?",
                   "U kojoj mjeri Vas je studij osposobio za rad u struci?",
                   "U kojoj mjeri Vas je studij pripremio za nastavak školovanja u struci?",
                   "Koliko ste zadovoljni mogućnostima zaposlenja nakon studija?",
                   "Koliko je općenito bio težak i zahtjevan Vaš studij?",
                   "U kojoj mjeri su ECTS bodovi usklađeni sa stvarnim opterećenjem na kolegijima?",
                   "U kojoj mjeri su informacije koje ste o studiju imali prije upisa bile točne i korisne?"])),
        ],
        yes_no=[
            _yes_no("F38", "F.38", "Kada biste ponovno birali biste li ponovno upisali isti diplomski studij?", "again"),
            _yes_no("F39", "F.39", "S obzirom na Vaša iskustva, biste li drugima preporučili svoj diplomski studij?",
                    "recommend"),
            _yes_no("F40", "F.40", "Namjeravate li u budućnosti upisati poslijediplomski studij (specijalistički ili "
                    "doktorski)?", "postgrad"),
        ],
        responsibility=("F41", "F.41"),
        overall=("F42", "F.42", "Koju biste opću ocjenu dali svom diplomskom studiju?"),
    ),
    # -------------------------------------------------------------- PREDDIPLOMSKI
    "Preddiplomski studij": dict(
        level="prijediplomski", level_gen="prijediplomskog", level_pl_gen="prijediplomskih",
        title="Vrjednovanje prijediplomskih studija od strane studenata koji su tijekom akademske godine {year} "
              "završili studij",
        intro=[
            "Za dio studija na Sveučilištu u Zagrebu, organiziranih prema načelu 3+2 i 4+1, protekle akademske "
            "godine završio je ciklus prijediplomskog studija provedenog prema načelima Bolonjskog procesa. Ured za "
            "upravljanje kvalitetom organizirao je i proveo anketu s ciljem prikupljanja informacija o iskustvima "
            "vezanima uz izvedbu prijediplomskih studija. U ovome izvještaju prikazani su rezultati procjena studenata "
            "prikupljenih tijekom akademske godine {year}.",
        ],
        overall_labels=OVERALL_IZVRSTAN,
        aim=("U sklopu vrjednovanja prijediplomskog studija pripremljena je anketa za studente koji su završili "
             "trogodišnji ili četverogodišnji prijediplomski studij, a čiji je cilj vrjednovanje različitih vidika "
             "prijediplomskog studija u cjelini."),
        sections=["A. Podaci o studentici/studentu i završenom prijediplomskom studiju",
                  "B. Rad službi i opći uvjeti studiranja na visokom učilištu",
                  "C. Studijski program (procjene se odnose na studij u cjelini)",
                  "D. Izvedba nastave i vrjednovanje znanja (procjene se odnose na studij u cjelini)",
                  "E. Odnos prema studentima i podrška u studiranju",
                  "F. Opća procjena ishoda", "G. Prijedlozi i komentari"],
        n_questions=75,
        population=("Ciljna populacija su studenti koji su tijekom akademske godine {year} završili neki od "
                    "prijediplomskih studija na Sveučilištu u Zagrebu. Sve procjene odnosile su se na njihov "
                    "prijediplomski studij. Za svaku anketiranu jedinicu tiskan je potreban broj anketnih upitnika u "
                    "skladu s brojem studenata koji su bili upisani u posljednju godinu svojeg prijediplomskog studija, "
                    "a svakom je studijskom programu dodijeljen zasebni kod. Studenti su anketu ispunili u studentskoj "
                    "službi (referadi) pri obavljanju administrativnih poslova vezanih za završetak prijediplomskog "
                    "studija."),
        a_items=[
            _a("Ocjena preddipl", "A.4", "Prosječna ocjena tijekom prijediplomskog studija navedena od strane studenata",
               GRADES, "grade", what="prijediplomskog studija"),
            _a("Vrsta studiranja", "A.5", "Vrsta studiranja",
               [(1, "Uz potporu Ministarstva"), (2, "Iz osobnih potreba")], "choice",
               phrases=["studiralo je uz potporu Ministarstva", "studiralo je iz osobnih potreba (uz plaćanje)"]),
            _a("Prvi izbor", "A.6", "Je li prijediplomski studij koji ste završili bio Vaš prvi izbor?",
               FIRST_CHOICE, "first_choice"),
            _a("Nastavak studija", "A.7", "Diplomski studij koji planiram upisati nakon završetka prijediplomskog studija",
               [(1, "Nastavak prijediplomskog studija"), (2, "Različit diplomski studij")], "choice",
               phrases=["planira upisati diplomski studij koji predstavlja nastavak njihovog prijediplomskog studija",
                        "planira upisati diplomski studij različit od prijediplomskog"]),
        ],
        groups=[
            _g("B1", "Rad administrativnih i stručnih službi", "rad administrativnih i stručnih službi", SATISFACTION,
               _items("B", [f"B{i}" for i in range(1, 11)], range(1, 11), B1_TEXTS)),
            _g("B2", "Uvjeti za studiranje u ustanovi", "uvjete za studiranje u ustanovi", SATISFACTION,
               _items("B", [f"b{i}" for i in range(11, 16)], range(11, 16), B2_TEXTS)),
            _g("C", "Studijski program: procjene se odnose na studij u cjelini", "studijski program", SATISFACTION,
               _items("C", [f"C{i}" for i in range(16, 24)], range(16, 24), PROGRAM_TEXTS_UG)),
            _g("D", "Izvedba nastave i vrjednovanje znanja: procjene se odnose na studij u cjelini",
               "izvedbu nastave i vrjednovanje znanja", SATISFACTION,
               _items("D", [f"D{i}" for i in range(24, 41)], range(24, 41), TEACHING_TEXTS)),
            _g("E1", "Odnos prema studentima", "odnos prema studentima", SATISFACTION,
               _items("E", [f"E{i}" for i in range(41, 46)], range(41, 46), E1_TEXTS)),
            _g("E2", "Podrška u studiranju", "podršku u studiranju", SATISFACTION,
               _items("E", [f"e{i}" for i in range(46, 53)], range(46, 53), E2_TEXTS)),
            _g("E3", "Ostali uvjeti studiranja", "ostale uvjete studiranja", SATISFACTION,
               _items("E", [f"e{i}e" for i in range(53, 59)], range(53, 59), E3_TEXTS_UG)),
            _g("F", "Opća procjena ishoda", "opću procjenu ishoda prijediplomskog studija", EXTENT,
               _items("F", [f"F{i}" for i in range(59, 67)], range(59, 67), [
                   "U kojoj mjeri je studij ispunio Vaša početna očekivanja?",
                   "U kojoj mjeri Vas je studij osposobio za rad u struci?",
                   "U kojoj mjeri Vas je studij pripremio za nastavak školovanja u struci?",
                   "Koliko ste zadovoljni cjelokupnim iskustvom studiranja?",
                   "Koliko ste zadovoljni mogućnostima zaposlenja nakon prijediplomskog studija?",
                   "Koliko je općenito bio težak i zahtjevan Vaš prijediplomski studij?",
                   "U kojoj mjeri su ECTS bodovi usklađeni sa stvarnim nastavnim opterećenjem na kolegijima?",
                   "U kojoj mjeri su informacije koje ste o studiju imali prije upisa bile točne i korisne?"])),
        ],
        yes_no=[
            _yes_no("f67", "F.67", "Smatrate li da je organizacija studija prema načelu 3+2 korisna za Vašu struku?",
                    "3+2"),
            _yes_no("f68", "F.68", "Kada biste ponovno birali biste li ponovno upisali isti prijediplomski studij?",
                    "again"),
            _yes_no("f69", "F.69", "S obzirom na Vaša iskustva, biste li drugima preporučili svoj prijediplomski "
                    "studij?", "recommend"),
            _yes_no("f70", "F.70", "Namjeravate li u budućnosti upisati poslijediplomski studij (specijalistički ili "
                    "doktorski)?", "postgrad"),
        ],
        responsibility=("odgovornost za studij", "F.71"),
        overall=("ukupna ocjena", "F.72", "Koju biste opću ocjenu dali svom prijediplomskom studiju?"),
    ),
    # ---------------------------------------------------------------- INTEGRIRANI
    "Integrirani studij": dict(
        level="integrirani", level_gen="integriranog", level_pl_gen="integriranih",
        title="Vrjednovanje integriranih prijediplomskih i diplomskih studija od strane studenata koji su tijekom "
              "akademske godine {year} završili studij",
        intro=[
            "Za dio studija na Sveučilištu u Zagrebu, organiziranih prema načelu 5+0, protekle akademske godine "
            "završio je ciklus integriranog prijediplomskog i diplomskog studija provedenog prema načelima Bolonjskog "
            "procesa. Ured za upravljanje kvalitetom organizirao je i proveo anketu s ciljem prikupljanja informacija "
            "o iskustvima vezanima uz izvedbu integriranih studija. U ovome izvještaju prikazani su rezultati procjena "
            "studenata prikupljenih tijekom akademske godine {year}.",
        ],
        overall_labels=OVERALL_IZVRSTAN,
        aim=("U sklopu vrjednovanja integriranih prijediplomskih i diplomskih studija pripremljena je anketa za "
             "studente koji su završili integrirani studij, a čiji je cilj vrjednovanje različitih vidika takvog "
             "studija u cjelini."),
        sections=["A. Podaci o studentici/studentu i završenom studiju",
                  "B. Rad službi i opći uvjeti studiranja na visokom učilištu",
                  "C. Studijski program (procjene se odnose na studij u cjelini)",
                  "D. Izvedba nastave i vrjednovanje znanja (procjene se odnose na studij u cjelini)",
                  "E. Odnos prema studentima i podrška u studiranju",
                  "F. Opća procjena ishoda", "G. Prijedlozi i komentari"],
        n_questions=73,
        population=("Ciljna populacija su studenti koji su tijekom akademske godine {year} završili neki od "
                    "integriranih prijediplomskih i diplomskih studija na Sveučilištu u Zagrebu. Sve procjene "
                    "odnosile su se na njihov studij. Za svaku anketiranu jedinicu tiskan je potreban broj anketnih "
                    "upitnika u skladu s brojem studenata koji su bili upisani u posljednju godinu svojeg studija, a "
                    "svakom je studijskom programu dodijeljen zasebni kod. Studenti su anketu ispunili u studentskoj "
                    "službi (referadi) pri obavljanju administrativnih poslova vezanih za završetak studija."),
        a_items=[
            _a("Ocjena studij", "A.4", "Prosječna ocjena tijekom studija navedena od strane studenata",
               GRADES, "grade", what="studija"),
            _a("Prvi izbor", "A.5", "Je li studij koji ste završili bio Vaš prvi izbor?", FIRST_CHOICE, "first_choice"),
        ],
        groups=[
            _g("B1", "Rad administrativnih i stručnih službi", "rad administrativnih i stručnih službi", SATISFACTION,
               _items("B", [f"B{i}" for i in range(1, 11)], range(1, 11), B1_TEXTS)),
            _g("B2", "Uvjeti za studiranje u ustanovi", "uvjete za studiranje u ustanovi", SATISFACTION,
               _items("B", [f"B{i}" for i in range(11, 16)], range(11, 16), B2_TEXTS)),
            _g("C", "Studijski program: procjene se odnose na studij u cjelini", "studijski program", SATISFACTION,
               _items("C", [f"C{i}" for i in range(16, 24)], range(16, 24), PROGRAM_TEXTS_UG)),
            _g("D", "Izvedba nastave i vrjednovanje znanja: procjene se odnose na studij u cjelini",
               "izvedbu nastave i vrjednovanje znanja", SATISFACTION,
               _items("D", [f"D{i}" for i in range(24, 41)], range(24, 41), TEACHING_TEXTS)),
            _g("E1", "Odnos prema studentima", "odnos prema studentima", SATISFACTION,
               _items("E", [f"E{i}" for i in range(41, 46)], range(41, 46), E1_TEXTS)),
            _g("E2", "Podrška u studiranju", "podršku u studiranju", SATISFACTION,
               _items("E", [f"E{i}" for i in range(46, 53)], range(46, 53), E2_TEXTS)),
            _g("E3", "Ostali uvjeti studiranja", "ostale uvjete studiranja", SATISFACTION,
               _items("E", [f"E{i}" for i in range(53, 58)], range(53, 58), E3_TEXTS_INT)),
            _g("F", "Opća procjena ishoda", "opću procjenu ishoda integriranog studija", EXTENT,
               _items("F", [f"F{i}" for i in range(58, 66)], range(58, 66), [
                   "U kojoj mjeri je studij ispunio Vaša početna očekivanja?",
                   "U kojoj mjeri Vas je studij osposobio za rad u struci?",
                   "U kojoj mjeri Vas je studij pripremio za nastavak školovanja u struci?",
                   "Koliko ste zadovoljni cjelokupnim iskustvom studiranja?",
                   "Koliko ste zadovoljni mogućnostima zaposlenja nakon studija?",
                   "Koliko je općenito bio težak i zahtjevan Vaš studij?",
                   "U kojoj mjeri su ECTS bodovi usklađeni sa stvarnim opterećenjem na kolegijima?",
                   "U kojoj mjeri su informacije koje ste o studiju imali prije upisa bile točne i korisne?"])),
        ],
        yes_no=[
            _yes_no("F66", "F.66", "Kada biste ponovno birali biste li ponovno upisali isti studij?", "again"),
            _yes_no("F67", "F.67", "S obzirom na Vaša iskustva, biste li drugima preporučili svoj studij?",
                    "recommend"),
            _yes_no("F68", "F.68", "Namjeravate li u budućnosti upisati poslijediplomski studij (specijalistički ili "
                    "doktorski)?", "postgrad"),
        ],
        responsibility=("odgovornost za studij", "F.69"),
        overall=("ukupna ocjena", "F.70", "Koju biste opću ocjenu dali svom studiju?"),
    ),
}

# ==============================================================================
# POMOĆNE FUNKCIJE
# ==============================================================================

M_COLOR, Z_COLOR = "#5B9BD5", "#ED7D31"


def np_text(n, total):
    """'3 (21,4%)' - broj i postotak od ukupnog."""
    return f"{n} ({fmt(n / total * 100 if total else 0, 1)}%)"


def _join(parts):
    parts = [p for p in parts if p]
    if len(parts) <= 1:
        return "".join(parts)
    return ", ".join(parts[:-1]) + " i " + parts[-1]


def _styles():
    st = base._styles()
    st.update({
        "doc_title": ParagraphStyle("doc_title", fontName="Serif-Bold", fontSize=15, leading=20, alignment=TA_CENTER),
        "h1": ParagraphStyle("h1", keepWithNext=1, fontName="Serif-Bold", fontSize=13, leading=16, spaceBefore=10, spaceAfter=8),
        "h2": ParagraphStyle("h2", keepWithNext=1, fontName="Serif-Bold", fontSize=12, leading=15, spaceBefore=8, spaceAfter=6),
        "h3": ParagraphStyle("h3", keepWithNext=1, fontName="Serif-Bold", fontSize=11, leading=14, spaceBefore=8, spaceAfter=5),
        "body": ParagraphStyle("body", fontName="Serif", fontSize=11, leading=14.5, alignment=TA_JUSTIFY,
                               spaceAfter=6),
        "list": ParagraphStyle("list", fontName="Serif-Italic", fontSize=11, leading=14.5, leftIndent=0.6 * cm,
                               spaceAfter=2),
        "cell_s": ParagraphStyle("cell_s", fontName="Serif", fontSize=9, leading=10.5, alignment=TA_LEFT),
        "cell_sb": ParagraphStyle("cell_sb", fontName="Serif-Bold", fontSize=9, leading=10.5, alignment=TA_LEFT),
    })
    return st


def _grid_table(data, col_widths, header_rows, spans=(), font_size=9, label_cols=1, bold_last=True):
    t = Table(data, colWidths=col_widths, hAlign="CENTER", repeatRows=header_rows)
    style = [
        ("FONTNAME", (0, 0), (-1, -1), "Serif"),
        ("FONTSIZE", (0, 0), (-1, -1), font_size),
        ("LEADING", (0, 0), (-1, -1), font_size + 1.5),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 0), (-1, header_rows - 1), base.GRAY),
        ("BACKGROUND", (0, header_rows), (label_cols - 1, -1), base.GRAY),
        ("BACKGROUND", (label_cols, header_rows), (-1, -1), base.BLUE),
        ("ALIGN", (label_cols, 0), (-1, -1), "RIGHT"),
        ("ALIGN", (label_cols, 0), (-1, header_rows - 1), "CENTER"),
        ("LINEBELOW", (0, 0), (-1, -1), 1, colors.white),
        ("LINEAFTER", (0, 0), (-1, -1), 1.5, colors.white),
        ("TOPPADDING", (0, 0), (-1, -1), 1.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
    ]
    if bold_last:
        style.append(("FONTNAME", (0, -1), (-1, -1), "Serif-Bold"))
    for (c0, r0), (c1, r1) in spans:
        style.append(("SPAN", (c0, r0), (c1, r1)))
    t.setStyle(TableStyle(style))
    return t


def _grouped_chart(labels, m_pct, z_pct, width_cm=15.5):
    """Grafikon po spolu (Muški i Ženski), postoci od ukupnog broja."""
    height_cm = max(5.0, 1.0 * len(labels) + 2.0)
    return base.hbar_chart(labels, [(m_pct, M_COLOR, "Muški"), (z_pct, Z_COLOR, "Ženski")],
                           width_cm=width_cm, height_cm=height_cm, suffix="%", legend=True)


def _pct_chart(labels, pcts, height_cm=None):
    """Ukupni grafikon (vrijednosti s oznakom %), prva kategorija dolje kao u Excel izvještajima."""
    return base._bar_chart(labels, pcts, width_cm=15.5, height_cm=height_cm, suffix="%")


# ==============================================================================
# IZRADA IZVJEŠTAJA
# ==============================================================================

class _Builder:
    def __init__(self, spec, df, info, academic_year, report_date):
        self.spec, self.df, self.info = spec, df.reset_index(drop=True), info
        self.year, self.date = academic_year, report_date
        self.n = len(self.df)
        self.st = _styles()
        self.story, self.warnings = [], []
        self.no = 0  # Broj tablice / grafikona
        if GENDER_COL in self.df.columns:
            self.gender, bad = clean_column(self.df[GENDER_COL], [0, 1])
            if bad:
                self.warnings.append(f"Spol: {bad} vrijednost(i) izvan ljestvice računa se kao 'Nedostaje'.")
        else:
            self.gender = pd.Series([np.nan] * self.n)
            self.warnings.append("Nedostaje stupac Spol.")

    # --- osnovno ---
    def col(self, name, allowed):
        if name not in self.df.columns:
            self.warnings.append(f"Nedostaje stupac {name} - računa se kao 'Nedostaje'.")
            return pd.Series([np.nan] * self.n)
        clean, bad = clean_column(self.df[name], allowed)
        if bad:
            self.warnings.append(f"{name}: {bad} vrijednost(i) izvan ljestvice računa se kao 'Nedostaje'.")
        return clean

    def p(self, text, style="body"):
        self.story.append(Paragraph(text, self.st[style]))

    def caption(self, text):
        return Paragraph(text, self.st["caption"])

    # --- tablica po spolu (postoci od ukupnog broja anketa) ---
    def gender_table(self, rows):
        """rows: lista (oznaka, maska)."""
        g = self.gender
        head1 = ["Odgovor", "Spol", "", "", "", "", "", "Ukupno", ""]
        head2 = ["", "M", "", "Ž", "", "Nedostaje", "", "", ""]
        head3 = ["", "N", "%", "N", "%", "N", "%", "N", "%"]
        data = [head1, head2, head3]
        sums = np.zeros(4, dtype=int)
        for label, mask in rows:
            counts = [int((mask & (g == 1)).sum()), int((mask & (g == 0)).sum()), int((mask & g.isna()).sum())]
            counts.append(sum(counts))
            sums += np.array(counts)
            row = [Paragraph(_esc(label), self.st["cell_s"])]
            for c in counts:
                row += [str(c), fmt(c / self.n * 100, 1) + "%"]
            data.append(row)
        total_row = [Paragraph("Ukupno", self.st["cell_sb"])]
        for c in sums:
            total_row += [str(c), fmt(c / self.n * 100, 1) + "%"]
        data.append(total_row)
        spans = [((0, 0), (0, 2)), ((1, 0), (6, 0)), ((1, 1), (2, 1)), ((3, 1), (4, 1)), ((5, 1), (6, 1)),
                 ((7, 0), (8, 1))]
        return _grid_table(data, [4.2 * cm] + [1.25 * cm, 1.45 * cm] * 4, header_rows=3, spans=spans)

    def gender_charts(self, rows, title):
        labels = [r[0] for r in rows]
        total = [r[1].sum() / self.n * 100 for r in rows]
        m = [(r[1] & (self.gender == 1)).sum() / self.n * 100 for r in rows]
        z = [(r[1] & (self.gender == 0)).sum() / self.n * 100 for r in rows]
        return [
            _pct_chart(labels, total), self.caption(f"Grafikon {self.no}a. {title} - distribucija odgovora"),
            Spacer(1, 0.3 * cm),
            _grouped_chart(labels, m, z), self.caption(f"Grafikon {self.no}b. {title} - distribucija odgovora po spolu"),
        ]

    def item_rows(self, clean, cats, missing_label=MISSING_LABEL):
        rows = [(missing_label, clean.isna())] if missing_label else []
        rows += [(label, clean == value) for value, label in cats]
        return rows

    # --- dijelovi izvještaja ---
    def cover(self):
        st = self.st
        if base.os.path.exists(base.LOGO):
            logo = Image(base.LOGO, width=9.5 * cm, height=9.5 * cm * 159 / 420)
            logo.hAlign = "CENTER"
            self.story.append(logo)
        self.story += [Spacer(1, 3.0 * cm), Paragraph(_esc(self.spec["title"].format(year=self.year)), st["doc_title"]),
                       Spacer(1, 2.2 * cm), Paragraph(_esc(self.info.get("Fakultet", "")), st["cover_bold"]),
                       Spacer(1, 0.5 * cm), Paragraph(_esc(self.info.get("Studijski program", "")), st["cover_bold"]),
                       Spacer(1, 4.0 * cm), Paragraph("Ured za upravljanje kvalitetom", st["cover_bold"]),
                       Paragraph("Sveučilište u Zagrebu", st["cover_bold"]), Spacer(1, 0.5 * cm),
                       Paragraph(f"Zagreb, {self.date.day}.{self.date.month}.{self.date.year}.", st["cover_bold"]),
                       PageBreak()]

    def introduction(self):
        s = self.spec
        self.p("1. Uvod", "h1")
        for para in s["intro"]:
            self.p(_esc(para.format(year=self.year)))
        self.p("1.1. Cilj i sadržaj ankete", "h2")
        self.p(_esc(s["aim"] + " Rezultate je moguće obraditi na razini Sveučilišta, anketirane jedinice te, u "
               "slučaju dovoljnog broja upitnika, na razini svake studijske grupe. Ova anketa omogućuje anketiranim "
               f"jedinicama i voditeljima studijskih grupa analizu provedbe {s['level_gen']} studija iz perspektive studenata, "
               "analizu dobrih i loših strana studija u cjelini iz studentske perspektive te, u skladu s time, "
               "planiranje mogućih mjera unaprjeđenja kvalitete."))
        self.p(_esc(f"Anketni upitnik sastoji se od {len(s['sections'])} cjelina, a svaka cjelina odnosi se na jedan "
                    f"vidik {s['level_gen']} studija:"))
        for sec in s["sections"]:
            self.p(_esc(sec), "list")
        self.story.append(Spacer(1, 0.2 * cm))
        self.p(_esc(f"Anketa sadrži ukupno {s['n_questions']} pitanja. " + AUTHORS))
        self.p("1.2. Ciljna populacija i način primjene", "h2")
        self.p(_esc(s["population"].format(year=self.year)))
        self.story.append(PageBreak())

    def sample(self):
        s, g, n = self.spec, self.gender, self.n
        nm, nz, nn = int((g == 1).sum()), int((g == 0).sum()), int(g.isna().sum())
        program = self.info.get("Studijski program", "")
        self.p("2. Rezultati ankete", "h1")
        self.p("2.1. Struktura uzorka prema anketiranim jedinicama i studijima", "h2")
        self.p(_esc(f"Studij {program} ({s['level']} studij) obuhvaća {n} anketnih upitnika. Kriterij za ispis "
                    f"izvješća je minimalno {MIN_SURVEYS} ispunjenih anketnih upitnika."))
        self.p(_esc(f"U tablici 1. prikazan je broj anketnih upitnika po anketiranoj jedinici i spolu anketiranih "
                    f"studenata. U anketi je "
                    f"sudjelovalo {n} studenata: {np_text(nm, n)} muških, {np_text(nz, n)} ženskih, a za "
                    f"{np_text(nn, n)} spol nije označen."))
        self.no = 1
        cells = [Paragraph(_esc(program), self.st["cell_s"])]
        for c in (nm, nz, nn, n):
            cells += [str(c), fmt(c / n * 100, 1) + "%"]
        data = [[Paragraph(_esc(self.info.get("Fakultet", "") or "Studij"), self.st["cell_sb"]),
                 "M", "", "Ž", "", "Nedostaje", "", "Ukupno", ""],
                ["", "N", "%", "N", "%", "N", "%", "N", "%"], cells]
        spans = [((0, 0), (0, 1)), ((1, 0), (2, 0)), ((3, 0), (4, 0)), ((5, 0), (6, 0)), ((7, 0), (8, 0))]
        self.story.append(KeepTogether([
            self.caption("Tablica 1. Broj anketiranih studenata po anketiranoj jedinici"),
            _grid_table(data, [4.2 * cm] + [1.25 * cm, 1.45 * cm] * 4, header_rows=2, spans=spans, bold_last=False),
            Spacer(1, 0.3 * cm)]))

    def a_items(self):
        self.p(_esc(f"2.2. Podaci o studentici/studentu i završenom {self.spec['level_gen']} studiju"), "h2")
        for item in self.spec["a_items"]:
            self.no += 1
            clean = self.col(item["col"], [v for v, _ in item["cats"]])
            rows = self.item_rows(clean, item["cats"])
            n, miss = self.n, int(clean.isna().sum())
            title = f"{item['code']} {item['text']}"
            texts = [f"U tablici {self.no}. i na grafikonima {self.no}a. i {self.no}b. prikazana je distribucija "
                     f"odgovora na pitanje „{item['code']} {item['text']}”, ukupno i po spolu. Od ukupno {n} anketnih "
                     f"upitnika, za {np_text(miss, n)} nedostaje odgovor."]
            if item.get("lead"):
                texts.insert(0, item["lead"])
            if item["kind"] == "grade":
                groups = [((4.5, 5), "odličan (4,5 ili 5)"), ((3.5, 4), "vrlo dobar (3,5 ili 4)"),
                          ((2.5, 3), "dobar (2,5 ili 3)"), ((2,), "dovoljan (2)")]
                parts = [f"{np_text(int(clean.isin(vals).sum()), n)} {lbl}" for vals, lbl in groups]
                texts.append(f"Prema vlastitoj procjeni, prosjek ocjena tijekom {item['what']} je kod "
                             f"{parts[0]}, kod {parts[1]}, kod {parts[2]} te kod {parts[3]}.")
            elif item["kind"] == "first_choice":
                c = [int((clean == v).sum()) for v, _ in item["cats"]]
                texts.append(f"Potvrdno je odgovorilo {np_text(c[0], n)} studenata, „uglavnom da” {np_text(c[1], n)}, "
                             f"a {np_text(c[2], n)} navodi da studij koji su završili nije bio njihov prvi izbor.")
            else:
                parts = [f"{np_text(int((clean == v).sum()), n)} {ph}"
                         for (v, _), ph in zip(item["cats"], item["phrases"])]
                texts.append("Od ukupnog broja anketiranih studenata " + _join(parts) + ".")
            block = [Paragraph(_esc(" ".join(texts)), self.st["body"]),
                     self.caption(f"Tablica {self.no}. {_esc(title)}"), self.gender_table(rows), Spacer(1, 0.3 * cm)]
            self.story.append(KeepTogether(block))
            self.story.append(KeepTogether(self.gender_charts(rows, _esc(title)) + [Spacer(1, 0.4 * cm)]))

    def groups(self):
        self.story.append(PageBreak())
        self.p("2.3. Pregled po skupinama pitanja", "h2")
        self.p(_esc(METHOD_TEXT))
        k, sub, current_parent = 0, 0, None
        for grp in self.spec["groups"]:
            self.no += 1
            if grp["parent"]:
                if grp["parent"] != current_parent:
                    k, sub, current_parent = k + 1, 0, grp["parent"]
                    self.p(f"2.3.{k}. {_esc(PARENTS[grp['parent']])} ({grp['parent']})", "h3")
                sub += 1
                number = f"2.3.{k}.{sub}."
            else:
                k, current_parent = k + 1, None
                number = f"2.3.{k}."
            self.group(number, grp)
        self.single_number = f"2.3.{k + 1}."

    def group(self, number, grp):
        st, n, no = self.st, self.n, self.no
        scale = grp["scale"]
        words = SATISFACTION_TEXT if scale is SATISFACTION else EXTENT_TEXT
        cols = {code: self.col(c, [v for v, _ in scale]) for c, code, _ in grp["items"]}
        frame = pd.DataFrame(cols)
        pooled = pd.concat(list(cols.values()), ignore_index=True)
        total = len(pooled)
        stats_all = describe(pooled)
        complete = int(frame.notna().all(axis=1).sum())
        counts_all = [int(pooled.isna().sum())] + [int((pooled == v).sum()) for v, _ in scale]
        first, last = grp["items"][0][1], grp["items"][-1][1]

        self.p(f"{number} {_esc(grp['title'])} ({grp['code']})", "h3")
        parts = [f"{np_text(c, total)} {w}" for c, w in zip(counts_all[1:], words)]
        self.p(_esc(
            f"Skup pitanja vezanih uz {grp['subject']} ({first} – {last}) prikazan je u tablicama {no}a. (opisna "
            f"statistika) i {no}b. (distribucija odgovora), a na grafikonima {no}a. i {no}b. prikazana je "
            f"distribucija svih odgovora u skupu, ukupno i po spolu. Sve odgovore u ovom skupu ispunilo je "
            f"{np_text(complete, n)} anketiranih studenata. Prosjek svih procjena je {fmt(stats_all['M'], 2)}, a "
            f"standardna devijacija {fmt(stats_all['SD'], 2)}. Od ukupno {total} odgovora (100%) nedostaje ih "
            f"{np_text(counts_all[0], total)} (neispunjeno, pogrešno ispunjeno ili „ne mogu procijeniti”), dok je "
            f"{_join(parts)}."))

        # Tablica a: opisna statistika
        data = [["Pitanje", "N", "NN", "M", "SD", "D", "Min", "Max"]]
        for c, code, text in grp["items"]:
            data.append(base._stats_row(Paragraph(_esc(f"{code} {text}"), st["cell_s"]), describe(cols[code])))
        data.append(base._stats_row(Paragraph("Ukupno za skup pitanja", st["cell_sb"]), stats_all))
        t_a = _grid_table(data, [8.0 * cm] + [1.1 * cm] * 7, header_rows=1, font_size=9)
        self.story += [self.caption(f"Tablica {no}a. Skup pitanja vezanih uz {_esc(grp['subject'])} - opisna "
                                    "statistika"), t_a, Spacer(1, 0.4 * cm)]

        # Tablica b: distribucija (postoci od ukupnog broja odgovora na pitanje)
        heads = [MISSING_LABEL] + [str(v) for v, _ in scale]
        head1 = ["Pitanje"] + sum([[h, ""] for h in heads], []) + ["Ukupno"]
        head2 = [""] + ["N", "%"] * len(heads) + ["N"]
        data = [head1, head2]
        for c, code, _ in grp["items"]:
            s = cols[code]
            cnt = [int(s.isna().sum())] + [int((s == v).sum()) for v, _ in scale]
            data.append([code] + sum([[str(x), fmt(x / n * 100, 1) + "%"] for x in cnt], []) + [str(n)])
        data.append(["Ukupno"] + sum([[str(x), fmt(x / total * 100, 1) + "%"] for x in counts_all], []) + [str(total)])
        spans = [((0, 0), (0, 1)), ((13, 0), (13, 1))] + [((1 + 2 * i, 0), (2 + 2 * i, 0)) for i in range(len(heads))]
        t_b = _grid_table(data, [1.35 * cm] + [0.95 * cm, 1.2 * cm] * len(heads) + [1.15 * cm], header_rows=2,
                          spans=spans, font_size=7.8)
        self.story += [self.caption(f"Tablica {no}b. Skup pitanja vezanih uz {_esc(grp['subject'])} - distribucija "
                                    "odgovora"), t_b, Spacer(1, 0.4 * cm)]

        # Grafikoni a i b (svi odgovori u skupu)
        labels = [MISSING_LABEL] + [lbl for _, lbl in scale]
        gender_long = pd.concat([self.gender] * len(cols), ignore_index=True)
        masks = [pooled.isna()] + [pooled == v for v, _ in scale]
        pct = [m.sum() / total * 100 for m in masks]
        m_pct = [(m & (gender_long == 1)).sum() / total * 100 for m in masks]
        z_pct = [(m & (gender_long == 0)).sum() / total * 100 for m in masks]
        title = f"Skup pitanja vezanih uz {_esc(grp['subject'])}"
        self.story += [KeepTogether([_pct_chart(labels, pct),
                                     self.caption(f"Grafikon {no}a. {title} - distribucija odgovora")]),
                       Spacer(1, 0.3 * cm),
                       KeepTogether([_grouped_chart(labels, m_pct, z_pct),
                                     self.caption(f"Grafikon {no}b. {title} - distribucija odgovora po spolu")]),
                       Spacer(1, 0.5 * cm)]

    def single_items(self):
        s, n = self.spec, self.n
        self.p(f"{self.single_number} Opća procjena ishoda (F) - pojedinačna pitanja", "h3")

        for item in s["yes_no"]:
            self.no += 1
            clean = self.col(item["col"], [1, 2])
            rows = self.item_rows(clean, YES_NO)
            yes, no_ = int((clean == 1).sum()), int((clean == 2).sum())
            yes_t, no_t = YES_NO_SENTENCES[item["key"]]
            title = f"{item['code']} {item['text']}"
            text = (f"U tablici {self.no}. i na grafikonima {self.no}a. i {self.no}b. prikazana je distribucija "
                    f"odgovora na pitanje „{title}”. Od ukupnog broja anketiranih ({n}), za "
                    f"{np_text(int(clean.isna().sum()), n)} nedostaje odgovor, {np_text(yes, n)} {yes_t}, a "
                    f"{np_text(no_, n)} {no_t}.")
            self._single_block(text, title, rows)

        # Odgovornost za nepripremljenost (odgovaraju samo oni koji se osjećaju nepripremljeno)
        self.no += 1
        col, code = s["responsibility"]
        clean = self.col(col, [v for v, _ in RESPONSIBILITY])
        rows = self.item_rows(clean, RESPONSIBILITY, missing_label=RESPONSIBILITY_PREPARED)
        c = [int(m.sum()) for _, m in rows]
        title = f"{code} Čija je odgovornost da se osjećate nepripremljeno za rad u struci?"
        text = (f"U tablici {self.no}. i na grafikonima {self.no}a. i {self.no}b. prikazana je distribucija odgovora na "
                f"pitanje „{title}”. Na ovo pitanje odgovaraju samo studenti koji se osjećaju nepripremljenima za rad "
                f"u struci, pa se studenti bez odgovora smatraju pripremljenima. Od ukupnog broja anketiranih ({n}), "
                f"{np_text(c[0], n)} osjeća se pripremljeno za rad u struci, {np_text(c[1], n)} odgovornost pripisuje "
                f"uglavnom studiju, {np_text(c[2], n)} dijelom sebi, a dijelom studiju, a {np_text(c[3], n)} "
                f"uglavnom sebi.")
        self._single_block(text, title, rows)

        # Opća ocjena studija
        self.no += 1
        col, code, qtext = s["overall"]
        labels = s.get("overall_labels", OVERALL)
        clean = self.col(col, [v for v, _ in labels])
        rows = self.item_rows(clean, labels)
        d = describe(clean)
        parts = [f"{np_text(int((clean == v).sum()), n)} ocjenom {v} ({lbl.split(' (')[0].lower()})"
                 for v, lbl in labels]
        title = f"{code} {qtext}"
        text = (f"Od ukupnog broja anketiranih ({n}) svoj {s['level']} studij ocjenjuje {_join(parts)}. Prosječna "
                f"opća ocjena kojom studenti ocjenjuju svoj {s['level']} studij iznosi {fmt(d['M'], 2)} "
                f"(SD = {fmt(d['SD'], 2)}). Na pitanje nije odgovorilo {np_text(d['NN'], n)} studenata.")
        self._single_block(text, title, rows)

    def _single_block(self, text, title, rows):
        self.story.append(KeepTogether([Paragraph(_esc(text), self.st["body"]),
                                        self.caption(f"Tablica {self.no}. {_esc(title)}"),
                                        self.gender_table(rows), Spacer(1, 0.3 * cm)]))
        self.story.append(KeepTogether(self.gender_charts(rows, _esc(title)) + [Spacer(1, 0.5 * cm)]))

    def build(self):
        if self.n == 0:
            raise ValueError("Nema anketa za izvještaj.")
        if self.n < MIN_SURVEYS:
            self.warnings.append(f"Samo {self.n} anketa - kriterij za ispis izvješća je najmanje {MIN_SURVEYS}.")
        self.cover()
        self.introduction()
        self.sample()
        self.a_items()
        self.groups()
        self.single_items()
        buf = io.BytesIO()
        doc = SimpleDocTemplate(
            buf, pagesize=A4, leftMargin=2.2 * cm, rightMargin=2.2 * cm, topMargin=2.0 * cm, bottomMargin=2.0 * cm,
            title=f"Vrjednovanje studija - {self.info.get('Studijski program', '')}",
            author="Ured za upravljanje kvalitetom, Sveučilište u Zagrebu",
        )
        doc.build(self.story, onFirstPage=lambda c, d: None, onLaterPages=base._page_number)
        return buf.getvalue(), self.warnings


def make_builder(form_type):
    spec = LEVELS[form_type]

    def build(df, info, semester, academic_year, report_date=None):
        return _Builder(spec, df, info, academic_year, report_date or date.today()).build()
    return build
