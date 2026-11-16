# Podešavanje spremanja u Google Drive

Program sprema rezultate (Excel) i izvorne skenove u vaš Google Drive, u mapu **OMR Ankete**.
Podešavanje se radi **jednom** (koraci 1–5), a prijava jednom **na svakom računalu** (korak 7).

> Program dobiva ovlast samo za datoteke koje je **sam stvorio** (`drive.file`).
> Ne vidi ostatak vašeg Drivea, e-poštu ni fotografije.

## 1. Napravite projekt u Google Cloud konzoli
1. Otvorite <https://console.cloud.google.com> i prijavite se **istim Google računom** na kojem je vaš Google AI Pro (Drive).
2. Gore lijevo kliknite izbornik projekata → **New project** → ime npr. `OMR Ankete` → **Create**.
   Ovo je besplatno, kartica nije potrebna.

## 2. Uključite Google Drive API
1. Izbornik → **APIs & Services** → **Library**.
2. Potražite **Google Drive API** → **Enable**.

## 3. Postavite zaslon za prijavu
1. Izbornik → **Google Auth Platform** (u starijem sučelju: *APIs & Services → OAuth consent screen*) → **Get started**.
2. App name: `OMR Ankete`, User support email: vaš e-mail.
3. Audience: **External**.
4. Contact information: vaš e-mail → prihvatite uvjete → **Create**.

## 4. Objavite aplikaciju (važno!)
1. U **Google Auth Platform → Audience** kliknite **Publish app** → **Confirm**.
2. Bez ovoga Google briše prijavu nakon 7 dana i morali biste se stalno ponovno prijavljivati.
   Za ovlast `drive.file` Google ne traži dodatnu provjeru.

## 5. Napravite OAuth klijenta i preuzmite ključ
1. **Google Auth Platform → Clients** → **Create client**.
2. Application type: **Desktop app**, ime: `OMR Ankete` → **Create**.
3. Kliknite **Download JSON**.
4. Preimenujte preuzetu datoteku u **`google_credentials.json`**.

## 6. Stavite ključ u mapu programa (na svakom računalu)
Kopirajte `google_credentials.json` u istu mapu gdje su `omr_aplikacija.py` i `Pokreni_OMR.bat`.

⚠️ **Nikada ne stavljajte ovu datoteku na GitHub** – repozitorij je javan.
Datoteka `.gitignore` to već sprječava, ali je nemojte ni ručno učitavati preko web stranice GitHuba.

Prvi put na svakom računalu instalirajte potrebne pakete:
```
pip install -r requirements.txt
```

## 7. Prva prijava
1. Pokrenite program (`Pokreni_OMR.bat`).
2. U lijevom izborniku kliknite **🔑 Prijavi se u Google Drive**.
3. Otvara se preglednik – odaberite svoj Google račun i odobrite pristup.
   Ako se pojavi poruka *"Google hasn't verified this app"*, kliknite **Advanced → Go to OMR Ankete** – to je vaša vlastita aplikacija.
4. Program sprema prijavu u `google_token.json` i sljedeći put se spaja sam.

## Kako se koristi
- Nakon očitavanja anketa i eventualnih ručnih ispravaka kliknite **☁️ Spremi u Google Drive**.
- Ako nakon spremanja još nešto ispravite u tablici, kliknite **☁️ Spremi izmjene u Google Drive** – zamijenit će se Excel te serije, skenovi se ne šalju ponovno.
- Na dnu stranice **📚 Svi spremljeni rezultati** učitava sve serije odabranog obrasca u jednu tablicu (s računala bilo kojeg kolege prijavljenog istim računom).

## Struktura u Google Driveu
```
OMR Ankete/
  Rezultati/<vrsta obrasca>/<datum_vrijeme>_<računalo>.xlsx
  Skenovi/<vrsta obrasca>/<datum_vrijeme>_<računalo>/<izvorne slike>
```
Nemojte ručno premještati ili preimenovati mape – program ih traži po imenu.
Datoteke koje sami dodate u te mape program neće vidjeti.

## Odjava
U lijevom izborniku **Odjava s Google Drivea** briše spremljenu prijavu s tog računala.
Pristup možete potpuno opozvati na <https://myaccount.google.com/permissions>.

## E-mail fakultetu uz gotov izvještaj

1. **Kontakti (jednom):** tablicu *Kontakti fakulteta – studentske ankete* preuzmite kao `.xlsx`
   (ili `.csv`) i učitajte je u bočnoj traci, pod **📇 Kontakti fakulteta** → *Spremi kontakte*.
   Program je sprema kao `kontakti.csv` uz `google_credentials.json` (ne ide na GitHub).
   Kad se tablica promijeni, učitajte je ponovno (*Zamijeni popis kontakata*).
2. **Izvještaj:** izradite PDF izvještaj (sprema se u Google Drive) i spremite seriju u Google Drive
   sa skenovima.
3. **E-mail:** ispod spremanja pojavi se **✉️ E-mail fakultetu**. Kontakti tog fakulteta za tu vrstu
   ankete označeni su ⭐. Odaberite primatelja, provjerite tekst (može se mijenjati) i kliknite
   **✉️ Pripremi e-mail** → **📧 Otvori u Outlooku**. Poruka se otvori u Outlooku na računalu, od računa
   koji je ondje prijavljen (mvinscak@unizg.hr); provjerite je i pritisnite *Pošalji*.
   - Izvještaj i mapa sa skenovima dijele se tako da ih **svatko s poveznicom može gledati**
     (kao *Dijeli → Svatko s vezom* u Driveu), umjesto WeTransfera.
   - Oslovljavanje (npr. *Poštovana kolegice Borošić,*) uzima se iz stupca *Oslovljavanje*;
     prazno = *Poštovani,*.
   - Ako Outlook sam dodaje potpis, isključite *Dodaj potpis*.
   - Ako se klikom ništa ne otvori, Outlook nije zadani program za e-poštu u Windowsima:
     *Postavke → Aplikacije → Zadane aplikacije → Outlook* (ili *MAILTO* → Outlook).
4. **Svi izvještaji odjednom:** kad izradite izvještaje za **jedan** fakultet, ispod se pojavi isti
   e-mail s poveznicom na mapu svih izvještaja tog fakulteta.

## 📦 Mapa za fakultet (izvještaji + skenovi, bez Excela)

Na dnu aplikacije: **Učitaj spremljene serije svih obrazaca** → odaberite fakultet i akademsku godinu →
provjerite serije sa skenovima → **📦 Pripremi mapu za fakultet**. Program složi i podijeli
(svatko s poveznicom može gledati) mapu:

    OMR Ankete/Za fakultete/<fakultet>/<2025-2026>/
        Prijediplomski studij/  <PDF izvještaji>   Skenirane ankete/<kod studij>/<skenovi>
        Diplomski studij/       <PDF izvještaji>   Skenirane ankete/<kod studij>/<skenovi>

Ispod se pojavi e-mail s poveznicom na tu mapu. Excel tablice ne idu u mapu. PDF izvještaje
prvo izradite (gore, *Izradi izvještaje* za taj fakultet). Ponovno pokretanje dodaje samo ono
čega još nema u mapi.
