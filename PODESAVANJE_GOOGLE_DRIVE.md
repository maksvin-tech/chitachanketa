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
