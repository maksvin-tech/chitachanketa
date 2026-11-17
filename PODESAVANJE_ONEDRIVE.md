# Spremanje u OneDrive (poslovni račun @unizg.hr)

Program sprema rezultate (Excel), skenirane ankete i PDF izvještaje u **OneDrive** računa
mvinscak@unizg.hr, u mapu **OMR Ankete**. Koristi samo OneDrive tog računa (ovlast *Files.ReadWrite*);
ne čita e-poštu, kalendar ni Teams.

Aplikacija **OMR Ankete** registrirana je u Microsoft Entri Sveučilišta u Zagrebu
(samo za taj tenant, „Allow public client flows” uključeno). Ništa se više ne podešava.

## Prijava (jednom)
1. U bočnoj traci, pod **☁️ OneDrive**, kliknite **🔑 Prijavi se u OneDrive**.
2. Program pokaže kratki kod. Otvorite <https://microsoft.com/devicelogin> i upišite ga.
3. Prijavite se računom **mvinscak@unizg.hr** i prihvatite pristup aplikaciji *OMR Ankete*.
4. Vratite se u program i kliknite **Završi prijavu**.

Na poslužitelju prijava vrijedi za sva računala i sama se obnavlja. Ponovna prijava treba tek ako
se program ne koristi 90 dana ili se promijeni lozinka. Lozinka nikad ne prolazi kroz program.

Ako piše da aplikaciju treba odobriti administrator (*admin consent*), to mora napraviti
Srce / informatičari Sveučilišta za aplikaciju *OMR Ankete* (ID `81b53565-04b2-4d78-9d51-a9b30fca601f`).

## Struktura u OneDriveu
```
OMR Ankete/
  Rezultati/<vrsta obrasca>/<fakultet>/<studijski program ili nastavnik>/<datum_vrijeme>_<računalo>_<naziv>.xlsx
  Skenovi/<vrsta obrasca>/<fakultet>/<studijski program ili nastavnik>/<datum_vrijeme>_<računalo>_<naziv>/<slike>
  Izvještaji/<vrsta obrasca>/<fakultet>/<izvještaj>.pdf
  Za fakultete/<fakultet>/<ak. god.>/<vrsta>/...      <- mapa koja se šalje fakultetu
  poslano.json                                        <- koji je fakultet već dobio mapu (zajedničko svim računalima)
```
Studijski program je mapa s kodom i nazivom (npr. `0190 Geoinženjerstvo`), nastavnik mapa s imenom.
Nemojte ručno premještati ili preimenovati mape – program ih traži po imenu.

Serije spremljene prije 17.11.2026. bile su složene samo po datumu. Jednokratno preuređenje
(samo premješta, ništa ne briše; sigurno ga je pokrenuti više puta; `--proba` samo prikaže plan):
```
sudo -u omr OMR_DATA_DIR=/var/lib/omr-ankete OMR_SERVER=1 \
    /opt/omr-ankete-venv/bin/python /opt/omr-ankete/preuredi_onedrive.py
```

Dijeli se **samo** ono što se šalje fakultetu: program stvori poveznicu *Svatko s vezom može gledati*
tek kad se pripremi e-mail. Sve ostalo vidi samo vlasnik OneDrivea.

## Prijenos starih podataka iz Google Drivea (jednom)
Na poslužitelju, nakon prijave u OneDrive:
```
sudo -u omr OMR_DATA_DIR=/var/lib/omr-ankete OMR_SERVER=1 \
    /opt/omr-ankete-venv/bin/python /opt/omr-ankete/prenesi_iz_google_drivea.py
```
Prenese cijelu mapu *OMR Ankete* (preskače što već postoji); Google Drive ostaje netaknut.

## Odjava
U bočnoj traci **Odjava s OneDrivea** briše spremljenu prijavu.
Pristup aplikaciji možete opozvati na <https://myapps.microsoft.com> → *Upravljanje aplikacijama*.

## Kontakti fakulteta (za e-mail)

Tablicu *Kontakti fakulteta – studentske ankete* preuzmite kao `.xlsx` i učitajte je u bočnoj traci,
pod **📇 Kontakti fakulteta** → *Spremi kontakte*. Spremi se kao `kontakti.csv` u mapu podataka
(ne ide na GitHub). Kad se tablica promijeni, učitajte je ponovno.
Postupak rada sa skenovima, izvještajima i slanjem: **UPUTE_ZA_STUDENTE.md**.
