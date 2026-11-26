# Studie-overzicht

Een lokale startpagina voor studenten **Systeem- en netwerkbeheer (SNE) aan PXL-Digital**. Je lessenrooster, deadlines, toetsen en schoolmails staan samen op één pagina. Per vak open je met één klik de cursus, Blackboard en de juiste programma's.

Alles draait op je eigen laptop. Er gaat niets naar een externe server. Het enige wat van het internet komt, is je rooster uit TimeEdit.

## Wat je krijgt

- **Dashboard** op `http://127.0.0.1:48765/`:
  - je dag als tijdlijn
  - vakknoppen
  - deadlines en toetsen per vak
  - schoolmails van de laatste 7 dagen
  - de komende 7 dagen
- **Deadlines uit Blackboard in je Outlook-agenda.** Ze komen uit de Blackboard-mails "Dagelijkse meldingen", in de rode categorie *Deadline*, met een herinnering 2 dagen vooraf. Zodra je mail "Inzending ontvangen" binnenkomt, wordt de opdracht gemarkeerd als *ingediend*.
- **Briefing bij aanmelden** (maximaal 1x per dag) en elk uur een stille update.
- **Weekoverzicht** van de volgende week, elke zondag om 20:00.
- **Vakknoppen**, op het dashboard en in het Startmenu (map *Studie-overzicht vakken*).

## Wat je nodig hebt

- Windows 10 of 11 met PowerShell (zit standaard in Windows)
- **Klassieke Outlook** (Microsoft 365-apps), aangemeld met je `@student.pxl.be`-account. De nieuwe Outlook mag je gewoon blijven gebruiken, zolang de klassieke er ook op staat: het script leest je mail en agenda via de klassieke Outlook.
- Je **TimeEdit-abonnementslink** (zie hieronder)

## Installeren

1. **Download** deze repository: groene knop *Code* > *Download ZIP*. Pak de ZIP uit naar bijvoorbeeld `Documenten\Studie-overzicht`. Of gebruik git:
   ```powershell
   git clone https://github.com/SeMDeC/Studie-overzicht.git "$env:USERPROFILE\Documents\Studie-overzicht"
   ```
2. **Zoek je TimeEdit-link.** Open je rooster in TimeEdit en kies *Abonneren*. Kopieer de link: die begint met `https://cloud.timeedit.net/` en eindigt op `.ics`.
3. **Open PowerShell in die map** en voer uit (met jouw link):
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\Installeer.ps1 -TimeEditLink "https://cloud.timeedit.net/be_pxl/web/stud/....ics"
   ```

Na een halve minuut opent je dashboard. Zet een bladwijzer op **http://127.0.0.1:48765/**.

> Je TimeEdit-link is persoonlijk: deel hem niet. Hij komt in `config.json`, en dat bestand staat in `.gitignore` zodat het nooit mee in git gaat.

## Vakknoppen instellen

Maak in je browser (Brave, Chrome of Edge) een bladwijzermap **Online cursussen** met de online cursussen van je vakken. Geef elke bladwijzer de naam van het vak, bijvoorbeeld `Networks 1`, `Linux Essentials` of `Windows Essentials`. De knoppen lezen die map op het moment dat je klikt, dus nieuwe bladwijzers worden meteen meegenomen.

Welke programma's en welke Blackboard-cursus bij elk vak horen, staat in `vakken.json`:

| Vak | Programma's |
|---|---|
| Networks 1 | Packet Tracer, Wireshark, Windows Terminal |
| Linux Essentials | VMware Workstation |
| Windows Essentials | VMware Workstation |
| Werkplekleren 1, IT Organisation 1, Communication Skills 1 | alleen Blackboard |

- **Ander vak, programma of Blackboard-cursus?** Pas `vakken.json` aan en voer `.\Installeer.ps1` opnieuw uit.
- **Blackboard-cursus:** het ID staat in de adresbalk wanneer je de cursus op Blackboard opent, bijvoorbeeld `_65586_1`. De ID's in deze repo zijn die van **1SNE, academiejaar 2026-27**.
- **Browser:** laat `"Browser"` leeg om je standaardbrowser te gebruiken. Wil je een bepaalde browser, vul dan het volledige pad naar dat programma in.

Controleer wat elke knop zou openen:

```powershell
.\OpenVak.ps1 -Lijst
```

## Je eigen deadlines en toetsen

Zet een afspraak in Outlook met de categorie **Deadline** of **PE / toets**, en ze komt vanzelf op het dashboard. Zet de vaknaam in de titel, bijvoorbeeld `Deadline: Labo 3 (Networks 1)`, dan staat ze onder het juiste vak.

## Wat de installatie op je laptop verandert

Alles hieronder geldt alleen voor jouw eigen Windows-account. Je hebt er geen beheerdersrechten voor nodig.

- **Geplande taken** in de Taakplanner, onder de map *Studie-overzicht*:
  - de briefing bij aanmelden
  - de update van elk uur
  - het weekoverzicht op zondag
  - de *Vakknoppen-server*, een kleine lokale webserver op `127.0.0.1:48765`. Die is alleen bereikbaar vanaf je eigen laptop en weigert verzoeken van andere websites.
- **Snelkoppelingen** in het Startmenu.
- **Outlook:** de categorieën *Deadline* en *PE / toets*, en afspraken voor de deadlines uit Blackboard. Verwijder je zo'n afspraak, dan wordt ze niet opnieuw aangemaakt.

## Aanpassen

In `config.json` stel je onder andere in:

| Instelling | Wat |
|---|---|
| `HerinneringMinutenVooraf` | 2880 = 2 dagen |
| `DeadlinesInOutlook` | zet op `false` om niets in je agenda te schrijven |
| `SignaalWoorden` | woorden waardoor een mail het label "Lijkt belangrijk" krijgt |
| `SchoolAfzenders` | welke afzenders als schoolmail tellen |

Wat er gebeurt, zie je in `data\log.txt`.

## Verwijderen

```powershell
powershell -ExecutionPolicy Bypass -File .\Installeer.ps1 -Verwijder
```

Dat verwijdert de geplande taken, de snelkoppelingen en de vakknoppen. Deadlines die al in je Outlook-agenda staan, blijven daar staan. Daarna mag je de map gewoon wissen.

## Hoe het werkt

| Bestand | Wat het doet |
|---|---|
| `StudieOverzicht.ps1` | Leest TimeEdit en je Outlook-mail, zet deadlines in je agenda en maakt `dashboard.html` en `weekoverzicht.html` |
| `OpenVak.ps1` | Opent de cursussen, Blackboard en programma's van een vak |
| `Server.ps1` | De lokale webserver voor het dashboard en de vakknoppen |
| `Installeer.ps1` | Installeren en verwijderen |
