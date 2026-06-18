# Globale Vorgaben (gelten für alle Projekte)

## Hausstil für Excel-Arbeitsmappen (Standard)

Wenn ich eine **Excel-Mappe / strukturierte Tabelle** erstelle, gilt standardmäßig dieser Stil und Aufbau
(Vorbild: das FBH-Auslegungs-Tool). Vor dem Bauen kurz bestätigen, sonst direkt so umsetzen.

**Technik:** Python + `openpyxl` (kein VBA/Makro – Sicherheitsvorgabe). Wiederverwendbare Palette + Helfer:
`~/.claude/templates/agn_xlsx_style.py` (importieren oder die Helfer in das Build-Skript kopieren).
Workflow: Build-Skript → Mappe; zum Prüfen via LibreOffice nach PDF (`soffice --headless --convert-to pdf`)
und mit PyMuPDF rendern. Muss in **Excel UND LibreOffice** funktionieren.

**Design (Graustufen):**
- Graustufen; Schrift immer **schwarz**. Farbe NUR im **Logo** und in der **Bewertung** (grün = im Rahmen
  `#1E8E1E`, rot = Grenzwert `#D32F2F`) – und zwar als **Schriftfarbe**, nicht als Zellhintergrund.
- Tabellenkopf einheitlich **kräftiges Grau `#D2D2D2`**, schwarze fette Schrift (nie weiße Schrift),
  darunter eine **schwarze Linie**. Kein Unterschied Eingabe-/Rechenspalte im Kopf.
- **Berechnete Zellen leicht grau `#ECECEC`**, **Eingabe-/editierbare Zellen weiß**. Haarlinien `#C9C9C9`
  statt Vollgitter.
- **Zebra** (jede zweite Zeile leicht abgesetzt) auf allen Datentabellen – **filterfest** via
  `MOD(SUBTOTAL(103,$<key>$<top>:$<key><row>),2)=0` (zählt nur sichtbare Zeilen). Bewertungs- und Zebra-
  Regel **in einer kombinierten CF-Regel** setzen (Excel/LibreOffice wenden je Zelle nur die oberste Regel an).
- **Briefkopf** je Blatt: Titel links schwarz, agn-Logo rechts, Projektzeile, dünne schwarze Trennlinie.
- **Blattschutz:** nur Formelzellen sperren (Eingaben editierbar), ohne Passwort; Zeilenhöhe/Spaltenbreite
  und Filtern trotz Schutz erlauben.
- Druck: A4 quer, auf Seitenbreite skaliert, Kopf als Wiederholzeile, Fußzeile mit Seite/Datum.

**Aufbau (Blattreihenfolge) – Standard für umfangreichere Mappen:**
Deckblatt · Anleitung · Grundeinstellungen · (Daten/Auslegung) · Kontrolle · (Auswertungen) ·
(Verifikation) · Methodik · Konstanten · **Changelog** (Versionsliste, neueste oben; Version je Mappe pflegen).

**Stolperfallen:** dxf-Füllungen mit `FF`-Alpha-Präfix; Zelltexte nie mit `=` beginnen (sonst Fehler:501);
Funktionen ab Excel 2007 (TEXTJOIN …) im XML mit `_xlfn.`-Präfix; keine Spill/UNIQUE-Funktionen.

**Versionierung & Repo:** Build-Skript + Mappe versionieren; sinnvolle Commits; bei GitHub-Projekten pushen.
Logo/Corporate-Design = IP → private Repos.
