# FBH-Auslegung – Auslegung von Fußbodenheizungen nach DIN EN 1264

Ein Excel-Werkzeug für die Auslegung von Fußbodenheizungen. Raumweise werden
Heizleistung, Heizlastdeckung, Rohrlängen, Massen-/Volumenströme und Druckverluste
ermittelt und bewertet (Status ✓ / !).

> Entwicklungsstand: **0.x (Vorab)**. Die aktuelle Version steht im Blatt *Changelog*.

## Inhalt des Repositories

**Hauptvariante – vollständige Berechnung nach DIN EN 1264 (Bauart A):**

| Datei | Zweck |
|-------|-------|
| `FBH_Auslegung_EN1264.xlsx` | Die fertige Arbeitsmappe (direkt nutzbar) |
| `build_fbh_en1264.py` | Erzeugt die Arbeitsmappe neu (Python + openpyxl) |
| `verify_fbh.py` | Prüft die Rechenlogik unabhängig nach |
| `agn-logo.png` | In die Mappe eingebettetes Logo |

**Archiv** (`archiv/`): die ältere, vereinfachte Variante mit Systemfaktor `f`
(`build_fbh_vereinfacht.py`, `FBH_Auslegung_vereinfacht.xlsx`) – nicht mehr aktiv gepflegt.

## Tabellenblätter der Mappe

Deckblatt · Anleitung · Grundeinstellungen · Auslegung · Kontrolle · HKV (Heizkreisverteiler) ·
Verifikation · Methodik · Konstanten · Changelog

## Arbeitsmappe neu erzeugen

Voraussetzung: Python 3 mit `openpyxl`.

```bash
pip install openpyxl
python build_fbh_en1264.py   # erzeugt FBH_Auslegung_EN1264.xlsx
python verify_fbh.py         # optional: Rechenlogik prüfen
```

Alle Eingaben (Temperaturen, Rohrsystem, Konstanten, Raumdaten …) werden direkt in
der Excel-Datei gemacht; das Build-Skript ist nur nötig, wenn die Struktur/Vorlage
geändert werden soll.

## Methodik (Kurzfassung)

- Spez. Heizleistung nach dem EN-1264-Produktverfahren (Bauart A):
  `q = B · a_B · a_Tᵐᵀ · a_üᵐü · a_Dᵐᴰ · ΔθH` (logarithmische Übertemperatur)
- Faktoren `a_B`, `a_T` (je R-Wert) sowie `B`, `a_ü`, `a_D` editierbar im Blatt *Konstanten*
- Druckverlust nach Darcy-Weisbach; maßgebliche Leistung = `min(Leistung; Heizlast) + Verlust nach unten`
- Zonenabhängige Oberflächentemperatur-Grenzwerte; Abgleich mit Herstellerdaten im Blatt *Verifikation*

Details und alle Formelzeichen stehen im Blatt *Methodik*; die angewandten
Vereinfachungen sind dort und auf dem *Deckblatt* dokumentiert.

## Hinweis

Überschlägiges Werkzeug – ersetzt **nicht** die TGA-Detailplanung (vollständiges
EN-1264-Verfahren, hydraulischer Abgleich, Herstellernachweise).
