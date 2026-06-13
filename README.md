# FBH-Auslegung – überschlägige Auslegung von Fußbodenheizungen

Ein Excel-Werkzeug für die **überschlägige** Auslegung von Fußbodenheizungen in der
frühen Planungsphase. Raumweise werden Heizleistung, Heizlastdeckung, Rohrlängen,
Massen-/Volumenströme und Druckverluste abgeschätzt und per Ampel bewertet.

> Entwicklungsstand: **0.x (Vorab)**. Version 1.0 erst nach Kalibrierung mit echten
> Herstellerdaten und Freigabe. Die aktuelle Version steht im Blatt *Changelog*.

## Inhalt des Repositories

| Datei | Zweck |
|-------|-------|
| `FBH_Auslegung.xlsx` | Die fertige Arbeitsmappe (direkt nutzbar) |
| `build_fbh.py` | Erzeugt die Arbeitsmappe neu (Python + openpyxl) |
| `verify_fbh.py` | Prüft die Rechenlogik unabhängig nach |
| `agn-logo.png` | In die Mappe eingebettetes Logo |

## Tabellenblätter der Mappe

Deckblatt · Anleitung · Grundeinstellungen · Auslegung · Kontrolle · HKV (Heizkreisverteiler) ·
Verifikation · Methodik · Konstanten · Changelog

## Arbeitsmappe neu erzeugen

Voraussetzung: Python 3 mit `openpyxl`.

```bash
pip install openpyxl
python build_fbh.py        # erzeugt FBH_Auslegung.xlsx
python verify_fbh.py       # optional: Rechenlogik prüfen
```

Alle Eingaben (Temperaturen, Rohrsystem, Konstanten, Raumdaten …) werden direkt in
der Excel-Datei gemacht; `build_fbh.py` ist nur nötig, wenn die Struktur/Vorlage
geändert werden soll.

## Methodik (Kurzfassung)

- Spez. Heizleistung: `q = f · η_VA · K_H · ΔθH` (logarithmische Übertemperatur)
- Druckverlust nach Darcy-Weisbach; maßgebliche Leistung = `min(Leistung; Heizlast) + Verlust nach unten`
- Zonenabhängige Oberflächentemperatur-Grenzwerte; Kalibrierung über den
  Korrekturfaktor `f` (Blatt *Verifikation*)

Details und alle Formelzeichen stehen im Blatt *Methodik*; die angewandten
Vereinfachungen sind dort und auf dem *Deckblatt* dokumentiert.

## Hinweis

Überschlägiges Werkzeug – ersetzt **nicht** die TGA-Detailplanung (vollständiges
EN-1264-Verfahren, hydraulischer Abgleich, Herstellernachweise).
