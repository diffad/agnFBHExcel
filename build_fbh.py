# -*- coding: utf-8 -*-
"""Aufbau der Excel-Arbeitsmappe zur ueberschlaegigen Auslegung von Fussbodenheizungen."""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.chart import LineChart, Reference

CHANGELOG = [
    ("0.1", "2026-06-07", [
        "Ersterstellung mit den Blättern Grundeinstellungen, Raumliste, Übersicht und Methodik.",
        "Heizleistung über die logarithmische Übertemperatur (q = η · K_H · ΔθH).",
        "Druckverlustberechnung nach Darcy-Weisbach.",
        "Editierbare Rohrbibliothek und Verlegeabstand-Faktoren.",
        "Prüfungen: Oberflächentemperatur, Heizlastdeckung, Kreislänge, Druckverlust.",
    ]),
    ("0.2", "2026-06-07", [
        "agn-Logo eingebunden und eigenes Changelog-Blatt ergänzt.",
        "Zuschlag Einzelwiderstände (5 %) und fester Verteiler-Aufschlag (500 Pa je Kreis).",
        "Einheiten als Zellformat; Projektnummer/Projektname erscheinen im Ausdruck.",
        "Richtwert-Tabellen für R-Werte und Verlegeabstand-Faktoren; weitere Musterräume.",
        "Druckaufbereitung für DIN A4.",
    ]),
    ("0.3", "2026-06-07", [
        "Zonentypen (Aufenthalt/Rand/Bad) mit eigener zulässiger Oberflächentemperatur.",
        "Einheitlicher Kopf: Titel oben links, agn-Logo oben rechts, Projektkopf darunter.",
        "Oberflächentemperatur θF und Volumenstrom eingeblendet.",
        "Neue Spalten 'spez. Druckverlust [Pa/m]' und Eingabespalte 'Verteiler (Anbindung)'.",
        "Warnung (rot), wenn die aktivierbare Fläche größer als die Raumfläche ist.",
    ]),
    ("0.4", "2026-06-07", [
        "Status-Spalten entfernt – Bewertung erfolgt jetzt direkt als Zellfärbung (grün/rot).",
        "Methodik um die Formelzeichen-Legende erweitert.",
        "Heizkreisverteiler (HKV) als erste Spalte der Auslegung; Standardwert max. Druckverlust 15.000 Pa.",
        "Versionsschema 0.x eingeführt (1.0 erst nach Test); interner Autor (dh).",
    ]),
    ("0.5", "2026-06-07", [
        "Versionsnummer und Autor werden nur noch im Changelog angezeigt.",
        "Editierbares Bearbeiter-Kürzel rechts oben unter dem Logo (Grundeinstellungen, Auslegung, Kontrolle, HKV).",
    ]),
    ("0.6", "2026-06-07", [
        "Bugfix: Zellfärbung (bedingte Formatierung) wird jetzt auch in Excel angezeigt – dxf-Füllfarbe mit voller Deckkraft.",
    ]),
    ("0.7", "2026-06-07", [
        "Deckung-%-Spalte färbt nur noch gefüllte Zellen (leere bleiben farblos).",
        "Neues Blatt 'Verifikation'; neue Spalte 'Strömungsbereich' (laminar/turbulent).",
        "Alle Blätter im A3-Querformat; einheitlicher Kopf je Blatt.",
    ]),
    ("0.8", "2026-06-07", [
        "Grundeinstellungen für A3 quer optimiert (Einstellungen links, Tabellen rechts).",
        "Auslegung passt auf eine Seite breit; Korrekturfaktor Heizleistung zur Kalibrierung ergänzt.",
        "Wärmeabgabe nach unten berücksichtigt (Verlust + Gesamtleistung).",
        "Blätter umbenannt: Raumliste → Auslegung, Übersicht → Kontrolle; zusätzliche Kontrollen.",
    ]),
    ("0.9", "2026-06-07", [
        "Neues Blatt 'Heizkreisverteiler' (Leistung, max. Druckverlust, Volumenstrom je HKV).",
        "Spaltenüberschriften der Auslegung dreizeilig (Name / Formelzeichen / Einheit).",
        "Korrekturfaktor f wird auf 'Verifikation' eingegeben (mit Vergleichsdiagramm).",
        "Geschwindigkeits-Grenzwerte v_min/v_max als Grundeinstellung.",
    ]),
    ("0.10", "2026-06-08", [
        "Heizkreisverteiler-Liste und Verifikation ohne dynamische Array-Funktionen → läuft in Excel UND LibreOffice.",
        "Verifikation mit Liniendiagramm; Konstanten/Bibliotheken in eigenes Blatt ausgelagert.",
        "Gewähltes Rohrsystem mit Außen-/Innendurchmesser, Wandstärke und Rauheit im Druckbereich.",
    ]),
    ("0.11", "2026-06-08", [
        "Bugfixes: Hinweistext wurde als Formel interpretiert (Fehler 501); MAXIFS ersetzt (#NAME? in LibreOffice).",
        "Verifikation: Korrekturfaktor prominent oben; Hersteller-q direkt neben q berechnet.",
        "Konstanten: alle Tabellen untereinander, mit Leerzeilen für weitere Einträge.",
        "Methodik: Variablen zuerst, dann Formeln mit Erklärung je Schritt; weitere Musterräume (Verwaltungsbau).",
    ]),
    ("0.12", "2026-06-08", [
        "Grundeinstellungen: Bezeichnungen ausgeschrieben (keine Abkürzungen).",
        "Neue Spalte 'spez. Heizlast [W/m²]' (Heizlast bezogen auf die Raumfläche).",
        "Massenstrom/Geschwindigkeit/Druckverlust auf Basis min(Leistung; Heizlast) + Verlust nach unten (realer Betriebspunkt).",
        "Heizkreisverteiler zusätzlich mit Massenstrom [kg/h]; Zonen-Kürzel (AZ/RZ/BAD).",
        "Kapazität 200 Räume / 50 Verteiler; Auslegung A3 quer, übrige Blätter A4 hoch.",
    ]),
    ("0.13", "2026-06-08", [
        "Strömungsbereich dreistufig: laminar (Re < 2300), Übergangsbereich (2300–4000), turbulent (Re > 4000).",
        "Spalte 'Strömungsbereich' ausgeblendet (für die Auslegung nicht relevant) und ohne Färbung.",
        "Auslegung schlanker: weitere Zwischenspalten standardmäßig ausgeblendet (jederzeit einblendbar).",
        "Verifikation: Korrekturfaktor kleiner dargestellt (passt nun vollständig in die Zelle).",
        "Changelog wieder ausführlich.",
    ]),
    ("0.14", "2026-06-09", [
        "Deckblatt als erstes Blatt: Übersichtsdiagramm Eingaben → Verarbeitung → Ergebnisse (inkl. Vereinfachungen).",
    ]),
    ("0.15", "2026-06-09", [
        "Deckblatt-Diagramm aus Excel-Zellen statt Bild – direkt in Excel bearbeitbar.",
        "Auslegung: Beispielräume mit gemischter Heizlast (≤ 50 W/m²); Formelzeichen in Zeile 5 ergänzt (Flächen u. a.).",
        "Heizkreisverteiler: zusätzliche Spalte Spreizung [K].",
    ]),
    ("0.16", "2026-06-09", [
        "Alle Blätter im A4-Querformat (nur Auslegung A3 quer); einheitlichere Schriftgrößen/Skalierung.",
        "Grundeinstellungen zweispaltig (Einstellwerte teilweise nebeneinander).",
        "Verifikation: Diagramm rechts neben den Tabellen.",
        "Methodik: Formelzeichen links, Berechnungsschritte in zwei Spalten rechts.",
        "Konstanten: Tabellen teilweise nebeneinander.",
        "Auslegung: Spalte Zone mittig/schmaler, Raum-Nr. breiter; Verifikation-Korrekturfaktor vertikal mittig.",
    ]),
    ("0.17", "2026-06-13", [
        "Neues Blatt 'Anleitung' (2. Blatt): Aufbau der Mappe, Schritt-für-Schritt, Farblegende.",
        "Anleitung enthält den Workflow zum Übernehmen der Räume aus Revit (Copy & Paste, Schedule, Dynamo/Power Query).",
    ]),
    ("0.18", "2026-06-13", [
        "Auslegung: Eingabespalten neu sortiert – zuerst die Revit-Spalten (A HKV, B Raum-Nr., C Bezeichnung, D Raumfläche, E Raumtemperatur, F Heizlast), dann G spez. Heizlast (berechnet), danach die manuellen Felder (H aktiv. Fläche, I R-Wert, J Verlegeabstand, K Anz. Kreise, L Zuleitung, M Zone).",
        "Anleitung: Excel-Export aus Revit ist jetzt der Standard-Workflow (direktes Kopieren funktioniert nicht); Spaltenliste an neue Reihenfolge angepasst.",
    ]),
    ("0.19", "2026-06-13", [
        "Grundeinstellungen: neuer Grenzwert 'Maximaler Volumenstrom je Kreis' (V̇_max, Default 150 l/h); Abschnitt 'Strömungsgeschwindigkeit' → 'Strömungsgrenzwerte'.",
        "Auslegung: Volumenstrom je Kreis wird gegen V̇_max geprüft und per Ampel eingefärbt (grün ≤ Grenze, rot darüber); zusätzliche Warnung im Blatt 'Kontrolle'.",
        "Anleitung: durchgehende Schritt-für-Schritt-Anleitung mit integriertem Revit-Export (vorgefertigte Bauteilliste muss nicht erst angelegt werden); redundanter Abschnitt 'Räume aus Revit' und der Hinweis zum direkten Kopieren entfernt.",
        "Deckblatt: Pfeil-Spalten breiter und Pfeilgröße angepasst (optisch sauberer).",
    ]),
]
VERSION = CHANGELOG[-1][0]
AUTHOR = "dh"
VERTXT = f"Version {VERSION} · {AUTHOR}"
LOGO = r"C:\Users\derDi\FBH_Auslegung_Excel\agn-logo.png"

FONT = "Arial"
NAME_ROW, SYM_ROW, UNIT_ROW = 4, 5, 6
R0, N_ROWS = 7, 200
R1 = R0 + N_ROWS - 1   # 206
EMU = 9525
BLUE, BLACK, WHITE, GREY, NAVY = "0000FF", "000000", "FFFFFF", "808080", "1F4E78"
HDR_FILL = PatternFill("solid", fgColor=NAVY)
SUB_FILL = PatternFill("solid", fgColor="D9E1F2")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
ACCENT_FILL = PatternFill("solid", fgColor="FFE699")
RED_FILL = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")
GREEN_FILL = PatternFill(start_color="FFC6EFCE", end_color="FFC6EFCE", fill_type="solid")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
CEN = Alignment(horizontal="center")
LEFT = Alignment(horizontal="left")

def f(bold=False, color=BLACK, size=10, italic=False):
    return Font(name=FONT, bold=bold, color=color, size=size, italic=italic)

def col_px(ws, idx):
    dim = ws.column_dimensions[get_column_letter(idx)]
    if dim.hidden: return 0
    w = dim.width if dim.width else 8.43
    return int(round(w * 7)) + 5

def add_logo_corner(ws, last_col, height=40):
    if not os.path.exists(LOGO): return
    img = XLImage(LOGO)
    ratio = img.width / img.height
    img.height = height; img.width = int(height * ratio)
    right_edge = sum(col_px(ws, i) for i in range(1, last_col + 1))
    left = max(0, right_edge - img.width - 4)
    acc, c = 0, 1
    while c <= last_col:
        cw = col_px(ws, c)
        if cw and acc + cw > left: break
        acc += cw; c += 1
    marker = AnchorMarker(col=c - 1, colOff=(left - acc) * EMU, row=0, rowOff=1 * EMU)
    img.anchor = OneCellAnchor(_from=marker, ext=XDRPositiveSize2D(cx=img.width * EMU, cy=img.height * EMU))
    ws.add_image(img)

def disp_header(ws, title, last_col, bearb_cell=None, bearb_merge=None, project=True):
    ws.row_dimensions[1].height = 34
    t = ws["A1"]; t.value = title
    t.font = Font(name=FONT, bold=True, color=NAVY, size=15); t.alignment = Alignment(horizontal="left", vertical="center")
    add_logo_corner(ws, last_col, height=40)
    if project:
        a2 = ws["A2"]; a2.value = f'="Projekt-Nr.: "&{gPNr}'; a2.font = f(bold=True, color=NAVY, size=11); a2.alignment = LEFT
        a3 = ws["A3"]; a3.value = f'="Projekt: "&{gPName}'; a3.font = f(bold=True, color=NAVY, size=11); a3.alignment = LEFT
    if bearb_cell:
        if bearb_merge: ws.merge_cells(bearb_merge)
        b = ws[bearb_cell]; b.value = f'="Bearbeiter: "&{gBearb}'
        b.font = f(italic=True, color=GREY, size=9); b.alignment = Alignment(horizontal="right", vertical="center")

def setup_print(ws, area, titles=None, orientation="landscape", paper=9):
    # paper: 9 = A4, 8 = A3. Auf eine Seite breit skaliert (fit-to-width).
    ws.page_setup.orientation = orientation
    ws.page_setup.paperSize = paper
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_area = area
    if titles: ws.print_title_rows = titles
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.oddFooter.left.text = "FBH überschlägige Auslegung"
    ws.oddFooter.center.text = "Seite &P von &N"
    ws.oddFooter.right.text = "Stand &D"

def mini_header(ws, row, col, text):
    c = ws.cell(row=row, column=col, value=text)
    c.font = f(bold=True, color=WHITE); c.fill = HDR_FILL
    c.alignment = Alignment(wrap_text=True, horizontal="center"); c.border = BORDER

def two_col_table(ws, title_row, c0, title, hdr1, hdr2, rows, fmt2, n_empty=0):
    ws.cell(row=title_row, column=c0, value=title).font = f(bold=True, color=NAVY)
    ws.merge_cells(start_row=title_row, start_column=c0, end_row=title_row, end_column=c0 + 1)
    mini_header(ws, title_row + 1, c0, hdr1); mini_header(ws, title_row + 1, c0 + 1, hdr2)
    top = title_row + 2
    total = len(rows) + n_empty
    for i in range(total):
        r = top + i
        a = ws.cell(row=r, column=c0); b = ws.cell(row=r, column=c0 + 1)
        a.font = f(color=BLUE); a.fill = INPUT_FILL; a.border = BORDER
        b.font = f(color=BLUE); b.fill = INPUT_FILL; b.border = BORDER; b.number_format = fmt2; b.alignment = CEN
        if i < len(rows):
            a.value = rows[i][0]; b.value = rows[i][1]
            a.alignment = Alignment(horizontal="left" if isinstance(rows[i][0], str) else "center")
        else:
            a.alignment = LEFT
    return top, top + total - 1

wb = Workbook()

# ---- Referenzen auf KONSTANTEN ----
K = "'Konstanten'!"
VA_RANGE = f"{K}$A$6:$B$13"
RW_RANGE = f"{K}$E$6:$E$13"; RW_LIST = RW_RANGE
ZONE_RANGE = f"{K}$H$6:$I$10"; ZONE_LIST = f"{K}$H$6:$H$10"   # Kürzel -> θF,max
PIPE_RANGE = f"{K}$A$18:$E$25"; PIPE_LIST = f"{K}$A$18:$A$25"

# =====================================================================
#  Blatt 1: GRUNDEINSTELLUNGEN  (Bezeichnungen ausgeschrieben)
# =====================================================================
g = wb.active
g.title = "Grundeinstellungen"
g.sheet_view.showGridLines = False
for col, w in (("A", 20), ("B", 15), ("C", 11), ("D", 7), ("E", 15), ("F", 7), ("G", 2),
               ("H", 20), ("I", 15), ("J", 11), ("K", 7), ("L", 15), ("M", 7)):
    g.column_dimensions[col].width = w
g.row_dimensions[1].height = 34
g["A1"].value = "Globale Grundeinstellungen – Fußbodenheizung"
g["A1"].font = Font(name=FONT, bold=True, color=NAVY, size=15); g["A1"].alignment = Alignment(horizontal="left", vertical="center")
add_logo_corner(g, 13, height=40)
for row, lab in ((2, "Projekt-Nr."), (3, "Projektname")):
    g.cell(row=row, column=1, value=lab).font = f(bold=True)
    g.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    v = g.cell(row=row, column=3, value=""); v.font = f(bold=True, color=BLUE); v.fill = INPUT_FILL; v.border = BORDER; v.alignment = LEFT
    g.merge_cells(start_row=row, start_column=3, end_row=row, end_column=4)
g.cell(row=2, column=12, value="Bearbeiter:").font = f(italic=True, color=GREY, size=9)
g["L2"].alignment = Alignment(horizontal="right")
bf = g.cell(row=2, column=13, value=""); bf.font = f(bold=True, color=BLUE); bf.fill = INPUT_FILL; bf.border = BORDER; bf.alignment = CEN

def section(row, text, cL):
    for col in range(cL, cL + 6): g.cell(row=row, column=col).fill = HDR_FILL
    g.cell(row=row, column=cL, value=text).font = f(bold=True, color=WHITE)
def param(row, label, value, unit, note, fmt, cL):
    g.cell(row=row, column=cL, value=label).font = f()
    g.merge_cells(start_row=row, start_column=cL, end_row=row, end_column=cL + 1)
    v = g.cell(row=row, column=cL + 2, value=value); v.font = f(bold=True, color=BLUE); v.alignment = CEN; v.border = BORDER; v.fill = INPUT_FILL
    if fmt: v.number_format = fmt
    g.cell(row=row, column=cL + 3, value=unit).font = f(color=GREY)
    g.cell(row=row, column=cL + 4, value=note).font = f(italic=True, color=GREY, size=9)
    g.merge_cells(start_row=row, start_column=cL + 4, end_row=row, end_column=cL + 5)
def calcrow(row, label, formula, unit, fmt, cL, note=""):
    g.cell(row=row, column=cL, value=label).font = f(); g.merge_cells(start_row=row, start_column=cL, end_row=row, end_column=cL + 1)
    v = g.cell(row=row, column=cL + 2, value=formula); v.font = f(bold=True); v.alignment = CEN; v.border = BORDER; v.number_format = fmt
    g.cell(row=row, column=cL + 3, value=unit).font = f(color=GREY)
    if note:
        g.cell(row=row, column=cL + 4, value=note).font = f(italic=True, color=GREY, size=9); g.merge_cells(start_row=row, start_column=cL + 4, end_row=row, end_column=cL + 5)

# ---- linke Spalte (Werte in C) ----
section(5, "Heizmedium / Temperaturen", 1)
param(6, "Vorlauftemperatur  θV", 35, "°C", "Auslegungs-Vorlauftemperatur", '0.0" °C"', 1)
param(7, "Rücklauftemperatur  θR", 28, "°C", "Auslegungs-Rücklauftemperatur", '0.0" °C"', 1)
calcrow(8, "Spreizung  θV − θR", "=C6-C7", "K", '0.0" K"', 1, "berechnet")
section(10, "Grenzwerte", 1)
param(11, "Maximale zulässige Kreislänge", 120, "m", "Obergrenze je Heizkreis", '0" m"', 1)
param(12, "Maximaler Druckverlust je Kreis", 15000, "Pa", "Warnschwelle bei Überschreitung", '#,##0" Pa"', 1)
section(14, "Stoffwerte Heizmedium (Wasser ~30 °C)", 1)
param(15, "Dichte  ρ", 995.7, "kg/m³", "Dichte des Mediums", '0.0', 1)
param(16, "spezifische Wärmekapazität  c_p", 4180, "J/(kg·K)", "spez. Wärmekapazität", '#,##0', 1)
param(17, "kinematische Viskosität  ν", 0.000000801, "m²/s", "für die Reynolds-Zahl", '0.00E+00', 1)
section(19, "Strömungsgrenzwerte", 1)
param(20, "Minimale Strömungsgeschwindigkeit  v_min", 0.10, "m/s", "untere empfohlene Grenze", '0.00" m/s"', 1)
param(21, "Maximale Strömungsgeschwindigkeit  v_max", 0.50, "m/s", "obere empfohlene Grenze", '0.00" m/s"', 1)
param(22, "Maximaler Volumenstrom je Kreis  V̇_max", 150, "l/h", "Warnschwelle je Heizkreis", '0" l/h"', 1)
section(23, "Druckverlust-Zuschläge", 1)
param(24, "Zuschlag Einzelwiderstände", 0.05, "-", "Rohr-Formstücke (0,05 = 5 %)", '0%', 1)
param(25, "Aufschlag Verteiler je Kreis", 500, "Pa", "fester Wert je Verteiler", '#,##0" Pa"', 1)

# ---- rechte Spalte (Werte in J) ----
section(5, "Wärmeübergang / Fußbodenaufbau", 8)
param(6, "Wärmeübergangskoeffizient Oberfläche  α", 10.8, "W/m²K", "Boden → Raum (EN 1264)", '0.0', 8)
param(7, "Estrichüberdeckung über Rohr  s_ü", 0.045, "m", "Estrich über den Rohren", '0.000', 8)
param(8, "Wärmeleitfähigkeit Estrich  λ_E", 1.2, "W/mK", "Zementestrich typ. 1,2", '0.00', 8)
section(10, "Wärmeabgabe nach unten (Dämmung, global)", 8)
param(11, "Wärmedurchlasswiderstand nach unten  R_u", 2.0, "m²K/W", "Dämmung + Aufbau unter den Rohren", '0.00', 8)
param(12, "Temperatur unter der Decke  θ_u", 10, "°C", "Raum / Erdreich unter der FBH", '0.0" °C"', 8)
calcrow(13, "Wärmestrom nach unten  q_u", "=((C6+C7)/2-J12)/J11", "W/m²", '0.0" W/m²"', 8, "Mittel θm = (θV+θR)/2")
section(15, "Rohrsystem-Auswahl (Werte aus Konstanten)", 8)
g.cell(row=16, column=8, value="Gewähltes Rohrsystem").font = f(); g.merge_cells("H16:I16")
v = g.cell(row=16, column=10, value="PE-Xa 17x2"); v.font = f(bold=True, color=BLUE); v.fill = INPUT_FILL; v.alignment = CEN; v.border = BORDER
g.merge_cells("J16:K16")
g.cell(row=16, column=12, value="Dropdown").font = f(italic=True, color=GREY, size=9); g.merge_cells("L16:M16")
calcrow(17, "→ Außendurchmesser  da", f"=VLOOKUP($J$16,{PIPE_RANGE},2,FALSE)", "mm", '0.0" mm"', 8, "aus Rohrbibliothek")
calcrow(18, "→ Wandstärke  s", f"=VLOOKUP($J$16,{PIPE_RANGE},3,FALSE)", "mm", '0.0" mm"', 8, "aus Rohrbibliothek")
calcrow(19, "→ Innendurchmesser  di", f"=VLOOKUP($J$16,{PIPE_RANGE},4,FALSE)", "mm", '0.0" mm"', 8, "di = da − 2·s")
calcrow(20, "→ Rauheit  k", f"=VLOOKUP($J$16,{PIPE_RANGE},5,FALSE)", "mm", '0.000" mm"', 8, "aus Rohrbibliothek")
calcrow(21, "→ Innendurchmesser  di", "=J19/1000", "m", '0.0000" m"', 8, "für die Berechnung")
calcrow(22, "→ Innenquerschnitt  A_i", "=PI()/4*J21^2", "m²", '0.000000" m²"', 8, "für die Berechnung")
dv_pipe = DataValidation(type="list", formula1=f"={PIPE_LIST}", allow_blank=False)
g.add_data_validation(dv_pipe); dv_pipe.add(g["J16"])
setup_print(g, "A1:M25")

# ---- Globale Referenzen ----
G = "'Grundeinstellungen'!"
gV, gR = f"{G}$C$6", f"{G}$C$7"
gMax, gWarn = f"{G}$C$11", f"{G}$C$12"
grho, gcp, gnu = f"{G}$C$15", f"{G}$C$16", f"{G}$C$17"
gvmin, gvmax = f"{G}$C$20", f"{G}$C$21"
gVdot = f"{G}$C$22"
gzus, gVarm = f"{G}$C$24", f"{G}$C$25"
galp, gsu, glam = f"{G}$J$6", f"{G}$J$7", f"{G}$J$8"
gqdown = f"{G}$J$13"
gdim, gAi = f"{G}$J$21", f"{G}$J$22"
gPNr, gPName, gBearb = f"{G}$C$2", f"{G}$C$3", f"{G}$M$2"
gfkorr = "'Verifikation'!$C$5"

# =====================================================================
#  Blatt 2: AUSLEGUNG
# =====================================================================
rl = wb.create_sheet("Auslegung")
rl.sheet_view.showGridLines = False
# Überschriften vom Nutzer optimiert (Wortrennungen beibehalten); Formelzeichen ergänzt.
columns = [
    ("Heizkreisverteiler", "HKV", "", 16, "text", BLUE),               # A
    ("Raum-Nr.", "", "", 10, "text", BLUE),                             # B
    ("Raumbezeichnung", "", "", 18, "text", BLUE),                     # C
    ("Raum-fläche", "A_R", "[m²]", 10, '0.0" m²"', BLUE),              # D
    ("Raumtemperatur", "θi", "[°C]", 11, '0.0" °C"', BLUE),            # E
    ("Heizlast", "Q", "[W]", 10, '#,##0" W"', BLUE),                   # F
    ("spez. Heizlast", "q_HL", "[W/m²]", 11, '0.0" W/m²"', BLACK),     # G  (=Q/Raumfläche)
    ("aktivierbare Fläche", "A_F", "[m²]", 11, '0.0" m²"', BLUE),      # H
    ("R-Wert Bodenbelag", "R_λ,B", "[m²·K/W]", 12, '0.000', BLUE),     # I
    ("Verlege–abstand", "VA", "[mm]", 11, '0" mm"', BLUE),             # J
    ("Anz.\nHK", "n", "[-]", 9, '0', BLUE),                            # K
    ("Zuleitungslänge", "L_zu", "[m]", 11, '0.0" m"', BLUE),           # L
    ("Zone", "", "", 7, "text", BLUE),                                 # M (Kürzel)
    ("log. Übertemperatur", "ΔθH", "[K]", 12, '0.00" K"', BLACK),      # N
    ("Wärmedurchgang", "K_H", "[W/m²K]", 11, '0.000', BLACK),          # O
    ("Verlegeabstand-Faktor", "η", "[-]", 11, '0.000', BLACK),         # P
    ("spez. Heizleistung", "q", "[W/m²]", 12, '0.0" W/m²"', BLACK),    # Q
    ("Oberflächentemperatur", "θF", "[°C]", 11, '0.0" °C"', BLACK),    # R
    ("max. Oberflächentemp.", "θF,max", "[°C]", 11, '0.0" °C"', BLACK),# S
    ("Leistung FBH nach oben", "Q_o", "[W]", 12, '#,##0" W"', BLACK),  # T
    ("Deckung absolut", "ΔQ", "[W]", 12, '#,##0" W"', BLACK),          # U
    ("Deckung", "", "[%]", 10, '0.0%', BLACK),                         # V
    ("Verlust nach unten", "Q_u", "[W]", 11, '#,##0" W"', BLACK),      # W
    ("maßgebl. Leistung", "Q_m", "[W]", 12, '#,##0" W"', BLACK),       # X  (=min(T;Q)+W)
    ("Rohrlänge Fläche", "L_F", "[m]", 11, '0.0" m"', BLACK),          # Y
    ("Rohrlänge Zuleitung", "L_Z", "[m]", 11, '0.0" m"', BLACK),       # Z
    ("Rohrlänge gesamt", "L_ges", "[m]", 11, '0.0" m"', BLACK),        # AA
    ("Rohrlänge je Kreis", "L_K", "[m]", 11, '0.0" m"', BLACK),        # AB
    ("Massenstrom je Kreis", "ṁ", "[kg/h]", 11, '0.0" kg/h"', BLACK),  # AC
    ("Volumen-strom je Kreis", "V̇", "[l/h]", 11, '0.0" l/h"', BLACK), # AD
    ("Geschwin-digkeit", "v", "[m/s]", 11, '0.000" m/s"', BLACK),      # AE
    ("Reynolds-Zahl", "Re", "[-]", 11, '#,##0', BLACK),                # AF
    ("Strömungsbereich", "", "[-]", 12, "text", BLACK),                # AG
    ("Rohrreibungszahl", "λ", "[-]", 11, '0.0000', BLACK),             # AH
    ("Δp Reibung", "Δp_R", "[Pa]", 11, '#,##0" Pa"', BLACK),           # AI
    ("spez. Druckverlust", "R", "[Pa/m]", 12, '#,##0" Pa/m"', BLACK),  # AJ
    ("Druckverlust", "Δp", "[Pa]", 12, '#,##0" Pa"', BLACK),           # AK
]  # A..AK = 37
NCOL = len(columns)
LASTCOL = get_column_letter(NCOL)
INPUT_COLS = [j for j, c in enumerate(columns, start=1) if c[5] == BLUE]
CALC_COLS = [j for j, c in enumerate(columns, start=1) if c[5] == BLACK]
INPUT_STYLED = {7}   # spez. Heizlast: berechnet, aber optisch wie die Eingabespalten
for j, (name, sym, unit, width, fmt, color) in enumerate(columns, start=1):
    rl.column_dimensions[get_column_letter(j)].width = width
    is_inp = (color == BLUE) or (j in INPUT_STYLED)
    fill = SUB_FILL if is_inp else HDR_FILL
    tcol = NAVY if is_inp else WHITE
    for row, txt, sz in ((NAME_ROW, name, 10), (SYM_ROW, sym, 10), (UNIT_ROW, unit, 9)):
        c = rl.cell(row=row, column=j, value=txt)
        c.fill = fill; c.font = f(bold=True, color=tcol, size=sz)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = BORDER
rl.row_dimensions[NAME_ROW].height = 34
rl.row_dimensions[SYM_ROW].height = 14
rl.row_dimensions[UNIT_ROW].height = 14

def F(r):
    return {
        7:  f'=IF($B{r}="","",IFERROR($F{r}/$D{r},""))',
        14: f'=IF($B{r}="","",IFERROR(({gV}-{gR})/LN(({gV}-$E{r})/({gR}-$E{r})),""))',
        15: f'=IF($B{r}="","",IFERROR(1/(1/{galp}+$I{r}+{gsu}/{glam}),""))',
        16: f'=IF($B{r}="","",IFERROR(VLOOKUP($J{r},{VA_RANGE},2,TRUE),""))',
        17: f'=IF(OR($B{r}="",$N{r}="",$O{r}="",$P{r}=""),"",{gfkorr}*$O{r}*$P{r}*$N{r})',
        18: f'=IF(OR($B{r}="",$Q{r}=""),"",$E{r}+$Q{r}/{galp})',
        19: f'=IF($B{r}="","",IFERROR(VLOOKUP($M{r},{ZONE_RANGE},2,FALSE),29))',
        20: f'=IF(OR($B{r}="",$Q{r}=""),"",$Q{r}*$H{r})',
        21: f'=IF(OR($B{r}="",$T{r}=""),"",$T{r}-$F{r})',
        22: f'=IF(OR($B{r}="",$T{r}=""),"",IFERROR($T{r}/$F{r},""))',
        23: f'=IF(OR($B{r}="",$H{r}=""),"",{gqdown}*$H{r})',
        24: f'=IF(OR($B{r}="",$T{r}=""),"",MIN($T{r},$F{r})+$W{r})',
        25: f'=IF($B{r}="","",IFERROR($H{r}/($J{r}/1000),""))',
        26: f'=IF($B{r}="","",IFERROR(2*$L{r}*$K{r},""))',
        27: f'=IF(OR($B{r}="",$Y{r}="",$Z{r}=""),"",$Y{r}+$Z{r})',
        28: f'=IF(OR($B{r}="",$AA{r}=""),"",IFERROR($AA{r}/$K{r},""))',
        29: f'=IF(OR($B{r}="",$X{r}=""),"",IFERROR(($X{r}/$K{r})/({gcp}*({gV}-{gR}))*3600,""))',
        30: f'=IF(OR($B{r}="",$AC{r}=""),"",$AC{r}/{grho}*1000)',
        31: f'=IF(OR($B{r}="",$AC{r}=""),"",IFERROR($AC{r}/3600/{grho}/{gAi},""))',
        32: f'=IF(OR($B{r}="",$AE{r}=""),"",IFERROR($AE{r}*{gdim}/{gnu},""))',
        33: f'=IF(OR($B{r}="",$AF{r}=""),"",IF($AF{r}<2300,"laminar",IF($AF{r}<4000,"Übergangsbereich","turbulent")))',
        34: f'=IF(OR($B{r}="",$AF{r}=""),"",IF($AF{r}<2300,64/$AF{r},0.316/$AF{r}^0.25))',
        35: f'=IF(OR($B{r}="",$AH{r}="",$AB{r}=""),"",IFERROR($AH{r}*($AB{r}/{gdim})*({grho}/2)*$AE{r}^2,""))',
        36: f'=IF(OR($B{r}="",$AI{r}="",$AB{r}=""),"",IFERROR($AI{r}/$AB{r},""))',
        37: f'=IF(OR($B{r}="",$AI{r}=""),"",$AI{r}*(1+{gzus})+{gVarm})',
    }

examples = [   # Heizlast so gewählt, dass spez. Heizlast (Q/Raumfläche) <= 50 W/m², bunte Mischung
    ["HKV EG", "001", "Foyer/Empfang", 40, 20, 1800, 36, 0.00, 100, 3, 6,  "RZ"],   # 45 W/m²
    ["HKV EG", "002", "Großraumbüro",  60, 20, 2400, 55, 0.10, 150, 4, 8,  "AZ"],   # 40 W/m²
    ["HKV EG", "003", "Büro 1",        18, 20,  630, 16, 0.10, 150, 1, 5,  "AZ"],   # 35 W/m²
    ["HKV EG", "004", "Besprechung",   30, 20, 1500, 28, 0.07, 100, 2, 7,  "AZ"],   # 50 W/m²
    ["HKV EG", "005", "Flur",          25, 20,  750, 22, 0.05, 200, 1, 3,  "AZ"],   # 30 W/m²
    ["HKV OG", "101", "WC / Sanitär",  12, 24,  600,  9, 0.01, 100, 1, 10, "BAD"],  # 50 W/m²
    ["HKV OG", "102", "Teeküche",      10, 20,  450,  8, 0.01, 100, 1, 9,  "AZ"],   # 45 W/m²
    ["HKV OG", "103", "Archiv",        20, 18,  500, 18, 0.05, 200, 1, 12, "AZ"],   # 25 W/m²
    ["HKV OG", "104", "Technikraum",   15, 15,  300, 12, 0.00, 250, 1, 14, "AZ"],   # 20 W/m²
]
for idx, r in enumerate(range(R0, R1 + 1)):
    fr = F(r)
    for ji, j in enumerate(INPUT_COLS):
        c = rl.cell(row=r, column=j)
        if idx < len(examples): c.value = examples[idx][ji]
        c.font = f(color=BLUE); c.fill = INPUT_FILL; c.border = BORDER
        fmt = columns[j - 1][4]
        if fmt != "text": c.number_format = fmt
        c.alignment = Alignment(horizontal="left" if j in (1, 2, 3) else "center")
    for j in CALC_COLS:
        c = rl.cell(row=r, column=j, value=fr[j])
        c.font = f(color=BLACK); c.border = BORDER
        if j in INPUT_STYLED: c.fill = INPUT_FILL
        fmt = columns[j - 1][4]
        if fmt != "text": c.number_format = fmt
        c.alignment = CEN

rl.freeze_panes = "C7"
dv_zone = DataValidation(type="list", formula1=f"={ZONE_LIST}", allow_blank=True, showErrorMessage=False)
rl.add_data_validation(dv_zone); dv_zone.add(f"M{R0}:M{R1}")
dv_rw = DataValidation(type="list", formula1=f"={RW_LIST}", allow_blank=True, showErrorMessage=False)
rl.add_data_validation(dv_rw); dv_rw.add(f"I{R0}:I{R1}")

def cf_pair(col, ok_formula, bad_formula):
    rng = f"{col}{R0}:{col}{R1}"
    rl.conditional_formatting.add(rng, FormulaRule(formula=[bad_formula], fill=RED_FILL))
    rl.conditional_formatting.add(rng, FormulaRule(formula=[ok_formula], fill=GREEN_FILL))
cf_pair("R", f'AND($R{R0}<>"",$R{R0}<=$S{R0})', f'AND($R{R0}<>"",$R{R0}>$S{R0})')
cf_pair("V", f'AND($V{R0}<>"",$V{R0}>=1)', f'AND($V{R0}<>"",$V{R0}<1)')
cf_pair("AB", f'AND($AB{R0}<>"",$AB{R0}<={gMax})', f'AND($AB{R0}<>"",$AB{R0}>{gMax})')
cf_pair("AD", f'AND($AD{R0}<>"",$AD{R0}<={gVdot})', f'AND($AD{R0}<>"",$AD{R0}>{gVdot})')
cf_pair("AK", f'AND($AK{R0}<>"",$AK{R0}<={gWarn})', f'AND($AK{R0}<>"",$AK{R0}>{gWarn})')
cf_pair("AE", f'AND($AE{R0}<>"",$AE{R0}>={gvmin},$AE{R0}<={gvmax})',
        f'AND($AE{R0}<>"",OR($AE{R0}<{gvmin},$AE{R0}>{gvmax}))')
rl.conditional_formatting.add(f"H{R0}:H{R1}", FormulaRule(
    formula=[f'AND($H{R0}<>"",$D{R0}<>"",$H{R0}>$D{R0})'], fill=RED_FILL))
# schlanker Standard: Zwischen-/Sekundärspalten ausblenden (jederzeit einblendbar)
for col in ["N", "O", "P", "S", "T", "U", "W", "X", "Y", "Z", "AA", "AC", "AF", "AG", "AH", "AI", "AJ"]:
    rl.column_dimensions[col].hidden = True
disp_header(rl, "Auslegung – Fußbodenheizung", NCOL, "AJ2", "AJ2:AK2")
setup_print(rl, f"A1:{LASTCOL}{R1}", titles="1:6", orientation="landscape", paper=8)

# =====================================================================
#  Blatt 3: KONTROLLE
# =====================================================================
ov = wb.create_sheet("Kontrolle")
ov.sheet_view.showGridLines = False
for col, w in (("A", 44), ("B", 16), ("C", 11), ("D", 40)):
    ov.column_dimensions[col].width = w
AUS = "'Auslegung'!"
def ovrow(r, label, formula, unit, fmt='#,##0', note=""):
    ov.cell(row=r, column=1, value=label).font = f()
    c = ov.cell(row=r, column=2, value=formula); c.font = f(bold=True); c.number_format = fmt
    c.alignment = CEN; c.border = BORDER
    ov.cell(row=r, column=3, value=unit).font = f(color=GREY)
    if note: ov.cell(row=r, column=4, value=note).font = f(italic=True, color=GREY, size=9)
c = ov.cell(row=5, column=1, value="Summen / Kennzahlen"); c.font = f(bold=True, color=WHITE); c.fill = HDR_FILL
for col in (2, 3, 4): ov.cell(row=5, column=col).fill = HDR_FILL
ovrow(6, "Anzahl Räume", f"=COUNTA({AUS}B{R0}:B{R1})", "-", '0')
ovrow(7, "Gesamte Heizlast", f"=SUM({AUS}F{R0}:F{R1})", "W")
ovrow(8, "Gesamte FBH-Leistung (nach oben)", f"=SUM({AUS}T{R0}:T{R1})", "W")
ovrow(9, "Verlustleistung nach unten", f"=SUM({AUS}W{R0}:W{R1})", "W")
ovrow(10, "Maßgebliche Gesamtleistung (Hydraulik)", f"=SUM({AUS}X{R0}:X{R1})", "W")
ovrow(11, "Deckungsgrad gesamt", "=IFERROR(B8/B7,0)", "%", '0.0%')
ovrow(12, "Gesamte Rohrlänge", f"=SUM({AUS}AA{R0}:AA{R1})", "m")
ovrow(13, "Anzahl Heizkreise gesamt", f"=SUM({AUS}K{R0}:K{R1})", "-", '0')
ovrow(14, "Gesamtmassenstrom (Pumpe)", f"=SUM({AUS}AC{R0}:AC{R1})", "kg/h", '#,##0')
ovrow(15, "Gesamtvolumenstrom (Pumpe)", f"=SUM({AUS}AD{R0}:AD{R1})", "l/h", '#,##0')
ovrow(16, "Max. Druckverlust eines Kreises", f"=IFERROR(MAX({AUS}AK{R0}:AK{R1}),0)", "Pa")
ovrow(17, "Max. Strömungsgeschwindigkeit", f"=IFERROR(MAX({AUS}AE{R0}:AE{R1}),0)", "m/s", '0.000')
c = ov.cell(row=19, column=1, value="Warnungen"); c.font = f(bold=True, color=WHITE); c.fill = HDR_FILL
for col in (2, 3, 4): ov.cell(row=19, column=col).fill = HDR_FILL
ovrow(20, "Räume unterdeckt", f'=COUNTIF({AUS}V{R0}:V{R1},"<1")', "-", '0', "Heizlast nicht gedeckt")
ovrow(21, "Kreise zu lang", f'=COUNTIF({AUS}AB{R0}:AB{R1},">"&{gMax})', "-", '0', "> max. Kreislänge")
ovrow(22, "Druckverlust über Warnschwelle", f'=COUNTIF({AUS}AK{R0}:AK{R1},">"&{gWarn})', "-", '0')
ovrow(23, "Oberflächentemperatur zu hoch", f'=SUMPRODUCT(({AUS}R{R0}:R{R1}<>"")*({AUS}R{R0}:R{R1}>{AUS}S{R0}:S{R1}))', "-", '0')
ovrow(24, "Geschwindigkeit außerhalb v_min–v_max", f'=SUMPRODUCT(({AUS}AE{R0}:AE{R1}<>"")*(({AUS}AE{R0}:AE{R1}<{gvmin})+({AUS}AE{R0}:AE{R1}>{gvmax})))', "-", '0', "Kreise")
ovrow(25, "Volumenstrom je Kreis zu hoch", f'=COUNTIF({AUS}AD{R0}:AD{R1},">"&{gVdot})', "-", '0', "> max. Volumenstrom")
for r in (20, 21, 22, 23, 24, 25):
    ov.conditional_formatting.add(f"B{r}", FormulaRule(formula=[f"$B${r}>0"], fill=RED_FILL))
    ov.conditional_formatting.add(f"B{r}", FormulaRule(formula=[f"$B${r}=0"], fill=GREEN_FILL))
disp_header(ov, "Kontrolle / Auswertung", 4, "D2")
setup_print(ov, "A1:D25")

# =====================================================================
#  Blatt 4: HEIZKREISVERTEILER  (+ Massenstrom)
# =====================================================================
hv = wb.create_sheet("HKV")
hv.sheet_view.showGridLines = False
for col, w in (("A", 22), ("B", 12), ("C", 12), ("D", 20), ("E", 18), ("F", 18), ("G", 20)):
    hv.column_dimensions[col].width = w
for j, h in enumerate(["Heizkreisverteiler (HKV)", "Anzahl Kreise", "Spreizung [K]", "maßgebl. Leistung [W]",
                       "Massenstrom [kg/h]", "Volumenstrom [l/h]", "max. Druckverlust [Pa]"], start=1):
    c = hv.cell(row=4, column=j, value=h)
    c.font = f(bold=True, color=WHITE); c.fill = HDR_FILL
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = BORDER
hv.row_dimensions[4].height = 30
Aaus = f"{AUS}$A${R0}:$A${R1}"
N_HKV = 50
for i in range(N_HKV):
    r = 5 + i
    af = (f'=IFERROR(INDEX({Aaus},MATCH(0,COUNTIF($A$4:$A{r-1},{Aaus})+IF({Aaus}="",1,0),0)),"")')
    cell = hv.cell(row=r, column=1); cell.value = ArrayFormula(f"A{r}", af); cell.font = f(); cell.alignment = LEFT; cell.border = BORDER
    b = hv.cell(row=r, column=2, value=f'=IF($A{r}="","",SUMIF({Aaus},$A{r},{AUS}$K${R0}:$K${R1}))')
    sp = hv.cell(row=r, column=3, value=f"=IF($A{r}=\"\",\"\",{gV}-{gR})")   # globale Spreizung (überall gleich)
    c = hv.cell(row=r, column=4, value=f'=IF($A{r}="","",SUMIF({Aaus},$A{r},{AUS}$X${R0}:$X${R1}))')
    d = hv.cell(row=r, column=5, value=f'=IF($A{r}="","",SUMIF({Aaus},$A{r},{AUS}$AC${R0}:$AC${R1}))')
    e = hv.cell(row=r, column=6, value=f'=IF($A{r}="","",SUMIF({Aaus},$A{r},{AUS}$AD${R0}:$AD${R1}))')
    p = hv.cell(row=r, column=7)
    p.value = ArrayFormula(f"G{r}", f'=IF($A{r}="","",MAX(IF({Aaus}=$A{r},{AUS}$AK${R0}:$AK${R1})))')
    for cc, fmt in ((b, '0'), (sp, '0.0" K"'), (c, '#,##0" W"'), (d, '#,##0" kg/h"'), (e, '#,##0" l/h"'), (p, '#,##0" Pa"')):
        cc.font = f(); cc.alignment = CEN; cc.number_format = fmt; cc.border = BORDER
hv.cell(row=5 + N_HKV + 1, column=1,
        value="Baut sich automatisch aus der Auslegung auf (Array-Formeln, Excel & LibreOffice).").font = f(italic=True, color=GREY, size=9)
disp_header(hv, "Heizkreisverteiler – Übersicht je HKV", 7, "F2", "F2:G2")
setup_print(hv, f"A1:G{5 + N_HKV + 1}")

# =====================================================================
#  Blatt 5: VERIFIKATION
# =====================================================================
vf = wb.create_sheet("Verifikation")
vf.sheet_view.showGridLines = False
for col, w in (("A", 20), ("B", 13), ("C", 14), ("D", 14), ("E", 11), ("F", 10), ("G", 11), ("H", 8), ("I", 2)):
    vf.column_dimensions[col].width = w
vf.cell(row=5, column=1, value="Korrekturfaktor Heizleistung  f =").font = f(bold=True, color=NAVY, size=11)
vf.merge_cells("A5:B5"); vf["A5"].fill = ACCENT_FILL; vf["B5"].fill = ACCENT_FILL
vf["A5"].alignment = Alignment(horizontal="right", vertical="center")
fk = vf.cell(row=5, column=3, value=1.00); fk.font = f(bold=True, color=BLUE, size=11); fk.fill = INPUT_FILL
fk.border = BORDER; fk.alignment = Alignment(horizontal="center", vertical="center"); fk.number_format = '0.000'
vf.cell(row=5, column=4, value="← an Herstellerdaten anpassen").font = f(italic=True, color=GREY, size=9)
vf.row_dimensions[5].height = 22
c = vf.cell(row=7, column=1, value="Vergleichsbedingungen (fix)")
c.font = f(bold=True, color=WHITE); c.fill = HDR_FILL
for col in (2, 3): vf.cell(row=7, column=col).fill = HDR_FILL
def vparam(row, label, value, unit, fmt):
    vf.cell(row=row, column=1, value=label).font = f()
    v = vf.cell(row=row, column=2, value=value); v.font = f(bold=True, color=BLUE); v.fill = INPUT_FILL; v.border = BORDER; v.alignment = CEN; v.number_format = fmt
    vf.cell(row=row, column=3, value=unit).font = f(color=GREY)
vparam(8, "Vorlauftemperatur  θV", 35, "°C", '0.0" °C"')
vparam(9, "Rücklauftemperatur  θR", 28, "°C", '0.0" °C"')
vparam(10, "Raumtemperatur  θi", 20, "°C", '0.0" °C"')
vparam(11, "R-Wert Bodenbelag", 0.10, "m²K/W", '0.000')
vf.cell(row=12, column=1, value="q berechnet = f · η(VA) · K_H · ΔθH").font = f(italic=True, color=GREY, size=9)
vf.merge_cells("A12:D12")
vhdr = ["Verlegeabstand [mm]", "q berechnet [W/m²]", "Hersteller q [W/m²]", "Abweichung [W/m²]",
        "Abweichung [%]", "ΔθH [K]", "K_H [W/m²K]", "η [-]"]
VT = 14
for j, h in enumerate(vhdr, start=1):
    c = vf.cell(row=VT, column=j, value=h)
    c.fill = HDR_FILL if j != 3 else SUB_FILL
    c.font = f(bold=True, color=(WHITE if j != 3 else NAVY))
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = BORDER
vf.row_dimensions[VT].height = 42
V0 = VT + 1
va_values = [50, 75, 100, 125, 150, 200, 250, 300]
hq_example = [54, 52, 51, 49, 47, 43, 40, 37]
V1 = V0 + len(va_values) - 1
for i, va in enumerate(va_values):
    r = V0 + i
    a = vf.cell(row=r, column=1, value=va); a.font = f(color=BLUE); a.fill = INPUT_FILL; a.alignment = CEN; a.number_format = '0" mm"'; a.border = BORDER
    cq = vf.cell(row=r, column=3, value=hq_example[i]); cq.font = f(color=BLUE); cq.fill = INPUT_FILL; cq.alignment = CEN; cq.number_format = '0.0" W/m²"'; cq.border = BORDER
    forms = {
        2: (f'=IF(OR($F{r}="",$G{r}="",$H{r}=""),"",$C$5*$G{r}*$H{r}*$F{r})', '0.0" W/m²"'),
        4: (f'=IF(OR($B{r}="",$C{r}=""),"",$C{r}-$B{r})', '0.0" W/m²"'),
        5: (f'=IF(OR($B{r}="",$C{r}=""),"",IFERROR(($C{r}-$B{r})/$B{r},""))', '0.0%'),
        6: ('=IFERROR(($B$8-$B$9)/LN(($B$8-$B$10)/($B$9-$B$10)),"")', '0.00" K"'),
        7: (f'=IFERROR(1/(1/{galp}+$B$11+{gsu}/{glam}),"")', '0.000'),
        8: (f'=IFERROR(VLOOKUP($A{r},{VA_RANGE},2,TRUE),"")', '0.000'),
    }
    for col, (formula, fmt) in forms.items():
        cc = vf.cell(row=r, column=col, value=formula); cc.font = f(); cc.alignment = CEN; cc.number_format = fmt; cc.border = BORDER
vf.conditional_formatting.add(f"E{V0}:E{V1}", FormulaRule(formula=[f'AND($E{V0}<>"",ABS($E{V0})>0.1)'], fill=RED_FILL))
vf.conditional_formatting.add(f"E{V0}:E{V1}", FormulaRule(formula=[f'AND($E{V0}<>"",ABS($E{V0})<=0.1)'], fill=GREEN_FILL))
chart = LineChart(); chart.title = "Heizleistung über Verlegeabstand"
chart.y_axis.title = "q [W/m²]"; chart.x_axis.title = "Verlegeabstand [mm]"
chart.height = 10; chart.width = 14; chart.style = 2
chart.add_data(Reference(vf, min_col=2, min_row=VT, max_row=V1), titles_from_data=True)
chart.add_data(Reference(vf, min_col=3, min_row=VT, max_row=V1), titles_from_data=True)
chart.set_categories(Reference(vf, min_col=1, min_row=V0, max_row=V1))
vf.add_chart(chart, "J5")   # rechts neben den Tabellen
nr = V1 + 2
for i, (txt, kind) in enumerate([
    ("Hinweise", "h"),
    ("• Vergleich EINES Berechnungsmodells mit EINEM Herstellermodell über den Verlegeabstand.", ""),
    ("• Hersteller-q je Verlegeabstand aus dem Datenblatt in Spalte C eintragen.", ""),
    ("• Korrekturfaktor f oben so einstellen, dass die Linien übereinanderliegen (Abweichung grün).", "i")]):
    c = vf.cell(row=nr + i, column=1, value=txt)
    if kind == "h": c.font = f(bold=True, color=NAVY, size=11)
    elif kind == "i": c.font = f(italic=True, color=GREY, size=9)
    else: c.font = f()
disp_header(vf, "Verifikation – Berechnung vs. Hersteller (über Verlegeabstand)", 16, project=True)
setup_print(vf, f"A1:Q{max(nr + 2, 24)}")

# =====================================================================
#  Blatt 6: METHODIK
# =====================================================================
m = wb.create_sheet("Methodik")
m.sheet_view.showGridLines = False
for col, w in (("A", 3), ("B", 46), ("C", 3), ("D", 48), ("E", 3), ("F", 48)):
    m.column_dimensions[col].width = w

def methcol(col, start, items):
    rr = start
    for text, kind in items:
        c = m.cell(row=rr, column=col, value=text)
        if kind == "h": c.font = Font(name=FONT, bold=True, color=NAVY, size=11)
        elif kind == "fb": c.font = Font(name=FONT, bold=True, color=BLACK, size=10)
        elif kind == "i": c.font = Font(name=FONT, italic=True, color=GREY, size=9)
        else: c.font = Font(name=FONT, size=10)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        rr += 1
    return rr

m["B4"].value = "Formelzeichen / Variablen"; m["B4"].font = Font(name=FONT, bold=True, color=NAVY, size=11)
m["D4"].value = "Berechnungsschritte"; m["D4"].font = Font(name=FONT, bold=True, color=NAVY, size=11)
m.merge_cells("D4:F4")
vars_ = [
    ("θV, θR — Vor-/Rücklauftemperatur [°C]", ""), ("θi — Raumtemperatur [°C]", ""),
    ("θ_u — Temperatur unter der Decke [°C]", ""), ("ΔθH — log. Übertemperatur [K]", ""),
    ("q — spez. Heizleistung nach oben [W/m²]", ""), ("α — Wärmeübergang Oberfläche [W/m²K]", ""),
    ("R_Belag — Widerstand Bodenbelag [m²K/W]", ""), ("s_ü — Estrichüberdeckung [m]", ""),
    ("λ_E — Wärmeleitfähigkeit Estrich [W/mK]", ""), ("K_H — Wärmedurchgang nach oben [W/m²K]", ""),
    ("η_VA — Verlegeabstandsfaktor [-]", ""), ("f — Korrekturfaktor Heizleistung [-]", ""),
    ("R_u — Widerstand nach unten [m²K/W]", ""), ("q_u — Wärmestrom nach unten [W/m²]", ""),
    ("θF, θF,max — Oberflächentemp. / Grenze [°C]", ""), ("di, A_i — Rohr-Innen-Ø / -Querschnitt", ""),
    ("ρ, c_p, ν — Dichte / Wärmekap. / Viskosität", ""), ("ṁ, V̇ — Massen-/Volumenstrom [kg/h]/[l/h]", ""),
    ("v, Re, λ — Geschw. / Reynolds / Reibung", ""), ("L_Kreis — Rohrlänge je Kreis [m]", ""),
    ("Δp — Druckverlust [Pa]", ""),
]
steps_l = [
    ("1)  ΔθH = (θV−θR) / ln((θV−θi)/(θR−θi))", "fb"),
    ("Mittlere treibende Temperaturdifferenz Heizwasser ↔ Raum (log. Mittel).", "i"), ("", ""),
    ("2)  K_H = 1 / (1/α + R_Belag + s_ü/λ_E)", "fb"),
    ("Wärmedurchgang des Aufbaus nach oben.", "i"), ("", ""),
    ("3)  q = f · η_VA · K_H · ΔθH", "fb"),
    ("Spez. Heizleistung nach oben; f kalibriert an Herstellerdaten.", "i"), ("", ""),
    ("4)  θF = θi + q/α", "fb"),
    ("Bodenoberflächentemperatur; Prüfung gegen θF,max.", "i"), ("", ""),
    ("5)  Leistung oben = q · Fläche ; Deckung % = Leistung / Heizlast", "fb"),
    ("Prüft, ob die FBH die Heizlast deckt (≥ 100 %).", "i"),
]
steps_r = [
    ("6)  q_u = ((θV+θR)/2 − θ_u)/R_u ; Verlust = q_u · Fläche", "fb"),
    ("Wärmeabgabe nach unten durch die Dämmung.", "i"), ("", ""),
    ("7)  maßgebl. Leistung = MIN(Leistung ; Heizlast) + Verlust", "fb"),
    ("Tatsächlich transportierte Wärme (Basis für Massenstrom).", "i"), ("", ""),
    ("8)  L_Kreis = (Fläche/Verlegeabstand + 2·Zuleitung·Kreise)/Kreise", "fb"),
    ("Rohrlänge je Kreis; Prüfung gegen max. Kreislänge.", "i"), ("", ""),
    ("9)  ṁ = (maßgebl. Leistung/Kreise)/(c_p·(θV−θR))", "fb"),
    ("v = V̇/A_i ; Re = v·di/ν ; Re < 2300 = laminar.", "i"), ("", ""),
    ("10) Δp = λ·(L/di)·(ρ/2)·v²·(1+Zuschlag) + Verteiler", "fb"),
    ("Druckverlust je Kreis; Kontrolle gegen Warnschwelle.", "i"), ("", ""),
    ("Vereinfachungen", "h"),
    ("• K_H vereinfacht + η-Faktor statt EN-1264; Rohrwand nicht enthalten (→ f).", ""),
    ("• Wärmeabgabe nach unten global; pauschale Δp-Zuschläge.", ""),
]
end_b = methcol(2, 5, vars_)
methcol(4, 6, steps_l)
methcol(6, 6, steps_r)
disp_header(m, "Methodik & Annahmen", 6, project=False)
setup_print(m, f"A1:F{end_b}")

# =====================================================================
#  Blatt 7: KONSTANTEN
# =====================================================================
kt = wb.create_sheet("Konstanten")
kt.sheet_view.showGridLines = False
for col, w in (("A", 22), ("B", 12), ("C", 9), ("D", 20), ("E", 12), ("F", 2), ("G", 16), ("H", 8), ("I", 11)):
    kt.column_dimensions[col].width = w
kt.cell(row=3, column=1, value="Nachschlage-Tabellen (editierbar). Leerzeilen = Platz für weitere Einträge.").font = f(italic=True, color=GREY, size=9)
# Oben nebeneinander: Verlegeabstand-Faktor | R-Werte | Zonen
two_col_table(kt, 4, 1, "Verlegeabstand-Faktor", "Verlegeabstand [mm]", "Faktor η [-]",
    [(50, 1.05), (75, 1.02), (100, 1.00), (125, 0.97), (150, 0.93), (200, 0.85), (250, 0.78), (300, 0.72)], '0.000')
two_col_table(kt, 4, 4, "R-Werte Bodenbeläge", "Bodenbelag", "R [m²·K/W]",
    [("Fliesen / Naturstein", 0.010), ("Linoleum / PVC", 0.05), ("Parkett / Laminat", 0.07),
     ("Holzdielen", 0.10), ("Teppich", 0.15)], '0.000', n_empty=3)
# Zonen 3-spaltig (Zone | Kürzel | θF,max) – Kürzel wird in der Auslegung verwendet
kt.cell(row=4, column=7, value="Zonen / max. Oberflächentemp.").font = f(bold=True, color=NAVY)
kt.merge_cells("G4:I4")
mini_header(kt, 5, 7, "Zone"); mini_header(kt, 5, 8, "Kürzel"); mini_header(kt, 5, 9, "θF,max [°C]")
zdata = [("Aufenthaltszone", "AZ", 29), ("Randzone", "RZ", 35), ("Badezimmer", "BAD", 33)]
for i in range(5):   # 3 + 2 leer
    r = 6 + i
    for col in (7, 8, 9):
        cell = kt.cell(row=r, column=col); cell.font = f(color=BLUE); cell.fill = INPUT_FILL; cell.border = BORDER
        cell.alignment = Alignment(horizontal="left" if col <= 8 else "center")
        if col == 9: cell.number_format = '0" °C"'
        if i < len(zdata): cell.value = zdata[i][col - 7]
# Darunter volle Breite: Rohrbibliothek
kt.cell(row=16, column=1, value="Rohrbibliothek (di = da − 2·s)").font = f(bold=True, color=NAVY)
for j, h in enumerate(["Rohrsystem", "da [mm]", "s [mm]", "di [mm]", "k [mm]"], start=1):
    mini_header(kt, 17, j, h)
pipes = [("PE-Xa 17x2", 17, 2.0, 0.007), ("PE-Xa 16x2", 16, 2.0, 0.007),
         ("PE-Xa 20x2", 20, 2.0, 0.007), ("PE-RT 14x2", 14, 2.0, 0.007)]
colmap = {1: 0, 2: 1, 3: 2, 5: 3}
for i in range(8):
    r = 18 + i
    for col in range(1, 6):
        cell = kt.cell(row=r, column=col); cell.border = BORDER
        if col == 4:
            cell.value = f'=IF(B{r}="","",B{r}-2*C{r})'; cell.font = f(color=BLACK); cell.number_format = '0.0'; cell.alignment = CEN
        else:
            cell.font = f(color=BLUE); cell.fill = INPUT_FILL
            cell.alignment = Alignment(horizontal="left" if col == 1 else "center")
            if col in (2, 3): cell.number_format = '0.0'
            if col == 5: cell.number_format = '0.000'
            if i < len(pipes): cell.value = pipes[i][colmap[col]]
disp_header(kt, "Konstanten / Bibliotheken", 9, project=False)
setup_print(kt, "A1:I26")

# =====================================================================
#  Blatt 8: CHANGELOG
# =====================================================================
cl = wb.create_sheet("Changelog")
cl.sheet_view.showGridLines = False
for col, w in (("A", 12), ("B", 14), ("C", 115)):
    cl.column_dimensions[col].width = w
for j, h in enumerate(["Version", "Datum", "Änderungen"], start=1):
    c = cl.cell(row=5, column=j, value=h)
    c.font = f(bold=True, color=WHITE); c.fill = HDR_FILL
    c.alignment = Alignment(horizontal="left", vertical="center"); c.border = BORDER
r = 6
for ver, datum, changes in reversed(CHANGELOG):
    start = r
    for ch in changes:
        cl.cell(row=r, column=3, value="• " + ch).font = f()
        cl.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="top")
        for col in (1, 2, 3): cl.cell(row=r, column=col).border = BORDER
        r += 1
    a = cl.cell(row=start, column=1, value=ver); a.font = f(bold=True); a.alignment = Alignment(vertical="top")
    b = cl.cell(row=start, column=2, value=datum); b.font = f(); b.alignment = Alignment(vertical="top")
    if r - start > 1:
        cl.merge_cells(start_row=start, start_column=1, end_row=r - 1, end_column=1)
        cl.merge_cells(start_row=start, start_column=2, end_row=r - 1, end_column=2)
disp_header(cl, "Changelog – FBH-Auslegung", 3, project=False)
cl.cell(row=3, column=3, value=VERTXT).font = f(italic=True, color=GREY, size=9)
cl.cell(row=3, column=3).alignment = Alignment(horizontal="right")
setup_print(cl, f"A1:C{r-1}")

# =====================================================================
#  Blatt 0: DECKBLATT  (Übersichtsdiagramm aus Zellen – in Excel editierbar)
# =====================================================================
db = wb.create_sheet("Deckblatt", 0)
db.sheet_view.showGridLines = False
for col, w in (("A", 14), ("B", 14), ("C", 14), ("D", 6), ("E", 14), ("F", 14),
               ("G", 14), ("H", 6), ("I", 14), ("J", 14), ("K", 14)):
    db.column_dimensions[col].width = w
RED_S = PatternFill("solid", fgColor="E2001A")
GREEN_S = PatternFill("solid", fgColor="2E7D32")
BODY_S = PatternFill("solid", fgColor="F7F7F7")
AMB_H = PatternFill("solid", fgColor="FFE699")
AMB_B = PatternFill("solid", fgColor="FBF3D9")
AMBER = "B88600"

def box(r1, c1, r2, c2, value=None, fill=None, font=None, align="left", valign="top", border=True):
    db.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cc = db.cell(row=r, column=c)
            if fill: cc.fill = fill
            if border: cc.border = BORDER
    tl = db.cell(row=r1, column=c1)
    if value is not None: tl.value = value
    if font: tl.font = font
    tl.alignment = Alignment(horizontal=align, vertical=valign, wrap_text=True)

db.row_dimensions[1].height = 28
db.merge_cells("A1:H1")
db["A1"].value = "Fußbodenheizung – überschlägige Auslegung"
db["A1"].font = Font(name=FONT, bold=True, color=NAVY, size=18); db["A1"].alignment = Alignment(vertical="center")
db.merge_cells("A2:H2"); db["A2"].value = "Eingaben · Verarbeitung · Ergebnisse"; db["A2"].font = f(color=GREY, size=12)
db.merge_cells("A3:K3")
db["A3"].value = f'="Projekt-Nr.: "&{gPNr}&"          Projekt: "&{gPName}&"          Bearbeiter: "&{gBearb}'
db["A3"].font = f(bold=True, color=NAVY, size=11)
add_logo_corner(db, 11, height=34)

# Spalten-Köpfe
box(5, 1, 5, 3, "INPUT  ·  Eingaben", fill=HDR_FILL, font=f(bold=True, color=WHITE, size=12), align="center", valign="center")
box(5, 5, 5, 7, "VERARBEITUNG  ·  Modell", fill=RED_S, font=f(bold=True, color=WHITE, size=12), align="center", valign="center")
box(5, 9, 5, 11, "OUTPUT  ·  Ergebnisse", fill=GREEN_S, font=f(bold=True, color=WHITE, size=12), align="center", valign="center")
# Pfeile
box(6, 4, 24, 4, "→", font=f(bold=True, color="E2001A", size=20), align="center", valign="center", border=False)
box(6, 8, 24, 8, "→", font=f(bold=True, color="E2001A", size=20), align="center", valign="center", border=False)
# Inhalte
INP = ("Global\n• Vor-/Rücklauftemperatur\n• Rohrsystem (da, s, di, k), Stoffwerte\n"
       "• α, Estrich, Dämmung nach unten\n• Grenzwerte (Kreislänge, Δp, v)\n\n"
       "Je Raum\n• Fläche (Raum / aktivierbar)\n• Heizlast, Raumtemperatur\n"
       "• R-Wert Belag, Verlegeabstand\n• Anzahl Kreise, Zuleitung, Zone, HKV\n\n"
       "Kalibrierung\n• Korrekturfaktor f")
VER = ("• log. Übertemperatur ΔθH\n• K_H · η (Verlegeabstand)\n• q = f · η · K_H · ΔθH\n"
       "• Oberflächentemperatur θF\n• Heizlastdeckung\n• Wärmeabgabe nach unten\n"
       "• maßgebl. Leistung → Massenstrom\n• Darcy-Weisbach: v, Re, λ, Δp\n• Rohrlängen je Kreis")
OUT = ("Je Raum\n• spez. Heizleistung q, θF\n• Deckungsgrad (Ampel)\n• Rohrlänge je Kreis (Ampel)\n"
       "• Massen-/Volumenstrom, v\n• Druckverlust je Kreis (Ampel)\n\n"
       "Je Verteiler (HKV)\n• Leistung, Massen-/Volumenstrom\n• max. Druckverlust\n\n"
       "Gesamt\n• Kontrolle & Warnübersicht")
box(6, 1, 24, 3, INP, fill=BODY_S, font=f(color=CHAR if False else BLACK, size=10.5), align="left", valign="top")
box(6, 5, 24, 7, VER, fill=BODY_S, font=f(color=BLACK, size=10.5), align="left", valign="top")
box(6, 9, 24, 11, OUT, fill=BODY_S, font=f(color=BLACK, size=10.5), align="left", valign="top")
for r in range(6, 25): db.row_dimensions[r].height = 13
db.row_dimensions[5].height = 22
# Vereinfachungen
box(26, 1, 26, 11, "Vereinfachungen (überschlägig)", fill=AMB_H, font=f(bold=True, color=AMBER, size=11), align="left", valign="center")
box(27, 1, 29, 11,
    "Vereinfachtes K_H + Verlegeabstand-Faktor statt vollständigem EN 1264   ·   "
    "Rohrwandwiderstand nicht enthalten (→ Korrekturfaktor f)   ·   "
    "Wärmeabgabe nach unten global (gleiche Dämmung)   ·   "
    "pauschaler Druckverlust-Zuschlag + fester Verteiler-Aufschlag",
    fill=AMB_B, font=f(color=BLACK, size=10.5), align="left", valign="center")
db.row_dimensions[26].height = 18
db.page_setup.orientation = "landscape"
db.page_setup.paperSize = 9
db.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
db.page_setup.fitToWidth = 1; db.page_setup.fitToHeight = 1
db.print_area = "A1:K29"
db.page_margins.left = db.page_margins.right = 0.3
db.page_margins.top = db.page_margins.bottom = 0.3

# =====================================================================
#  Blatt 1: ANLEITUNG  (zweites Blatt)
# =====================================================================
an = wb.create_sheet("Anleitung", 1)
an.sheet_view.showGridLines = False
for col, w in (("A", 2), ("B", 60), ("C", 3), ("D", 60)):
    an.column_dimensions[col].width = w

def anl(col, start, items):
    rr = start
    for text, kind in items:
        c = an.cell(row=rr, column=col, value=text)
        if kind == "h": c.font = Font(name=FONT, bold=True, color=NAVY, size=12)
        elif kind == "sh": c.font = Font(name=FONT, bold=True, color=BLACK, size=10.5)
        elif kind == "i": c.font = Font(name=FONT, italic=True, color=GREY, size=9)
        else: c.font = Font(name=FONT, size=10)
        c.alignment = Alignment(vertical="top")
        rr += 1
    return rr

left = [
    ("Aufbau der Mappe", "h"),
    ("• Deckblatt – Übersicht (Input / Verarbeitung / Output)", ""),
    ("• Anleitung – diese Seite", ""),
    ("• Grundeinstellungen – globale Werte", ""),
    ("• Auslegung – Raumliste mit allen Berechnungen", ""),
    ("• Kontrolle – Summen & Warnungen", ""),
    ("• HKV – Auswertung je Heizkreisverteiler", ""),
    ("• Verifikation – Abgleich mit Herstellerdaten", ""),
    ("• Methodik – Formeln & Annahmen", ""),
    ("• Konstanten – Nachschlagetabellen", ""),
    ("• Changelog – Versionen", ""),
    ("", ""),
    ("Spalten der Auslegung", "h"),
    ("Aus Revit (A–F):", "sh"),
    ("HKV, Raum-Nr., Bezeichnung, Raumfläche,", ""),
    ("Raumtemperatur, Heizlast.", ""),
    ("Berechnet (G):", "sh"),
    ("spez. Heizlast – nicht überschreiben.", ""),
    ("Im Tool ergänzen (H–M):", "sh"),
    ("aktiv. Fläche, R-Wert, Verlegeabstand,", ""),
    ("Anz. Kreise, Zuleitung, Zone.", ""),
    ("", ""),
    ("Farblegende", "h"),
    ("• Blau (gelb hinterlegt) = Eingabe", ""),
    ("• Schwarz = berechnet (nicht ändern)", ""),
    ("• Grün = OK,  Rot = Warnung / Grenzwert überschritten", ""),
    ("• Nur die blauen Eingabezellen bearbeiten.", ""),
]
right = [
    ("Schritt für Schritt", "h"),
    ("1. Grundeinstellungen ausfüllen: Projekt, Temperaturen,", ""),
    ("    Rohrsystem, Stoffwerte und Grenzwerte (Kreislänge,", ""),
    ("    Druckverlust, Geschwindigkeit, Volumenstrom je Kreis).", ""),
    ("2. Konstanten prüfen / erweitern (Rohre, R-Werte, Zonen).", ""),
    ("3. In Revit die vorgefertigte Bauteilliste öffnen – sie ist", ""),
    ("    bereits passend angelegt und enthält die Spalten A–F", ""),
    ("    (HKV, Raum-Nr., Bezeichnung, Raumfläche,", ""),
    ("    Raumtemperatur, Heizlast).", ""),
    ("4. Bauteilliste aus Revit nach Excel exportieren", ""),
    ("    ('Exportieren → Bericht/Schedule').", ""),
    ("5. Werte aus dem Export in die Auslegung übernehmen", ""),
    ("    (Spalten A–F, ab der ersten Datenzeile).", ""),
    ("6. Spalte G (spez. Heizlast) frei lassen – wird berechnet.", ""),
    ("7. Spalten H–M je Raum ergänzen: aktivierbare Fläche,", ""),
    ("    R-Wert Bodenbelag, Verlegeabstand, Anzahl Kreise,", ""),
    ("    Zuleitungslänge, Zone.", ""),
    ("8. Ergebnisse prüfen: Ampel in der Auslegung und im", ""),
    ("    Blatt 'Kontrolle' (Deckung, Kreislänge, Druckverlust,", ""),
    ("    Geschwindigkeit, Volumenstrom je Kreis).", ""),
    ("9. Verifikation: Korrekturfaktor f an Herstellerdaten", ""),
    ("    kalibrieren.", ""),
    ("", ""),
    ("Hinweis", "sh"),
    ("Die Bauteilliste in Revit muss nicht erst angelegt werden –", ""),
    ("sie ist passend zur Auslegung vorbereitet. Bei Bedarf lassen", ""),
    ("sich auch die Felder H–M (z. B. Zone, Verlegeabstand) als", ""),
    ("gemeinsame Parameter in Revit pflegen und mit exportieren.", ""),
]
end_l = anl(2, 3, left)
end_r = anl(4, 3, right)
disp_header(an, "Anleitung / Bedienung", 4, project=False)
setup_print(an, f"A1:D{max(end_l, end_r)}")
wb.active = 0

# =====================================================================
wb.properties.version = VERSION
wb.properties.creator = AUTHOR
wb.properties.lastModifiedBy = AUTHOR
wb.properties.keywords = f"FBH-Auslegung v{VERSION} ({AUTHOR})"
try:
    wb.calculation.fullCalcOnLoad = True
except Exception:
    pass
OUT = r"C:\Users\derDi\FBH_Auslegung_Excel\FBH_Auslegung.xlsx"
wb.save(OUT)
print("gespeichert:", OUT, "| Version", VERSION, AUTHOR)
