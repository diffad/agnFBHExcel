# Weiterarbeiten auf einem anderen Rechner (z. B. Linux Mint)

Dieser Ordner enthält eine Kopie der **globalen Hausstil-Konfiguration** für Claude Code,
damit das Excel-Design (Graustufen, Deckblatt, Changelog …) auch auf einem neuen Rechner
standardmäßig genutzt wird. Der Ordner ist nur für die Einrichtung – das Projekt selbst
braucht ihn nicht.

## 1. Werkzeuge installieren (Linux Mint)

```bash
sudo apt update
sudo apt install -y git python3 python3-pip libreoffice
pip3 install --user openpyxl pymupdf
```

Claude Code (CLI) installieren und anmelden – siehe https://claude.com/claude-code .

## 2. Projekt holen

```bash
git clone https://github.com/diffad/agnFBHExcel.git
cd agnFBHExcel
python3 build_fbh_en1264.py      # erzeugt FBH_Auslegung_EN1264.xlsx (Pfade sind plattformunabhängig)
```

## 3. Globalen Hausstil einrichten (einmalig je Rechner)

```bash
mkdir -p ~/.claude/templates
cp claude-setup/CLAUDE.global.md          ~/.claude/CLAUDE.md
cp claude-setup/templates/agn_xlsx_style.py   ~/.claude/templates/
cp claude-setup/templates/agn_xlsx_example.py ~/.claude/templates/
```

Damit liest Claude Code in **jedem** Projekt auf diesem Rechner den Excel-Hausstil.

## Hinweise

- `~` ist unter Linux `/home/<name>`, unter Windows `C:\Users\<name>` – die Pfade passen auf beiden.
- Falls du die globale Config später änderst, einfach die Datei `~/.claude/CLAUDE.md` (bzw. die
  Vorlage in `~/.claude/templates/`) bearbeiten. Diese Kopie hier ist nur der Übertragungsstand.
- Projekt-spezifisches Wissen (z. B. die FBH-Methodik) merkt sich Claude pro Projekt selbst neu;
  der globale Teil ist nur der allgemeine Tabellen-Hausstil.
- Eine kleine Beispiel-Mappe erzeugen: `python3 ~/.claude/templates/agn_xlsx_example.py`.
