# Rechnungsprogramm

Ein Programm zur Erstellung von Rechnungen für MacBook.

## Funktionen

- Eingabemaske für Firmendaten (inkl. Logo)
- Datenbank für Mandanten
- Datenbank für Leistungen
- Zuordnung von Leistungen zu Terminen
- Rechnungserstellung für Zeiträume
- Mehrwertsteuerberechnung
- Export als Word-Dokument
- Statistikmodul

## Installation

1. Python 3.8+ installieren.
2. Abhängigkeiten installieren: `pip install -r requirements.txt`
3. Datenbank initialisieren: `python db_init.py`
4. Programm starten: `python main.py`

## Executable erstellen

Für Windows/Mac:
1. PyInstaller installieren: `pip install pyinstaller`
2. Bauen: `pyinstaller --onefile main.py`
3. Das Executable ist in `dist/main` (Linux) oder entsprechend.

Für Mac: Auf Mac bauen, da GUI abhängig von OS.

## GitHub

Das Repository ist auf GitHub: https://github.com/ophtahackathon/Rechnungsprogramm

Um zu pushen:
1. `git add .`
2. `git commit -m "Initial commit"`
3. `git push origin main`

## Vorhandene Lösungen

Basierend auf Recherche, ähnliche Programme:
- InvoicePlane (Open Source)
- Dolibarr (ERP mit Rechnungen)
- Odoo (Modular, inkl. Rechnungen)

Dieses Programm ist eine einfache, maßgeschneiderte Lösung.