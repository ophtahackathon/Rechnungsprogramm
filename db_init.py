import sqlite3

def init_db():
    conn = sqlite3.connect('rechnungen.db')
    c = conn.cursor()

    # Tabelle für Firma
    c.execute('''CREATE TABLE IF NOT EXISTS firma (
                    id INTEGER PRIMARY KEY,
                    name TEXT,
                    adresse TEXT,
                    logo_path TEXT,
                    mwst_satz REAL DEFAULT 19.0,
                    mwst_text TEXT DEFAULT 'Es wird keine Mehrwertsteuer erhoben'
                )''')

    # Tabelle für Mandanten
    c.execute('''CREATE TABLE IF NOT EXISTS mandanten (
                    id INTEGER PRIMARY KEY,
                    name TEXT,
                    adresse TEXT,
                    email TEXT,
                    telefon TEXT
                )''')

    # Tabelle für Leistungen
    c.execute('''CREATE TABLE IF NOT EXISTS leistungen (
                    id INTEGER PRIMARY KEY,
                    kuerzel TEXT,
                    beschreibung TEXT,
                    betrag REAL
                )''')

    # Tabelle für Termine
    c.execute('''CREATE TABLE IF NOT EXISTS termine (
                    id INTEGER PRIMARY KEY,
                    mandant_id INTEGER,
                    datum TEXT,
                    FOREIGN KEY (mandant_id) REFERENCES mandanten(id)
                )''')

    # Tabelle für Termin-Leistungen
    c.execute('''CREATE TABLE IF NOT EXISTS termin_leistungen (
                    id INTEGER PRIMARY KEY,
                    termin_id INTEGER,
                    leistung_id INTEGER,
                    anzahl INTEGER DEFAULT 1,
                    FOREIGN KEY (termin_id) REFERENCES termine(id),
                    FOREIGN KEY (leistung_id) REFERENCES leistungen(id)
                )''')

    # Tabelle für Rechnungen
    c.execute('''CREATE TABLE IF NOT EXISTS rechnungen (
                    id INTEGER PRIMARY KEY,
                    mandant_id INTEGER,
                    von_datum TEXT,
                    bis_datum TEXT,
                    summe_netto REAL,
                    mwst REAL,
                    summe_brutto REAL,
                    erstellt_am TEXT,
                    FOREIGN KEY (mandant_id) REFERENCES mandanten(id)
                )''')

    # Tabelle für Rechnungspositionen
    c.execute('''CREATE TABLE IF NOT EXISTS rechnung_positionen (
                    id INTEGER PRIMARY KEY,
                    rechnung_id INTEGER,
                    leistung_id INTEGER,
                    anzahl INTEGER,
                    betrag REAL,
                    FOREIGN KEY (rechnung_id) REFERENCES rechnungen(id),
                    FOREIGN KEY (leistung_id) REFERENCES leistungen(id)
                )''')

    conn.commit()
    conn.close()

if __name__ == '__main__':
    init_db()