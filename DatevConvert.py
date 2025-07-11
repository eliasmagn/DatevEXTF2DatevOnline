
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import logging
import sys
import csv
import os
import json
from datetime import datetime
import getpass
import re

DEBUG_MODE = "--debug" in sys.argv
CONFIG_PATH = "datev_konverter_config.json"

def setup_logging(debug=DEBUG_MODE):
    level = logging.DEBUG if debug else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler("konverter.log", mode='w', encoding="utf-8"),
            logging.StreamHandler(sys.stdout)
        ]
    )
    logging.info("Starte Konverter (Debug=%s)", debug)

# Exakte DATEV-Kassenbuchfelder in der korrekten Reihenfolge
DATEV_HEADERS = [
    "Währung", "VorzBetrag", "RechNr", "BelegDatum", "Belegtext",
    "UStSatz", "BU", "Gegenkonto", "Kost1", "Kost2", "Kostmenge", "Skonto", "Nachricht"
]

# Mapping aus dem EXTF-Export ins DATEV-Format (Anpassen falls Feldnamen abweichen!)
EXTF_TO_DATEV = {
    "Währung": "WKZ Umsatz",
    "VorzBetrag": "Umsatz (ohne Soll/Haben-Kz)",
    "SollHaben": "Soll/Haben-Kennzeichen",
    "RechNr": "Belegfeld_1",
    "BelegDatum": "Belegdatum",
    "Belegtext": "Buchungstext",
    "UStSatz": "Beleginfo - Inhalt 2",
    "BU": "BU-Schlüssel",
    "Gegenkonto": "Gegenkonto (ohne BU-Schlüssel)",
    "Kost1": "KOST1 - Kostenstelle",
    "Kost2": "KOST2 - Kostenstelle",
    "Kostmenge": "Kost-Menge",
    "Skonto": "Skonto",
    "Nachricht": "",  # Optional/Freitext
}

def vorzeichen_betrag(betrag_raw, sollhaben):
    # Betrag nach DATEV: Komma als Dezimaltrennzeichen, Vorzeichen vorangestellt (+/-)
    betrag_str = str(betrag_raw).replace('.', '').replace(',', '.').strip()
    try:
        betrag = float(betrag_str)
    except Exception:
        logging.warning("Betrag '%s' nicht konvertierbar.", betrag_raw)
        return ''
    # Soll: Minus, Haben: Plus
    if str(sollhaben).strip().upper() == "S":
        val = f"-{abs(betrag):.2f}".replace('.', ',')
    elif str(sollhaben).strip().upper() == "H":
        val = f"+{abs(betrag):.2f}".replace('.', ',')
    else:
        val = f"{betrag:.2f}".replace('.', ',')
    logging.debug("vorzeichen_betrag: %s %s -> %s", betrag_raw, sollhaben, val)
    return val

def belegdatum_formatieren(dateval):
    # DATEV verlangt TTMM (z.B. 0207 für 2. Juli)
    try:
        # Versucht, verschiedene Formate zu akzeptieren
        for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d.%m.%y", "%d%m"):
            try:
                dt = datetime.strptime(str(dateval), fmt)
                return f"{dt.day:02d}{dt.month:02d}"
            except Exception:
                continue
        # Falls bereits TTMM, nichts ändern
        if re.match(r"^\d{4}$", str(dateval)):
            return str(dateval)
    except Exception:
        pass
    return ""

def robust_datev_import_mit_korrektur(dateipfad, delimiter=";", encoding="latin1"):
    rows = []
    with open(dateipfad, encoding=encoding) as f:
        reader = list(csv.reader(f, delimiter=delimiter, quoting=csv.QUOTE_MINIMAL))
        header = reader[1]
        erwartete_spalten = len(header)
        for i, row in enumerate(reader[2:]):  # ab Zeile 3 (Index 2)
            if len(row) != erwartete_spalten:
                fehlerzeile = delimiter.join(row)
                logging.warning("Zeile %d hat %d statt %d Spalten: %s", i+3, len(row), erwartete_spalten, fehlerzeile)
                korrigiert = zeile_korrigieren_gui(i+2, fehlerzeile, delimiter)
                if korrigiert is None:
                    logging.info("User hat Abbruch gewählt (bei Zeile %d).", i+3)
                    return None, header
                new_row = korrigiert.split(delimiter)
                while len(new_row) != erwartete_spalten:
                    messagebox.showwarning(
                        "Korrektur noch fehlerhaft",
                        f"Die Korrektur ergibt {len(new_row)} Spalten, erwartet werden {erwartete_spalten}.\nBitte erneut anpassen."
                    )
                    korrigiert = zeile_korrigieren_gui(i+2, korrigiert, delimiter)
                    if korrigiert is None:
                        logging.info("User hat Abbruch gewählt (bei Zeile %d)", i+3)
                        return None, header
                    new_row = korrigiert.split(delimiter)
                rows.append(new_row)
                logging.info("Zeile %d erfolgreich korrigiert.", i+3)
            else:
                rows.append(row)
    df = pd.DataFrame(rows, columns=header, dtype=str)
    return df, header

def zeile_korrigieren_gui(zeilennr, fehlerzeile, delimiter=";"):
    korrigiert = []
    abgebrochen = []
    def submit():
        korrigiert.append(textfeld.get("1.0", tk.END).strip())
        fenster.destroy()
    def abbrechen():
        abgebrochen.append(True)
        fenster.destroy()
    fenster = tk.Tk()
    fenster.title(f"Fehlerhafte Zeile {zeilennr+1} manuell korrigieren")
    tk.Label(fenster, text="Bitte korrigieren Sie die Zeile so, dass sie exakt zur Kopfzeile passt.\nAbbrechen bricht die gesamte Konvertierung ab!").pack(pady=4)
    textfeld = scrolledtext.ScrolledText(fenster, width=120, height=4, wrap=tk.NONE)
    textfeld.insert(tk.END, fehlerzeile)
    textfeld.pack(padx=8, pady=8)
    button_frame = tk.Frame(fenster)
    button_frame.pack(pady=8)
    tk.Button(button_frame, text="Korrigiert übernehmen", command=submit, bg="#4caf50", fg="white", width=22).pack(side=tk.LEFT, padx=5)
    tk.Button(button_frame, text="Abbrechen", command=abbrechen, bg="#d32f2f", fg="white", width=22).pack(side=tk.LEFT, padx=5)
    fenster.mainloop()
    if abgebrochen:
        return None
    return korrigiert[0] if korrigiert else fehlerzeile

def konvertieren(quellpfad, zielpfad):
    setup_logging()
    df, header = robust_datev_import_mit_korrektur(quellpfad, delimiter=";", encoding="latin1")
    if df is None:
        messagebox.showwarning("Abbruch", "Die Konvertierung wurde abgebrochen.")
        return

    output_rows = []
    for idx, row in df.iterrows():
        buchung = {}
        # Währung
        buchung["Währung"] = row.get(EXTF_TO_DATEV["Währung"], "EUR") or "EUR"
        # Betrag
        betrag = vorzeichen_betrag(row.get(EXTF_TO_DATEV["VorzBetrag"], ""), row.get(EXTF_TO_DATEV["SollHaben"], ""))
        buchung["VorzBetrag"] = betrag
        # Rechnungsnr/Feld 1
        buchung["RechNr"] = row.get(EXTF_TO_DATEV["RechNr"], "")
        # BelegDatum in TTMM
        buchung["BelegDatum"] = belegdatum_formatieren(row.get(EXTF_TO_DATEV["BelegDatum"], ""))
        # Freitext
        buchung["Belegtext"] = row.get(EXTF_TO_DATEV["Belegtext"], "")
        # USt, BU, Gegenkonto usw.
        for feld in ["UStSatz", "BU", "Gegenkonto", "Kost1", "Kost2", "Kostmenge", "Skonto", "Nachricht"]:
            quelle = EXTF_TO_DATEV[feld]
            value = row.get(quelle, "") if quelle else ""
            if feld == "BU":
                if str(value).strip().lower() in ["0", "null", "none", "nan"]:
                    value = ""
            buchung[feld] = value
        output_rows.append(buchung)
    df_out = pd.DataFrame(output_rows, columns=DATEV_HEADERS)
    try:
        df_out.to_csv(zielpfad, sep=";", index=False, encoding="utf-8", header=True, quoting=csv.QUOTE_MINIMAL)
        msg = f"Datei erfolgreich konvertiert: {zielpfad}"
        logging.info(msg)
        messagebox.showinfo("Fertig", msg)
    except Exception as e:
        logging.error("Fehler beim Speichern: %s", e)
        messagebox.showerror("Fehler beim Speichern", str(e))

def load_config():
    if os.path.isfile(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_config(cfg):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f)

def vorschlagsname_from(quellpfad, rule="Konvertiert_{basename}_{date}.csv"):
    basename = os.path.splitext(os.path.basename(quellpfad))[0]
    now = datetime.now()
    year = now.strftime("%Y")
    month = now.strftime("%m")
    mon = now.strftime("%b")
    week = now.strftime("%V")
    day = now.strftime("%d")
    date = now.strftime("%Y%m%d-%H%M")
    user = getpass.getuser()
    match = re.search(r'_von_(\d{4})_(\d{2})_(\d{2})_bis_(\d{4})_(\d{2})_(\d{2})', basename)
    von, bis, zeitraum, monat, jahr = '', '', '', '', ''
    if match:
        jahr_von, monat_von, tag_von, jahr_bis, monat_bis, tag_bis = match.groups()
        von = f"{jahr_von}-{monat_von}-{tag_von}"
        bis = f"{jahr_bis}-{monat_bis}-{tag_bis}"
        monat = monat_von
        jahr = jahr_von
        zeitraum = f"{jahr_von}-{monat_von}"
    else:
        von = bis = zeitraum = ''
        monat = month
        jahr = year
    vorschlag = rule.format(
        basename=basename, date=date, user=user,
        year=year, month=month, mon=mon, week=week, day=day,
        von=von, bis=bis, monat=monat, jahr=jahr, zeitraum=zeitraum
    )
    ordner = os.path.dirname(quellpfad)
    return os.path.join(ordner, vorschlag)

def gui_start():
    config = load_config()
    last_quell = config.get("last_quell", "")
    last_ziel = config.get("last_ziel", "")
    name_rule = config.get("name_rule", "Konvertiert_{basename}_{date}.csv")
    root = tk.Tk()
    root.title("DATEV-Kassenbuch Konverter")
    frame = tk.Frame(root, padx=15, pady=15)
    frame.pack()
    tk.Label(frame, text="Regel für Ausgabedatei:").grid(row=0, column=0, sticky="w")
    rule_entry = tk.Entry(frame, width=50)
    rule_entry.grid(row=0, column=1, columnspan=2, sticky="w")
    rule_entry.insert(0, name_rule)
    tk.Label(frame, text="Platzhalter: {basename}, {date}, {year}, {month}, {mon}, {week}, \n {day}, {user}, {von}, {bis}, {monat}, {jahr}, {zeitraum}").grid(row=1, column=1, sticky="w")
    tk.Label(frame, text="Quell-CSV (DATEV-Export):").grid(row=2, column=0, sticky="w")
    quell_entry = tk.Entry(frame, width=50)
    quell_entry.grid(row=2, column=1)
    quell_entry.insert(0, last_quell)
    def quell_browse():
        path = filedialog.askopenfilename(filetypes=[("CSV Dateien", "*.csv")])
        if path:
            quell_entry.delete(0, tk.END)
            quell_entry.insert(0, path)
            rule = rule_entry.get()
            ziel_vorschlag = vorschlagsname_from(path, rule)
            ziel_entry.delete(0, tk.END)
            ziel_entry.insert(0, ziel_vorschlag)
    tk.Button(frame, text="Durchsuchen...", command=quell_browse).grid(row=2, column=2)
    tk.Label(frame, text="Ziel-CSV (für Import):").grid(row=3, column=0, sticky="w")
    ziel_entry = tk.Entry(frame, width=50)
    ziel_entry.grid(row=3, column=1)
    ziel_entry.insert(0, last_ziel)
    def ziel_browse():
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Dateien", "*.csv")])
        if path:
            ziel_entry.delete(0, tk.END)
            ziel_entry.insert(0, path)
    tk.Button(frame, text="Speichern unter...", command=ziel_browse).grid(row=3, column=2)
    def start_konvertierung():
        quell = quell_entry.get()
        ziel = ziel_entry.get()
        rule = rule_entry.get()
        if not quell or not ziel:
            messagebox.showerror("Fehler", "Bitte Quell- und Zieldatei auswählen!")
            return
        konvertieren(quell, ziel)
        cfg = {
            "last_quell": quell,
            "last_ziel": ziel,
            "name_rule": rule
        }
        save_config(cfg)
    tk.Button(frame, text="Konvertieren", command=start_konvertierung, width=20, bg="#4caf50", fg="white").grid(row=4, column=1, pady=15)
    root.mainloop()

if __name__ == "__main__":
    gui_start()
