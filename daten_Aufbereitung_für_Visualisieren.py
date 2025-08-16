"""
Dieses Skript dient der Datenaufbereitung für spätere Visualisierungen von Artikel-Nutzungen.

Ablauf:
1. Abfrage des Analysejahres vom Benutzer.
2. Einlesen von Aufträgen und Artikeldaten aus den Excel-Dateien:
   - auftraege_{year}.xlsx
   - Artikel_20250624.xlsx
3. Bereinigung und Vereinheitlichung der Datentypen.
4. Filterung der Aufträge auf das gewählte Jahr und Ergänzung einer Monats-Spalte.
5. Verknüpfung der Artikel mit ihren Artikelgruppen sowie den Aufträgen.
6. Zählung der Artikel-Nutzungen pro Monat und Gruppe.
7. Export:
   - Zusammengefuehrte_Datei.xlsx (Rohdaten)
   - Anzahl_genutzte_Produkte_{year}.xlsx (Auswertung nach Monaten)

Ergebnis:
Die erzeugten Excel-Dateien bilden die Grundlage für Visualisierungen
(z. B. Top-N-Analysen oder Heatmaps) und werden in einem Archivordner gespeichert.
"""
"""
BITTE ÄNDERN SIE DIE PFADE FÜR excel_file, artikel_file, path_to_create - ZEILE 53, 54, 55

"""

import os

from pathlib import Path
import pandas as pd

# ==== Jahr abfragen ====
def ask_for_year() -> int:
    """Fragt den Benutzer nach einem Jahr und gibt es als int zurück."""
    while True:
        try:
            year = int(input("Welches Jahr möchten Sie auswerten? "))
            if 2000 <= year <= 2100:  # Plausibilitätscheck
                return year
            else:
                print("⚠ Bitte geben Sie ein Jahr zwischen 2000 und 2100 ein.")
        except ValueError:
            print("⚠ Ungültige Eingabe! Bitte eine ganze Zahl eingeben.")

year = ask_for_year()

"""
BITTE ÄNDERN SIE DIE 3 PFADE: excel_file, artikel_file, path_to_create
excel_file: der Pfad für Excel-Datei "auftraege_2023.xlsx"
artikel_file: der Pfad für Excel-Datei "Artikel_20250624.xlsx"
path_to_create: Ordner für aufgebreitete Dateien
"""

excel_file = f"C:/Users/minhn/Documents/IPH_Praktikum/Aufgabe_Python/given_Excel_datas/auftraege_{year}.xlsx"
artikel_file = f"C:/Users/minhn/Documents/IPH_Praktikum/Aufgabe_Python/given_Excel_datas/Artikel_20250624.xlsx"
path_to_create = Path(f"C:/Users/minhn/Documents/IPH_Praktikum/Aufgabe_Python/Archive_Excels_{year}")

def create_archive_dir() -> Path:
    path = path_to_create
    try:
        os.makedirs(path, exist_ok=True)  # exist_ok=True = kein Fehler wenn Ordner schon da
    except OSError as e:
        print(f"Fehler beim Erstellen des Ordners: {e}")
    return path

archive_dir = create_archive_dir()

# ==== Hilfsfunktionen ====

def read_data():
    """Liest Excel-Dateien ein und gibt DataFrames zurück."""
    sheets = pd.read_excel(excel_file, sheet_name=None)
    artikel_sheet = pd.read_excel(artikel_file)

    auftraege = sheets['Auftraege'][['AU_Nummer', 'ProduktionsEnde']].copy()
    artikel = sheets['Auftraege_Artikel'][['AU_Nummer', 'ArtikelCode']].copy()
    artikel_gruppe = artikel_sheet[['ArtikelCode', 'ArtikelGruppe']].copy()

    return auftraege, artikel, artikel_gruppe

def clean_and_convert(auftraege, artikel, artikel_gruppe):
    """Bereinigt Datentypen und konvertiert Datumswerte."""
    # Einheitliche Datentypen für Merge
    auftraege['AU_Nummer'] = auftraege['AU_Nummer'].astype(str)
    artikel['AU_Nummer'] = artikel['AU_Nummer'].astype(str)
    artikel_gruppe['ArtikelCode'] = artikel_gruppe['ArtikelCode'].astype(str)
    artikel['ArtikelCode'] = artikel['ArtikelCode'].astype(str)

    # ProduktionsEnde zu datetime konvertieren
    auftraege['ProduktionsEnde'] = pd.to_datetime(auftraege['ProduktionsEnde'], errors='coerce')

    return auftraege, artikel, artikel_gruppe

def filter_by_year(auftraege, year):
    """Filtert Aufträge nach Jahr und fügt Monatsspalte hinzu."""
    df = auftraege[auftraege['ProduktionsEnde'].dt.year == year].copy()
    df['month'] = df['ProduktionsEnde'].dt.month
    return df

def merge_artikel_with_gruppe(artikel, artikel_gruppe):
    """Fügt ArtikelGruppe zu Artikeln hinzu."""
    merged = pd.merge(
        artikel,
        artikel_gruppe,
        left_on=artikel['ArtikelCode'].str.lower(),
        right_on=artikel_gruppe['ArtikelCode'].str.lower(),
        how="left"
    )
    merged['ArtikelCode'] = merged['ArtikelCode_x']
    merged = merged.drop(columns=['key_0', 'ArtikelCode_x', 'ArtikelCode_y'])

    if merged.empty:
        print("⚠ Merge fehlgeschlagen: keine gemeinsamen ArtikelCode gefunden.")
    return merged


def merge_with_month_data(artikel_gruppe_df, auftraege_filted):
    """Verknüpft Artikel mit Monatsinformationen."""
    merged = pd.merge(
        artikel_gruppe_df,
        auftraege_filted[['AU_Nummer', 'month']],
        on="AU_Nummer",
        how="left"
    )
    if merged.empty:
        print("⚠ Merge fehlgeschlagen: keine gemeinsamen AU_Nummern gefunden.")
    return merged


def count_products_per_month(merged):
    """Zählt Produktnutzungen pro Monat und gibt DataFrame zurück."""
    return (
        merged.groupby(["month", "ArtikelCode", "ArtikelGruppe"])
        .size()
        .reset_index(name="Number of Usages")
    )


def save_to_excel(df, filename):
    """Speichert DataFrame in Excel-Datei."""
    df.to_excel(archive_dir / filename, index=False)

# ==== Hauptlogik ====

def main():

        auftraege, artikel, artikel_gruppe = read_data()
        auftraege, artikel, artikel_gruppe = clean_and_convert(auftraege, artikel, artikel_gruppe)

        auftraege_filted = filter_by_year(auftraege, year)
        merged_gr = merge_artikel_with_gruppe(artikel, artikel_gruppe)
        merged = merge_with_month_data(merged_gr, auftraege_filted)
        save_to_excel(merged, "Zusammengefuehrte_Datei.xlsx")

        product_counts = count_products_per_month(merged)
        save_to_excel(product_counts, f"Anzahl_genutzte_Produkte_{year}.xlsx")

if __name__ == "__main__":
    main()
