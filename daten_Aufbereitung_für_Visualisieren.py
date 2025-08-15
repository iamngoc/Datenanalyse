from typing import Final
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt


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

# ==== Benutzerabfrage ====
YEAR: Final[int] = ask_for_year()
NUMBER_OF_PRODUCTS: Final[int] = 15
TOP_PRODUCTS_OF_GROUP: Final[int] = 50
ARCHIV_DIR = Path("archiv_Excel")

EXCEL_FILE = ARCHIV_DIR / f"auftraege_{YEAR}.xlsx"
ARTIKEL_FILE = ARCHIV_DIR / "Artikel_20250624.xlsx"

# ==== Hilfsfunktionen ====

def read_data():
    """Liest Excel-Dateien ein und gibt DataFrames zurück."""
    sheets = pd.read_excel(EXCEL_FILE, sheet_name=None)
    artikel_sheet = pd.read_excel(ARTIKEL_FILE)

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


def top_products_per_month(product_counts, top_n):
    """Ermittelt Top-N Produkte pro Monat."""
    return (
        product_counts
        .sort_values(["month", "Number of Usages"], ascending=[True, False])
        .groupby("month")
        .head(top_n)
    )

def save_to_excel(df, filename):
    """Speichert DataFrame in Excel-Datei."""
    df.to_excel(ARCHIV_DIR / filename, index=False)

# ==== Hauptlogik ====

def main():

        auftraege, artikel, artikel_gruppe = read_data()
        auftraege, artikel, artikel_gruppe = clean_and_convert(auftraege, artikel, artikel_gruppe)

        auftraege_filted = filter_by_year(auftraege, YEAR)
        merged_gr = merge_artikel_with_gruppe(artikel, artikel_gruppe)
        merged = merge_with_month_data(merged_gr, auftraege_filted)
        save_to_excel(merged, "Zusammengefuehrte_Datei.xlsx")

        product_counts = count_products_per_month(merged)
        save_to_excel(product_counts, f"Anzahl_genutzte_Produkte_{YEAR}.xlsx")

        # Top15 pro Monat (für mögliche spätere Auswertungen)
        top_products = top_products_per_month(product_counts, NUMBER_OF_PRODUCTS)
        save_to_excel(top_products, f"top{NUMBER_OF_PRODUCTS}_pro_monat_{YEAR}.xlsx")

if __name__ == "__main__":
    main()
