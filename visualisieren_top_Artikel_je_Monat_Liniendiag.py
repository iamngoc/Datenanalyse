"""
Dieses Skript fragt den Benutzer nach einem Jahr und einer Anzahl N, die der am meisten verwendeten Artikel pro Monat im Jahr entspricht.
Es lädt anschließend die vorbereiteten Excel-Daten (Nutzungszahlen von Artikeln) für das gewählte Jahr,
ermittelt die Top-N Artikel pro Monat und erstellt daraus ein Liniendiagramm.
Die Diagramme werden automatisch in einem passenden Ausgabeordner gespeichert.

WICHTIG: Sie müssen zuerst das Skript "daten_Aufarbeitung_für_Visualisieren.py" ausführen!
"""
"""
BITTE ÄNDERN SIE DEN PFAD FÜR archive_dir IN MAIN FUNKTION - ZEILE 110

"""

from pathlib import Path

import pandas as pd
import matplotlib.pyplot as plt

def ask_for_year() -> int:
    """Fragt den Benutzer nach einem Jahr und gibt es als int zurück."""
    while True:
        try:
            year = int(input("Welches Jahr möchten Sie auswerten? "))
            if 2000 <= year <= 2100:
                return year
            else:
                print("Bitte geben Sie ein Jahr zwischen 2000 und 2100 ein.")
        except ValueError:
            print("Ungültige Eingabe! Bitte eine ganze Zahl eingeben.")

def ask_for_n_articles() -> int:
    """Fragt den Benutzer nach einer Anzahl der am meistens Artikelnpro Monat und gibt es als int zurück."""
    while True:
        try:
            top_n_articles_per_month = int(input("Geben Sie die Anzahl der am meistens verwendeten Artikel je Monat ein? "))
            if 1 <= top_n_articles_per_month <= 1000:
                return top_n_articles_per_month
            else:
                print("Bitte geben Sie eine Nummer zwischen 1 und 1000 ein.")
        except ValueError:
            print("Ungültige Eingabe! Bitte eine ganze Zahl eingeben.")

def load_data(year: int, archive_dir: Path) -> pd.DataFrame:
    """Lädt die Excel-Datei mit den Verwendungszahlen."""
    file_path = archive_dir / f"Anzahl_genutzte_Produkte_{year}.xlsx"
    if not file_path.exists():
        raise FileNotFoundError(f"Datei nicht gefunden: {file_path}")
    return pd.read_excel(file_path)

def top_articles_per_month(df: pd.DataFrame, n: int) -> pd.DataFrame:
    """Ermittelt die Top-N Produkte pro Monat."""
    return (
        df.sort_values(["month", "Number of Usages"], ascending=[True, False])
          .groupby("month")
          .head(n)
          .reset_index(drop=True)
    )

def top_articles_per_month(df: pd.DataFrame, n: int) -> pd.DataFrame:
    df_sorted = (
        df.sort_values(["month", "Number of Usages"], ascending=[True, False])
          .groupby("month")
          .head(n)
          .reset_index(drop=True)
    )
    # Rank je Monat hinzufügen
    df_sorted["Rank"] = (
        df_sorted.groupby("month")["Number of Usages"]
                 .rank(method="first", ascending=False)
                 .astype(int)
    )
    return df_sorted

def prepare_output_folder(year: int, top_n_articles_per_month: int, archive_dir: Path) -> Path:
    """Erstellt den Ausgabeordner für Plots."""
    out_dir = archive_dir / f"Top_{top_n_articles_per_month}_articles_of_month_{year}"
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir

def create_line_chart (df: pd.DataFrame, year: int, top_n_articles_per_month: int, output_dir: Path):
    pivot_df = df.pivot(index='month', columns='ArtikelCode', values='Number of Usages').fillna(0)
    pivot_df.plot(kind='line', marker='o', figsize=(20, 14))

    plt.title(f"Verwendungsverlauf der Top {top_n_articles_per_month} Artikel pro Monat im Jahr {year}")
    plt.xlabel("Monat")
    plt.ylabel("Anzahl der Verwendungen")
    plt.xticks(range(1, 13))
    plt.grid(True)
    plt.legend(title="ArtikelCode", bbox_to_anchor=(1.05, 1), loc='upper left')

    # Speichern
    filename = output_dir / f"Top {top_n_articles_per_month} Artikel pro Monat im Jahr {year}.png"
    plt.savefig(filename, dpi=100, bbox_inches="tight")
    plt.close()

def save_top_n_to_excel(df: pd.DataFrame, year: int, n: int, output_dir: Path):
    """Speichert die Top-N Produkte pro Monat in eine Excel-Datei."""
    out = output_dir / f"Top_{n}_Artikel_pro_Monat_{year}.xlsx"
    df_sorted = df.sort_values(["month", "Rank"])
    df_sorted.to_excel(out, index=False)
    print(f"Kontroll-Excel gespeichert: {out}")


def main():
    year = ask_for_year()
    top_n_articles_per_month = ask_for_n_articles()
    """
    BITTE ÄNDERN SIE DEN PFAD FÜR archive_dir hier
    """
    archive_dir = Path(f"C:/Users/minhn/Documents/IPH_Praktikum/Aufgabe_Python/Archive_Excels_{year}")

    df = load_data(year, archive_dir)
    df_top = top_articles_per_month(df, top_n_articles_per_month)

    output_dir = prepare_output_folder(year, top_n_articles_per_month, archive_dir)

    create_line_chart(df_top, year, top_n_articles_per_month, output_dir)
    save_top_n_to_excel(df_top, year, top_n_articles_per_month, output_dir)

    print("Diagramme und Kontroll-Excel wurden erstellt.")


if __name__ == "__main__":
    main()
