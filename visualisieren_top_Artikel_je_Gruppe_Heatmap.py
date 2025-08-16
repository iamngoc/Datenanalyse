"""
Dieses Skript fragt den Benutzer nach einem Jahr und einer Anzahl N, die der am meisten verwendeten Artikel je Artikelgruppe pro Monat im Jahr entspricht.
Es lädt anschließend die vorbereiteten Excel-Daten (Nutzungszahlen von Artikeln) für das gewählte Jahr,
ermittelt die Top-N Artikel je Artikelgruppe pro Monat und erstellt daraus ein Heatmap.
Die Diagramme werden automatisch in einem passenden Ausgabeordner gespeichert.

WICHTIG: Sie müssen zuerst das Skript "daten_Aufarbeitung_für_Visualisieren.py" ausführen!
"""
"""
BITTE ÄNDERN SIE DEN PFAD FÜR archive_dir IN MAIN FUNKTION - ZEILE 150

"""

from pathlib import Path

import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# ==== Hilfsfunktionen ====
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
    """Fragt den Benutzer nach einer Anzahl der Artikeln und gibt es als int zurück."""
    while True:
        try:
            top_articles_of_group = int(input("Geben Sie die Anzahl der am meistens verwendeten Artikel je Gruppe ein? "))
            if 1 <= top_articles_of_group <= 200:
                return top_articles_of_group
            else:
                print("Bitte geben Sie eine Nummer zwischen 1 und 200 ein.")
        except ValueError:
            print("Ungültige Eingabe! Bitte eine ganze Zahl eingeben.")


def load_data(year: int, archive_dir: Path) -> pd.DataFrame:
    """Lädt die Excel-Datei mit den Nutzungszahlen."""
    file_path = archive_dir / f"Anzahl_genutzte_Produkte_{year}.xlsx"
    if not file_path.exists():
        raise FileNotFoundError(f"Datei nicht gefunden: {file_path}")
    return pd.read_excel(file_path)

def prepare_output_folder(year: int, top_articles_of_group: int, archive_dir: Path) -> Path:
    """Erstellt den Ausgabeordner für Plots."""
    out_dir = archive_dir / f"Plots_top_{top_articles_of_group}_je_Gruppe_{year}"
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir


def top_n_preview_all_groups_multisheet(df: pd.DataFrame, top_n: int, year: int,
                                        archive_dir: Path, save_excel: bool = True) -> dict:
    """
    Gibt die Top-N Produkte je Monat für alle Gruppen zurück.
    - Jede Gruppe bekommt ein eigenes Excel-Blatt
    - Alle Blätter in einer Datei
    """
    all_results = {}
    for gruppe, daten in df.groupby("ArtikelGruppe"):
        # Top-N je Monat
        top_per_month = (
            daten.sort_values(["month", "Number of Usages"], ascending=[True, False])
                  .groupby("month", as_index=False)
                  .head(top_n)
                  .copy()
        )

        top_per_month["Rank"] = (
            top_per_month.groupby("month")["Number of Usages"]
                         .rank(method="first", ascending=False)
                         .astype(int)
        )

        preview = (
            top_per_month[["month", "Rank", "ArtikelCode", "Number of Usages"]]
            .sort_values(["month", "Rank"])
            .reset_index(drop=True)
        )

        all_results[gruppe] = preview

    if save_excel:
        out = archive_dir / f"Check_Top{top_n}_AllGroups_{year}.xlsx"
        with pd.ExcelWriter(out) as writer:
            for gruppe, df_gruppe in all_results.items():
                sheet_name = str(gruppe)[:31]  # Excel-Blattname max. 31 Zeichen
                df_gruppe.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"Multi-Sheet-Excel gespeichert: {out}")

    return all_results

def create_heatmaps(df: pd.DataFrame, top_per_group: int, output_dir: Path):
    """Erstellt Balkendiagramme für jede Artikelgruppe."""
    sns.set_theme(style="whitegrid")

    for gruppe, daten in df.groupby("ArtikelGruppe"):
        # Top-N pro Monat auswählen
        top_per_month = (
            daten.sort_values(["month", "Number of Usages"], ascending=[True, False])
            .groupby("month")
            .head(top_per_group)
        )

        # Pivot-Tabelle: Artikel = Zeilen, Monate = Spalten
        pivot_df = top_per_month.pivot_table(
            index="ArtikelCode",
            columns="month",
            values="Number of Usages",
            fill_value=0
        )

        # Sicherstellen, dass alle Monate 1–12 drin sind
        for m in range(1, 13):
            if m not in pivot_df.columns:
                pivot_df[m] = 0
        pivot_df = pivot_df[sorted(pivot_df.columns)]

        # Heatmap erstellen
        plt.figure(figsize=(40, 20))
        sns.heatmap(
            pivot_df,
            annot=True, fmt=".0f", cmap="YlOrRd", linewidths=0.5,
            cbar_kws={"label": "Nutzungen"}
        )
        plt.title(f"Heatmap – Top {top_per_group} Artikel pro Monat – Gruppe: {gruppe}")
        plt.xlabel("Monat")
        plt.ylabel("ArtikelCode")

        # Speichern
        filename = output_dir / f"heatmap_top{top_per_group}_{gruppe}.png"
        plt.savefig(filename, dpi=100, bbox_inches="tight")
        plt.close()


# ==== Hauptlogik ====
def main():
    year = ask_for_year()
    """
    BITTE ÄNDERN SIE DEN PFAD FÜR archive_dir hier
    """
    archive_dir = Path(f"C:/Users/minhn/Documents/IPH_Praktikum/Aufgabe_Python/Archive_Excels_{year}")

    top_articles_of_group = ask_for_n_articles()
    df = load_data(year, archive_dir)
    top_per_group = top_articles_of_group

    output_dir = prepare_output_folder(year, top_articles_of_group, archive_dir)
    _ = top_n_preview_all_groups_multisheet(df,
                                   top_n=top_articles_of_group,
                                   year=year,
                                   archive_dir=archive_dir,
                                   save_excel=True)

    create_heatmaps(df, top_per_group, output_dir)
    print("Diagramme und Check-Excels wurden erstellt.")


if __name__ == "__main__":
    main()
