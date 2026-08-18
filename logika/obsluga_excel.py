import json
import openpyxl
from pathlib import Path
from openpyxl.styles import Alignment
import datetime

def generuj_pojedynczy_psw(dane_z_formularza):
    """
    Mapuje dane formularza do szablonu PSW i zapisuje dokument Excel.
    """
    # Katalog główny projektu
    BASE_DIR = Path(__file__).resolve().parent.parent

    # Katalog nadrzędny dla wygenerowanych dokumentów
    ROOT_DIR = BASE_DIR.parent

    sciezka_json = BASE_DIR / "konfiguracje_json" / "vda_config.json"
    sciezka_excel = BASE_DIR / "puste_formatki" / "vda_2020.xlsx"

    # Pobranie numeru części i wersji raportu
    part_number = str(dane_z_formularza.get("PartNumber", "Brak_PN")).strip()
    wersja_raportu = str(dane_z_formularza.get("ReportVersion", "00")).strip()

    # Folder docelowy: vda/<numer części>/<wersja raportu>
    docelowy_folder = ROOT_DIR / "vda" / part_number / wersja_raportu

    # Utworzenie folderu docelowego, jeśli nie istnieje
    docelowy_folder.mkdir(parents=True, exist_ok=True)

    # Domyślna nazwa pliku wynikowego
    sciezka_zapisu = docelowy_folder / f"PSW_{part_number}.xlsx"

    # Dodanie daty i licznika zapobiega nadpisaniu istniejącego pliku
    licznik = 1

    today = datetime.datetime.now().strftime("%Y-%m-%d")

    while sciezka_zapisu.exists():
        nowa_nazwa = f"PSW_{part_number}_{today}_{licznik}.xlsx"
        sciezka_zapisu = docelowy_folder / nowa_nazwa
        licznik += 1

    # Wczytanie mapowania pól i szablonu Excel
    with open(sciezka_json, 'r', encoding='utf-8') as f:
        mapa = json.load(f)

    wb = openpyxl.load_workbook(sciezka_excel)
    sheet = wb.active

    # Uzupełnienie komórek zgodnie z konfiguracją
    for klucz, wartosc in dane_z_formularza.items():
        if klucz in mapa["Fields"]:
            konf = mapa["Fields"][klucz]
            komorka = sheet[konf["cell"]]
            typ = konf.get("type", "standard")

            if typ in ["checkbox", "merged_checkbox"]:
                if wartosc is True:
                    komorka.value = "X"
                    komorka.alignment = Alignment(horizontal='center', vertical='center')
            elif typ == "merged_multiline":
                komorka.value = wartosc
                komorka.alignment = Alignment(wrapText=True, vertical='top')
            else:
                komorka.value = wartosc

    wb.save(sciezka_zapisu)
    wb.close()
    return sciezka_zapisu