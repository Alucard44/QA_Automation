import os
import time
import win32com.client
import pythoncom
import gc
from pathlib import Path


def eksportuj_do_pdf(sciezka_excel):
    sciezka_excel_str = os.path.abspath(sciezka_excel).replace("/", "\\")
    sciezka_pdf_str = sciezka_excel_str.replace(".xlsx", ".pdf")

    excel = None
    wb = None
    new_wb = None  # Tymczasowy skoroszyt zawierający eksportowany arkusz

    try:
        pythoncom.CoInitialize()
        time.sleep(1)

        if not os.path.exists(sciezka_excel_str):
            raise Exception(f"BŁĄD: Plik nie istnieje pod adresem -> {sciezka_excel_str}")

        # Osobna instancja Excela ogranicza konflikty z otwartymi skoroszytami
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False

        # Otwarcie skoroszytu źródłowego tylko do odczytu
        wb = excel.Workbooks.Open(sciezka_excel_str, 0, True)
        arkusz = wb.Worksheets(1)

        # Skopiowanie pierwszego arkusza pozwala wyeksportować go bez pozostałych zakładek
        arkusz.Copy()
        new_wb = excel.ActiveWorkbook
        nowy_arkusz = new_wb.Worksheets(1)

        # Dopasowanie arkusza do jednej strony PDF
        nowy_arkusz.PageSetup.Zoom = False
        nowy_arkusz.PageSetup.FitToPagesWide = 1
        nowy_arkusz.PageSetup.FitToPagesTall = 1

        # Eksport skoroszytu zawierającego wyłącznie wybrany arkusz
        new_wb.ExportAsFixedFormat(0, sciezka_pdf_str, 0, True, False, 1, 1, False)

    except Exception as e:
        raise Exception(f"Błąd silnika PDF: {e}")

    finally:
        # Zamknięcie obiektów Excela niezależnie od wyniku eksportu
        if new_wb is not None:
            try:
                # Zamknięcie skoroszytu tymczasowego bez zapisywania zmian
                new_wb.Close(SaveChanges=False)
            except:
                pass
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except:
                pass

        # Zwolnienie referencji do obiektów COM
        new_wb = None
        wb = None
        excel = None
        gc.collect()

        pythoncom.CoUninitialize()

    return Path(sciezka_pdf_str)