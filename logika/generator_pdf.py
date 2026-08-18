import os
import time
import win32com.client
import pythoncom
import gc
from pathlib import Path


def export_to_pdf(excel_path):
    excel_path_str = os.path.abspath(excel_path).replace("/", "\\")
    pdf_path_str = excel_path_str.replace(".xlsx", ".pdf")

    excel = None
    wb = None
    new_wb = None  # Temporary workbook containing the worksheet selected for export

    try:
        pythoncom.CoInitialize()
        time.sleep(1)

        if not os.path.exists(excel_path_str):
            raise Exception(f"BŁĄD: Plik nie istnieje pod adresem -> {excel_path_str}")

        # A separate Excel instance prevents conflicts with other open workbooks
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False

        # Open the source workbook in read-only mode
        wb = excel.Workbooks.Open(excel_path_str, 0, True)
        source_worksheet = wb.Worksheets(1)

        # Copy the first worksheet to export it without the remaining tabs
        source_worksheet.Copy()
        new_wb = excel.ActiveWorkbook
        export_worksheet = new_wb.Worksheets(1)

        # Fit the worksheet to a single PDF page
        export_worksheet.PageSetup.Zoom = False
        export_worksheet.PageSetup.FitToPagesWide = 1
        export_worksheet.PageSetup.FitToPagesTall = 1

        # Export the workbook containing only the selected worksheet
        new_wb.ExportAsFixedFormat(0, pdf_path_str, 0, True, False, 1, 1, False)

    except Exception as e:
        raise Exception(f"Błąd silnika PDF: {e}")

    finally:
        # Close Excel objects regardless of the export result
        if new_wb is not None:
            try:
                # Close the temporary workbook without saving changes
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

        # Release references to COM objects
        new_wb = None
        wb = None
        excel = None
        gc.collect()

        pythoncom.CoUninitialize()

    return Path(pdf_path_str)
