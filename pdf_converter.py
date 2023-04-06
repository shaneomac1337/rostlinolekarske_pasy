import os
import openpyxl
import win32com.client

def save_excel_as_pdf(excel_file, sheet_name):
    """Save the given sheet in excel file as pdf in the 'pdf' folder."""
    pdf_folder = "pdf"
    if not os.path.exists(pdf_folder):
        os.makedirs(pdf_folder)

    pdf_file = f"{pdf_folder}/{sheet_name}.pdf"

    try:
        xlApp = win32com.client.Dispatch("Excel.Application")
        xlApp.Visible = False

        wb = xlApp.Workbooks.Open(os.path.abspath(excel_file), ReadOnly=1)
        ws = wb.Worksheets(sheet_name)
        ws.ExportAsFixedFormat(0, os.path.abspath(pdf_file))

    except Exception as e:
        print(f"Failed to convert {excel_file} - {sheet_name} to PDF: {e}")

    finally:
        wb.Close(SaveChanges=False)
        xlApp.Quit()

def save_all_excels_as_pdfs():
    has_excel_files = False
    for filename in os.listdir('.'):
        if filename.endswith('.xlsx') and filename != 'template.xlsx':
            has_excel_files = True
            wb = openpyxl.load_workbook(filename)
            for sheet in wb:
                save_excel_as_pdf(filename, sheet.title)
                print(f"Uloženo jako:{sheet.title}\n")

    if not has_excel_files:
        print("A teď zase Olinka nedodala Excely na konvertování do PDF. Bože muj.")
    else:
        print("Všechny listy v excelu byly uloženy jako samostatné PDF.")

if __name__ == '__main__':
    save_all_excels_as_pdfs()
