import os
import sys
import win32com.client as win32

def convert_xlsx_to_xlsb(xlsx_path, xlsb_path):
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False

        workbook = excel.Workbooks.Open(os.path.abspath(xlsx_path))

        workbook.SaveAs(os.path.abspath(xlsb_path), FileFormat=50)

        workbook.Close(SaveChanges=False)
        excel.Quit()

        print(f"Conversão concluída: {xlsb_path}")

    except Exception as e:
        print(f"Erro ao converter arquivo: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Uso: python converter.py <caminho_xlsx> <caminho_xlsb>")
        sys.exit(1)
    
    xlsx_file = sys.argv[1]
    xlsb_file = sys.argv[2]

    convert_xlsx_to_xlsb(xlsx_file, xlsb_file)
