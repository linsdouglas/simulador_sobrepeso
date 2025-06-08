
import sys
import os
import comtypes.client
import time

def exportar_pdf(file_path, aba_nome, nome_remessa):
    comtypes.CoInitialize()
    excel = comtypes.client.CreateObject("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(file_path)
    ws = wb.Worksheets(aba_nome)

    pdf_dir = os.path.join(os.path.dirname(file_path), 'Relatório_Saida')
    os.makedirs(pdf_dir, exist_ok=True)
    pdf_path = os.path.join(pdf_dir, f"SOBREPESOSIMULADO - {nome_remessa}.pdf")

    ws.ExportAsFixedFormat(Type=0, Filename=pdf_path)
    wb.Close(False)
    excel.Quit()
    comtypes.CoUninitialize()

    print(pdf_path)

if __name__ == "__main__":
    file_path = sys.argv[1]
    aba_nome = sys.argv[2]
    nome_remessa = sys.argv[3]
    exportar_pdf(file_path, aba_nome, nome_remessa)
