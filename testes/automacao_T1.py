import os
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox
from datetime import datetime
import win32com.client as win32
import pandas as pd
import tempfile
import threading
import traceback
import time

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

fonte_dir="Simulador_T1"
download_dir = os.path.join(os.environ['USERPROFILE'], 'Downloads', fonte_dir)

def executar_processo(mapa_frete_path, mes_usuario, log_callback):
    try:
        log_callback("Iniciando processo...")
        print("iniciando processo")

        consol_path = os.path.join(download_dir,"Consol.xlsx")
        simulador_path = os.path.join(download_dir,"Simulador T1_AA_Mar v.2.xlsx")
        dashboard_path = os.path.join(download_dir,"DASHBOARD_FRETE.xlsx")

        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False

        log_callback("Abrindo arquivos...")
        print("abrindo arquivos")
        mapa_wb = excel.Workbooks.Open(mapa_frete_path)
        consol_wb = excel.Workbooks.Open(consol_path)
        simulador_wb = excel.Workbooks.Open(simulador_path)
        dashboard_wb = excel.Workbooks.Open(dashboard_path)
        time.sleep(2)

        simulador_wb.Sheets("CÓD").Range("B3").Value = mes_usuario

        ws_mapa = consol_wb.Sheets("Mapa Frete")
        last_row = ws_mapa.Cells(ws_mapa.Rows.Count, "AD").End(-4162).Row
        ws_mapa.Range(f"AD4:DL{last_row}").Clear()

        ws_origem = mapa_wb.Sheets(1)
        last_row_map = ws_origem.Cells(ws_origem.Rows.Count, "A").End(-4162).Row
        ws_origem.Range(f"A2:CI{last_row_map}").Copy()
        ws_mapa.Range(f"AD4:DL{last_row_map}").PasteSpecial(Paste=-4104)

        log_callback("Atualizando pivôs do consolidado...")
        for ws in consol_wb.Sheets:
            try:
                for pt in ws.PivotTables():
                    pt.RefreshTable()
            except Exception as e:
                log_callback(f"Aviso: erro ao atualizar pivô na aba {ws.Name} - {str(e)}")

        simulador_wb.Sheets("Base Transf Real").Range("E3:T2299").ClearContents()
        source_range = consol_wb.Sheets("Link Real T1").Range("A3:P2299")
        dest_range = simulador_wb.Sheets("Base Transf Real").Range("E3:T2299")
        dest_range.ClearContents()
        dest_range.Value = source_range.Value

        log_callback("Atualizando pivôs do simulador...")
        for ws in simulador_wb.Sheets:
            try:
                for pt in ws.PivotTables():
                    pt.RefreshTable()
            except Exception as e:
                log_callback(f"Aviso: erro ao atualizar pivô na aba {ws.Name} - {str(e)}")

        efeitos_range = simulador_wb.Sheets("Efeitos Regional").Range("B3:I17")

        try:
            bd_sheet = dashboard_wb.Sheets("BD")
            efeitos_range.Copy()
            bd_sheet.Range("A1").PasteSpecial(Paste=-4163)
            log_callback("Efeitos colados na aba 'BD' do DASHBOARD_FRETE.")
        except Exception as e:
            log_callback(f"Aviso: não foi possível acessar a aba 'BD' do DASHBOARD_FRETE - {str(e)}")

        log_callback("Finalizando e salvando arquivos...")
        mapa_wb.Close(SaveChanges=False)
        consol_wb.Close(SaveChanges=True)
        simulador_wb.Close(SaveChanges=True)
        dashboard_wb.Close(SaveChanges=True)
        excel.Quit()

        log_callback("Processo concluído com sucesso.")

    except Exception as e:
        log_callback(f"Erro: {str(e)}\n{traceback.format_exc()}")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Automação Dashboard")
        self.geometry("700x500")

        self.mapa_path = tk.StringVar()
        self.mes_usuario = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        ctk.CTkLabel(self, text="Arquivo Mapa de Frete (.xlsx):").pack(pady=10)
        ctk.CTkEntry(self, textvariable=self.mapa_path, width=500).pack(pady=5)
        ctk.CTkButton(self, text="Selecionar Arquivo", command=self.selecionar_arquivo).pack(pady=5)

        ctk.CTkLabel(self, text="Mês (1-12):").pack(pady=10)
        ctk.CTkEntry(self, textvariable=self.mes_usuario).pack(pady=5)

        ctk.CTkButton(self, text="Executar Processo", command=self.iniciar_processo).pack(pady=20)

        self.log_box = ctk.CTkTextbox(self, width=600, height=200)
        self.log_box.pack(pady=10)

    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(filetypes=[
                        ("Excel or CSV files", "*.xlsm *.xlsx *.csv"),
                        ("Excel files", "*.xlsm *.xlsx"),
                        ("CSV files", "*.csv")
                    ])
        if caminho:
            self.mapa_path.set(caminho)

    def log(self, mensagem):
        timestamp = datetime.now().strftime("[%H:%M:%S]")
        self.log_box.insert("end", f"{timestamp} {mensagem}\n")
        self.log_box.see("end")

    def iniciar_processo(self):
        mapa = self.mapa_path.get() 
        mes = self.mes_usuario.get()

        if not mapa or not os.path.isfile(mapa):
            messagebox.showerror("Erro", "Selecione um arquivo válido.")
            return

        if mapa.endswith('.csv'):
            caminho_excel = mapa
            if not caminho_excel:
                return
        else:
            caminho_excel = mapa

        def processo_em_thread():
            self.log("Iniciando execução...")
            executar_processo(caminho_excel, mes, self.log)

        thread = threading.Thread(target=processo_em_thread)
        thread.start()

if __name__ == "__main__":
    app = App()
    app.mainloop()