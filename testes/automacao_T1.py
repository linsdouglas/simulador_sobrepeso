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
import ctypes
import shutil
import win32com.client
win32com.client.gencache.is_readonly = False
win32com.client.gencache.Rebuild()

temp_copy_path = None
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

def encontrar_pasta_onedrive_empresa():
    user_dir = os.environ["USERPROFILE"]
    possiveis = os.listdir(user_dir)
    for nome in possiveis:
        if "DIAS BRANCO" in nome.upper():
            caminho_completo = os.path.join(user_dir, nome)
            if os.path.isdir(caminho_completo) and "Gestão de Estoque - Documentos" in os.listdir(caminho_completo):
                return os.path.join(caminho_completo, "Gestão de Estoque - Documentos")
    return None
unidade='Simulador_T1'
fonte_dir = encontrar_pasta_onedrive_empresa()
#fonte_dir = os.path.join(os.environ['USERPROFILE'], 'Downloads') 
if not fonte_dir:
    raise FileNotFoundError("Não foi possível localizar a pasta sincronizada do SharePoint via OneDrive.")

download_dir = os.path.join(fonte_dir, 'SIMULADOR_T1')
consol_path = os.path.join(download_dir,"Consol.xlsx")
simulador_path = os.path.join(download_dir,"Simulador_T1.xlsx")
dashboard_path = os.path.join(download_dir,"DASHBOARD_FRETE.xlsx")

def copiar_arquivo_temporario(origem, log_callback):
    try:
        temp_dir = tempfile.gettempdir()
        destino = os.path.join(temp_dir, os.path.basename(origem))
        shutil.copy2(origem, destino)
        desbloquear_arquivo(destino, log_callback)
        log_callback(f"Arquivo copiado para uso temporário: {destino}")
        return destino
    except Exception as e:
        log_callback(f"Erro ao copiar arquivo temporário: {str(e)}")
        raise

def desbloquear_arquivo(path, log_callback):
    try:
        if os.path.exists(path + ":Zone.Identifier"):
            os.remove(path + ":Zone.Identifier")
        ctypes.windll.kernel32.DeleteFileW(f"{path}:Zone.Identifier")
    except Exception as e:
        log_callback(f"Aviso: falha ao remover bloqueio de segurança: {e}")

def abrir_workbook(excel, path, log_callback, tentativas=3, espera=2):
    for i in range(tentativas):
        try:
            return excel.Workbooks.Open(path)
        except Exception as e:
            log_callback(f"Tentativa {i + 1} falhou ao abrir {path}: {e}")
            time.sleep(espera)
            if i == tentativas - 1:
                log_callback(f"Falha definitiva ao abrir {path}")
                raise

def executar_processo(mapa_frete_path, mes_usuario, log_callback):
    arquivos_temp = []

    try:
        log_callback("Iniciando processo sequencial...")

        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.Interactive = False

        # Copiar arquivos
        mapa_temp = copiar_arquivo_temporario(mapa_frete_path, log_callback)
        consol_temp = copiar_arquivo_temporario(consol_path, log_callback)
        simulador_temp = copiar_arquivo_temporario(simulador_path, log_callback)
        dashboard_temp = copiar_arquivo_temporario(dashboard_path, log_callback)

        arquivos_temp.extend([mapa_temp, consol_temp, simulador_temp, dashboard_temp])

        # Passo 1 - Processar Mapa e Consol
        log_callback("Abrindo Mapa Frete...")
        mapa_wb = abrir_workbook(excel, mapa_temp, log_callback)

        log_callback("Abrindo Consol...")
        consol_wb = abrir_workbook(excel, consol_temp, log_callback)

        ws_mapa = consol_wb.Sheets("Mapa Frete")
        last_row = ws_mapa.Cells(ws_mapa.Rows.Count, "AD").End(-4162).Row
        ws_mapa.Range(f"AD4:DL{last_row}").Clear()

        ws_origem = mapa_wb.Sheets(1)
        last_row_map = ws_origem.Cells(ws_origem.Rows.Count, "A").End(-4162).Row
        ws_origem.Range(f"A2:CI{last_row_map}").Copy()
        ws_mapa.Range(f"AD4:DL{last_row_map}").PasteSpecial(Paste=-4104)

        mapa_wb.Close(SaveChanges=False)

        log_callback("Atualizando pivôs do Consol...")
        for ws in consol_wb.Sheets:
            try:
                for pt in ws.PivotTables():
                    pt.RefreshTable()
            except Exception as e:
                log_callback(f"Aviso: erro ao atualizar pivô na aba {ws.Name}: {e}")

        consol_wb.Close(SaveChanges=True)

        # Passo 2 - Processar Simulador
        log_callback("Abrindo Simulador...")
        simulador_wb = abrir_workbook(excel, simulador_temp, log_callback)

        simulador_wb.Sheets("CÓD").Range("B3").Value = mes_usuario
        simulador_wb.Sheets("Base Transf Real").Range("E3:T2299").ClearContents()

        log_callback("Reabrindo Consol para extrair Link Real T1...")
        consol_wb = abrir_workbook(excel, consol_temp, log_callback)

        source_range = consol_wb.Sheets("Link Real T1").Range("A3:P2299")
        dest_range = simulador_wb.Sheets("Base Transf Real").Range("E3:T2299")
        dest_range.Value = source_range.Value

        consol_wb.Close(SaveChanges=False)

        log_callback("Atualizando pivôs do Simulador...")
        for ws in simulador_wb.Sheets:
            try:
                for pt in ws.PivotTables():
                    pt.RefreshTable()
            except Exception as e:
                log_callback(f"Aviso: erro ao atualizar pivô na aba {ws.Name}: {e}")

        efeitos_range = simulador_wb.Sheets("Efeitos Regional").Range("B3:I17")
        efeitos_valor = efeitos_range.Value
        log_callback("Efeitos capturados da aba 'Efeitos Regional' como matriz de valores.")
        simulador_wb.Close(SaveChanges=True)

        # Passo 3 - Processar Dashboard
        log_callback("Abrindo Dashboard...")
        dashboard_wb = abrir_workbook(excel, dashboard_temp, log_callback)

        try:
            bd_sheet = dashboard_wb.Sheets("BD")
            bd_range = bd_sheet.Range("A1").Resize(len(efeitos_valor), len(efeitos_valor[0]))
            bd_range.Value = efeitos_valor
            log_callback("Efeitos colados na aba 'BD' do DASHBOARD_FRETE.")
        except Exception as e:
            log_callback(f"Aviso: não foi possível colar efeitos na aba 'BD': {e}")

        dashboard_wb.Close(SaveChanges=True)

        excel.Quit()

        log_callback("Processo sequencial concluído com sucesso.")

    except Exception as e:
        log_callback(f"Erro: {str(e)}\n{traceback.format_exc()}")

    finally:
        for temp_file in arquivos_temp:
            try:
                if os.path.exists(temp_file):
                    os.remove(temp_file)
                    log_callback(f"Arquivo temporário removido: {temp_file}")
            except Exception as e:
                log_callback(f"Não foi possível remover a cópia temporária {temp_file}: {str(e)}")

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