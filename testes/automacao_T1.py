import os
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox
from datetime import datetime
import win32com.client as win32

# Configurações do CustomTkinter
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Função principal do processo
fonte_dir="Simulador_T1"
download_dir = os.path.join(os.environ['USERPROFILE'], 'Downloads', fonte_dir)
def executar_processo(mapa_frete_path, mes_usuario, log_callback):
    try:
        log_callback("Iniciando processo...")

        consol_path = os.path.join(download_dir,"Consol.xlsx")
        simulador_path = os.path.join(download_dir,"Simulador_T1_FINAL.xlsx")

        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False

        # Abrindo os arquivos
        log_callback("Abrindo arquivos...")
        mapa_wb = excel.Workbooks.Open(mapa_frete_path)
        consol_wb = excel.Workbooks.Open(consol_path)
        simulador_wb = excel.Workbooks.Open(simulador_path)

        # Inputando mês
        simulador_wb.Sheets("CÓD").Range("B3").Value = mes_usuario

        # Limpando antigo mapa de frete
        ws_mapa = consol_wb.Sheets("Mapa Frete")
        last_row = ws_mapa.Cells(ws_mapa.Rows.Count, "AD").End(-4162).Row
        ws_mapa.Range(f"AD4:DL{last_row}").Clear()

        # Copiando novo mapa
        ws_origem = mapa_wb.Sheets("mapa_frete")
        last_row_map = ws_origem.Cells(ws_origem.Rows.Count, "A").End(-4162).Row
        ws_origem.Range(f"A2:CI{last_row_map}").Copy()
        ws_mapa.Range(f"AD4:DL{last_row_map}").PasteSpecial(Paste=-4104)

        # Atualiza Pivôs do consol
        log_callback("Atualizando pivôs do consolidado...")
        for ws in consol_wb.Sheets:
            for pt in ws.PivotTables():
                pt.RefreshTable()

        # Cópia dos dados de Link Real T1 para Base Transf Real
        consol_wb.Sheets("Link Real T1").Range("C3:P2299").Copy()
        simulador_wb.Sheets("Base Transf Real").Range("C3").PasteSpecial(Paste=-4104)

        # Atualiza pivôs do simulador
        log_callback("Atualizando pivôs do simulador...")
        for ws in simulador_wb.Sheets:
            for pt in ws.PivotTables():
                pt.RefreshTable()

        # Cópia dos efeitos regionais
        efeitos_range = simulador_wb.Sheets("Efeitos Regional").Range("B3:I17")
        bd_sheet = mapa_wb.Sheets("BD")
        efeitos_range.Copy()
        bd_sheet.Range("A1").PasteSpecial(Paste=-4163)

        # Salvando e fechando
        log_callback("Finalizando e salvando arquivos...")
        mapa_wb.Close(SaveChanges=False)
        consol_wb.Close(SaveChanges=True)
        simulador_wb.Close(SaveChanges=True)
        excel.Quit()

        log_callback("✅ Processo concluído com sucesso.")

    except Exception as e:
        log_callback(f"Erro: {str(e)}")

# Interface Gráfica
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Dashboard de Frete - Executor")
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
        caminho = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsm")])
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

        if not mes.isdigit() or not (1 <= int(mes) <= 12):
            messagebox.showerror("Erro", "Digite um mês válido entre 1 e 12.")
            return

        self.log("🟡 Iniciando execução...")
        executar_processo(mapa, mes, self.log)

# Executar o app
if __name__ == "__main__":
    app = App()
    app.mainloop()
