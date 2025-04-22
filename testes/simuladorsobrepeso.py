import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from datetime import datetime
import customtkinter as ctk
from tkinter import messagebox, StringVar
import os
import win32com.client as win32

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

# === Função principal de cálculo ===
def calcular_peso_final(remessa_num, peso_veiculo_vazio, qtd_paletes, df_exp, df_sku, df_sap, df_sobrepeso, log_callback):
    try:
        remessa_num = int(remessa_num)
    except ValueError:
        log_callback("Remessa inválida.")
        return None

    df_remessa = df_exp[df_exp['REMESSA'] == remessa_num]
    if df_remessa.empty:
        log_callback("Remessa não encontrada em data_exp.")
        return None

    sku = df_remessa.iloc[0]['ITEM']
    qtd_caixas = df_remessa['QUANTIDADE'].sum()
    log_callback(f"SKU: {sku}, Quantidade de caixas: {qtd_caixas}")

    df_sku_filtrado = df_sku[(df_sku['COD_PRODUTO'] == sku) & (df_sku['DESC_UNID_MEDID'] == 'Caixa')]
    if df_sku_filtrado.empty:
        log_callback("SKU não encontrado ou unidade diferente de 'Caixa'.")
        return None

    peso_por_caixa = df_sku_filtrado.iloc[0]['QTDE_PESO_LIQ']
    peso_base = qtd_caixas * peso_por_caixa
    log_callback(f"Peso por caixa: {peso_por_caixa}, Peso base: {peso_base}")

    chaves_pallet = df_remessa['CHAVE_PALETE'].unique()
    df_pallets = df_sap[df_sap['Chave Pallet'].isin(chaves_pallet)]
    if df_pallets.empty:
        log_callback("Nenhum pallet encontrado em data_sap.")
        return None

    total_overweight_adjustment = 0
    for _, row in df_pallets.iterrows():
        lote = row['Lote']
        data_producao = row['Data de produção']
        hora_inicio = row['Hora de criação'][:2] + ":00:00"
        hora_fim = row['Hora de modificação'][:2] + ":00:00"

        linha_produzida = "L" + lote[-3:]
        log_callback(f"Processando pallet: Lote {lote}, Linha {linha_produzida}, Data {data_producao}")

        df_sp_filtro = df_sobrepeso[
            (df_sobrepeso['Linhas'] == linha_produzida) &
            (df_sobrepeso['Data e Hora'] >= f"{data_producao} {hora_inicio}") &
            (df_sobrepeso['Data e Hora'] <= f"{data_producao} {hora_fim}")
        ]

        if not df_sp_filtro.empty:
            media_sp = df_sp_filtro['Sobrepesohora'].mean()
            ajuste = peso_base * media_sp
            log_callback(f"Sobrepeso médio: {media_sp:.4f}, Ajuste: {ajuste:.2f}")
            total_overweight_adjustment += ajuste

    peso_paletes = qtd_paletes * 26
    peso_final = peso_veiculo_vazio + peso_base + total_overweight_adjustment + peso_paletes
    log_callback(f"Peso dos paletes: {peso_paletes}, Peso final: {peso_final}")

    return peso_base, total_overweight_adjustment, peso_paletes, peso_final
def preencher_formulario(wb, dados, sobrepesos_por_item):
    ws = wb["FORMULARIO"]
    ws["B4"] = dados['remessa']
    ws["B6"] = dados['qtd_skus']
    ws["B7"] = dados['placa']
    ws["B8"] = dados['turno']
    ws["B9"] = dados['peso_vazio']
    ws["B10"] = dados['peso_base']
    ws["B11"] = dados['sp_total']
    ws["B12"] = dados['peso_com_sp']
    ws["B13"] = dados['peso_total_final']
    ws["B14"] = dados['media_sp']
    ws["B16"] = dados['peso_total_final'] * 1.02
    ws["B17"] = dados['peso_total_final']
    ws["B18"] = dados['peso_total_final'] * 0.99
    ws["D4"] = dados['qtd_paletes']
    ws["D9"] = dados['qtd_paletes'] * 26

    linha = 12
    for sku, sp in sobrepesos_por_item.items():
        ws[f"C{linha}"] = sku
        ws[f"D{linha}"] = sp
        linha += 1

def exportar_para_pdf(caminho_arquivo, aba_nome="FORMULARIO"):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(caminho_arquivo)
    ws = wb.Worksheets(aba_nome)
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = 1
    pdf_path = caminho_arquivo.replace(".xlsm", f"_{aba_nome}.pdf")
    ws.Range("A1:D49").ExportAsFixedFormat(0, pdf_path)
    wb.Close(False)
    excel.Quit()
    return pdf_pat

# === GUI Principal ===
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Simulador de Sobrepeso")
        self.geometry("700x600")

        self.placa = StringVar()
        self.turno = StringVar()
        self.remessa = StringVar()
        self.qtd_paletes = StringVar()
        self.peso_vazio = StringVar()
        self.peso_balanca = StringVar()
        self.log_text = []

        ctk.CTkLabel(self, text="Placa:").pack()
        ctk.CTkEntry(self, textvariable=self.placa).pack()

        ctk.CTkLabel(self, text="Turno:").pack()
        ctk.CTkComboBox(self, values=["Manhã", "Tarde", "Noite"], variable=self.turno).pack()

        ctk.CTkLabel(self, text="Remessa:").pack()
        ctk.CTkEntry(self, textvariable=self.remessa).pack()

        ctk.CTkLabel(self, text="Quantidade de Paletes:").pack()
        ctk.CTkEntry(self, textvariable=self.qtd_paletes).pack()

        ctk.CTkLabel(self, text="Peso Veículo Vazio:").pack()
        ctk.CTkEntry(self, textvariable=self.peso_vazio).pack()

        ctk.CTkLabel(self, text="Peso Final Balança:").pack()
        ctk.CTkEntry(self, textvariable=self.peso_balanca).pack()

        ctk.CTkButton(self, text="Calcular", command=self.processar).pack(pady=10)

        ctk.CTkLabel(self, text="Histórico de Logs:").pack(pady=5)
        self.log_display = ctk.CTkLabel(self, text="", wraplength=600, justify="left")
        self.log_display.pack(pady=5)

    def add_log(self, msg):
        print(msg)
        timestamp = datetime.now().strftime("%H:%M:%S")
        entrada = f"[{timestamp}] {msg}"
        self.log_text.append(entrada)
        self.log_display.configure(text="\n".join(self.log_text[-10:]))

    def processar(self):
        try:
            self.log_text.clear()
            self.log_display.configure(text="")
            file_path = os.path.expanduser("C://Users//xql80316//Downloads//SIMULADOR_BALANÇA_3.0_1.xlsm")
            self.add_log("Abrindo planilha Excel...")

            xl = pd.ExcelFile(file_path)
            df_exp = xl.parse("dado_exp")
            df_sap = xl.parse("dado_sap")
            df_sku = xl.parse("dado_sku")
            df_sobrepeso = xl.parse("dado_sobrepeso")
            df_sobrepeso = df_sobrepeso[~df_sobrepeso['Data e hora'].astype(str).str.contains("Redimensionar", na=False)]
            df_sobrepeso['Data e hora'] = pd.to_datetime(df_sobrepeso['Data e hora'], errors='coerce')
            df_sobrepeso = df_sobrepeso.dropna(subset=['Data e hora'])


            peso_vazio = float(self.peso_vazio.get())
            qtd_paletes = int(self.qtd_paletes.get())

            self.add_log("Iniciando cálculo do peso final...")
            resultado = calcular_peso_final(
                self.remessa.get(), peso_vazio, qtd_paletes,
                df_exp, df_sku, df_sap, df_sobrepeso,
                self.add_log
            )

            if resultado:
                peso_base, sp, peso_paletes, peso_final = resultado
                wb = load_workbook(file_path)
                ws = wb["FORMULARIO"]
                ultima_linha = ws.max_row + 1
                ws.append([
                    datetime.now().strftime("%d/%m/%Y %H:%M:%S"), self.placa.get(), self.turno.get(),
                    self.remessa.get(), peso_vazio, peso_base, sp, peso_paletes, peso_final,
                    float(self.peso_balanca.get()), float(self.peso_balanca.get()) - peso_final
                ])
                wb.save(file_path)
                self.add_log("Resultado salvo na aba FORMULARIO.")
                messagebox.showinfo("Sucesso", "Resultado salvo com sucesso!")
            else:
                self.add_log("Falha no cálculo. Verifique os dados inseridos.")
                messagebox.showwarning("Aviso", "Cálculo não pôde ser realizado.")

        except Exception as e:
            self.add_log(f"Erro: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao processar: {str(e)}")

if __name__ == "__main__":
    app = App()
    app.mainloop()
