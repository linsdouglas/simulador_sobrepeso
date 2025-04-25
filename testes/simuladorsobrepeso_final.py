import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from datetime import datetime
import customtkinter as ctk
from tkinter import messagebox, StringVar
import os
import win32com.client as win32
import comtypes.client
from pathlib import Path
import winreg
from shutil import copyfile
import os
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import pandas as pd
import numpy as np
import os
import threading

def encontrar_pasta_onedrive_empresa():
    user_dir = os.environ["USERPROFILE"]
    possiveis = os.listdir(user_dir)
    for nome in possiveis:
        if "DIAS BRANCO" in nome.upper():
            caminho_completo = os.path.join(user_dir, nome)
            if os.path.isdir(caminho_completo) and "Gestão de Estoque - Documentos" in os.listdir(caminho_completo):
                return os.path.join(caminho_completo, "Gestão de Estoque - Documentos")
    return None

fonte_dir = encontrar_pasta_onedrive_empresa()
if not fonte_dir:
    raise FileNotFoundError("Não foi possível localizar a pasta sincronizada do SharePoint via OneDrive.")

MODELO_FORMULARIO = os.path.join(fonte_dir, "SIMULADOR_BALANÇA_LIMPO_2.xlsx")

def criar_copia_planilha(fonte_dir, nome_arquivo, log_callback):
    try:
        origem = os.path.join(fonte_dir, nome_arquivo)
        destino_dir = os.path.join(os.environ['USERPROFILE'], 'Downloads')
        timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
        nome_copia = f"copia_temp_{timestamp}_{nome_arquivo}"
        destino = os.path.join(destino_dir, nome_copia)
        copyfile(origem, destino)
        log_callback(f"Cópia criada com sucesso: {destino}")
        return destino
    except Exception as e:
        log_callback(f"Erro ao criar cópia da planilha: {e}")
        raise

def integrar_itens_detalhados(df_remessa, df_sap, df_sobrepeso_real, log_callback):
    itens_detalhados = []

    for _, row in df_remessa.iterrows():
        sku = row['ITEM']
        chave_pallet = row.get('CHAVE_PALETE', None)
        sp = 0.0

        if pd.notna(chave_pallet) and chave_pallet in df_sap['Chave Pallet'].values:
            try:
                lote_info = df_sap[df_sap['Chave Pallet'] == chave_pallet].iloc[0]
                lote = lote_info['Lote']
                data_producao = lote_info['Data de produção']
                hora_inicio = f"{lote_info['Hora de criação'].hour:02d}:00:00"
                hora_fim = f"{lote_info['Hora de modificação'].hour:02d}:00:00"
                linha_coluna = "LB" + lote[-2:]  # Ex: LB01

                df_sp_filtro = df_sobrepeso_real[
                    (df_sobrepeso_real['DataHora'] >= pd.to_datetime(f"{data_producao} {hora_inicio}")) &
                    (df_sobrepeso_real['DataHora'] <= pd.to_datetime(f"{data_producao} {hora_fim}"))
                ]

                if linha_coluna in df_sp_filtro.columns:
                    sp_valores = df_sp_filtro[linha_coluna].fillna(0)
                    if not sp_valores.empty:
                        media_sp = sp_valores.mean() / 100
                        log_callback(f"[{linha_coluna}] Média SP para SKU {sku}: {media_sp:.4f}")
                        sp = media_sp

            except Exception as e:
                log_callback(f"Erro ao calcular SP para pallet {chave_pallet}: {e}")

        itens_detalhados.append({'sku': sku, 'sp': round(sp, 4)})

    return itens_detalhados

def calcular_peso_final(remessa_num, peso_veiculo_vazio, qtd_paletes, df_exp, df_sku, df_sap, df_sobrepeso_real, log_callback):
    try:
        remessa_num = int(remessa_num)
    except ValueError:
        log_callback("Remessa inválida.")
        return None

    df_remessa = df_exp[df_exp['REMESSA'] == remessa_num]
    if df_remessa.empty:
        log_callback("Remessa não encontrada em data_exp.")
        return None

    skus = df_remessa['ITEM'].unique()
    qtd_caixas_total = df_remessa['QUANTIDADE'].sum()

    peso_base_total = 0
    peso_base_total_liq = 0
    sobrepesos_por_item = {}
    sp_total = 0

    for sku in skus:
        qtd_caixas = df_remessa[df_remessa['ITEM'] == sku]['QUANTIDADE'].sum()
        df_sku_filtrado = df_sku[df_sku['COD_PRODUTO'] == sku]
        if not df_sku_filtrado.empty:
            df_sku_filtrado = df_sku_filtrado.sort_values(by='DESC_UNID_MEDID')  
            unidade_usada = df_sku_filtrado.iloc[0]['DESC_UNID_MEDID']
        else:
            continue
        if df_sku_filtrado.empty:
            continue

        peso_por_caixa_bruto = df_sku_filtrado.iloc[0]['QTDE_PESO_BRU']
        peso_por_caixa_liquido = df_sku_filtrado.iloc[0]['QTDE_PESO_LIQ']
        peso_base_liq=qtd_caixas * peso_por_caixa_liquido
        peso_base = qtd_caixas * peso_por_caixa_bruto
        peso_base_total += peso_base
        peso_base_total_liq += peso_base_liq

        chaves_pallet = df_remessa[df_remessa['ITEM'] == sku]['CHAVE_PALETE'].unique()
        df_pallets = df_sap[df_sap['Chave Pallet'].isin(chaves_pallet)]

        total_overweight = 0
        count_sp = 0
        for idx, row in df_pallets.iterrows():
            try:
                log_callback(f"Processando pallet {idx+1}/{len(df_pallets)}...")

                lote = row['Lote']
                log_callback(f"Lote: {lote}")
                
                data_producao = row['Data de produção']
                log_callback(f"Data produção: {data_producao}")

                hora_inicio = f"{row['Hora de criação'].hour:02d}:00:00"
                hora_fim = f"{row['Hora de modificação'].hour:02d}:00:00"
                log_callback(f"Intervalo hora: {hora_inicio} - {hora_fim}")

                linha_coluna = "LB" + lote[-2:]
                log_callback(f"Linha produzida: {linha_coluna}")
                df_sp_filtro = df_sobrepeso_real[
                    (df_sobrepeso_real['DataHora'] >= pd.to_datetime(f"{data_producao} {hora_inicio}")) &
                    (df_sobrepeso_real['DataHora'] <= pd.to_datetime(f"{data_producao} {hora_fim}"))
                ]

                if linha_coluna in df_sp_filtro.columns:
                    sp_valores = df_sp_filtro[linha_coluna].fillna(0)
                    if not sp_valores.empty:
                        media_sp = sp_valores.mean() / 100

                log_callback(f"Linhas sobrepeso encontradas: {len(df_sp_filtro)}")  

                if linha_coluna in df_sp_filtro.columns:
                    sp_valores = df_sp_filtro[linha_coluna].fillna(0)
                    if not sp_valores.empty:
                        media_sp = sp_valores.mean() / 100
                        log_callback(f"[{linha_coluna}] Média SP: {media_sp:.4f}")
                        ajuste = peso_base_liq * media_sp
                        log_callback(f"Ajuste aplicado ao peso base ({peso_base_liq:.2f}): {ajuste:.2f}kg")
                        total_overweight += media_sp
                        count_sp += 1
                else:
                    log_callback(f"[AVISO] Coluna {linha_coluna} não encontrada na base de sobrepeso.")


            except Exception as e:
                log_callback(f"Erro ao processar pallet {idx+1}: {e}")


        if count_sp > 0:
            sp_medio = total_overweight / count_sp
            sobrepesos_por_item[sku] = sp_medio
            sp_total_ajuste = peso_base_liq * sp_medio
            log_callback(f"Total SP médio do SKU {sku}: {sp_medio:.4f}, ajuste total: {sp_total_ajuste:.2f}kg")
            sp_total += sp_total_ajuste

    peso_com_sobrepeso = peso_base_total + sp_total
    log_callback(f"Peso com sobrepeso: {peso_com_sobrepeso:.2f} kg")
    peso_total_com_paletes = peso_com_sobrepeso + (qtd_paletes * 26) + peso_veiculo_vazio
    log_callback(f"Peso total com paletes ({qtd_paletes} x 26kg): {peso_total_com_paletes:.2f} kg")
    if sobrepesos_por_item:
        media_sp_geral = sum(sobrepesos_por_item.values()) / len(sobrepesos_por_item)
    else:
        media_sp_geral = 0.0
    log_callback(f"Média geral de sobrepeso (entre {len(sobrepesos_por_item)} itens): {media_sp_geral:.4f}")
    
    itens_detalhados = integrar_itens_detalhados(df_remessa, df_sap, df_sobrepeso_real, log_callback)

    return peso_base_total, sp_total, peso_com_sobrepeso, peso_total_com_paletes, media_sp_geral, itens_detalhados

def preencher_formulario_com_openpyxl(path_copia, dados, sobrepesos_por_item, log_callback):
    try:
        wb = load_workbook(path_copia)
        ws = wb["FORMULARIO"]

        log_callback("Preenchendo cabeçalhos principais com openpyxl...")
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
        log_callback("Preenchendo SKUs e sobrepesos...")
        for item in sobrepesos_por_item: 
            ws[f"C{linha}"] = item['sku']
            ws[f"D{linha}"] = item['sp']
            linha += 1

        wb.save(path_copia)
        log_callback("Formulário preenchido e salvo com sucesso.")

    except Exception as e:
        log_callback(f"Erro no preenchimento: {e}")
        raise

def exportar_pdf_com_comtypes(path_xlsx, aba_nome="FORMULARIO", nome_remessa="REMESSA", log_callback=None):

    try:
        if log_callback:
            log_callback("Iniciando exportação via comtypes...")

        comtypes.CoInitialize()
        excel = comtypes.client.CreateObject("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(path_xlsx)

        if aba_nome not in [sheet.Name for sheet in wb.Sheets]:
            raise Exception(f"Aba '{aba_nome}' não encontrada.")

        ws = wb.Worksheets(aba_nome)

        pdf_dir = "C:\\temp"
        os.makedirs(pdf_dir, exist_ok=True)

        pdf_path = os.path.join(pdf_dir, f"SOBREPESOSIMULADO - {nome_remessa}.pdf")

        if log_callback:
            log_callback(f"Tentando exportar para: {pdf_path}")

        ws.ExportAsFixedFormat(Type=0, Filename=pdf_path)
        wb.Close(False)
        excel.Quit()
        comtypes.CoUninitialize()

        if log_callback:
            log_callback(f"PDF exportado com sucesso: {pdf_path}")

        return pdf_path

    except Exception as e:
        if log_callback:
            log_callback(f"Erro ao exportar PDF: {e}")
        raise

def gerar_relatorio_diferenca(remessa_num, peso_final_balança, peso_veiculo_vazio, df_remessa, df_sku, peso_estimado_total, pasta_excel):
    import os
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_pdf import PdfPages
    import pandas as pd

    # Criar pasta de destino
    pasta_destino = os.path.join(pasta_excel, 'Analise_divergencia')
    os.makedirs(pasta_destino, exist_ok=True)

    # Preparar dados da remessa
    skus = df_remessa['ITEM'].unique()
    dados_relatorio = []
    peso_base_total_liq = 0

    for sku in skus:
        qtd = df_remessa[df_remessa['ITEM'] == sku]['QUANTIDADE'].sum()
        df_sku_filtrado = df_sku[df_sku['COD_PRODUTO'] == sku]

        if df_sku_filtrado.empty:
            continue

        unidade = df_sku_filtrado.iloc[0]['DESC_UNID_MEDID']
        peso_unit_liq = df_sku_filtrado.iloc[0]['QTDE_PESO_LIQ']
        peso_total_liq = qtd * peso_unit_liq
        peso_base_total_liq += peso_total_liq

        dados_relatorio.append({
            'SKU': sku,
            'Unidade': unidade,
            'Quantidade': qtd,
            'Peso Total Líquido': peso_total_liq,
            'Peso Unit. Líquido': peso_unit_liq
        })

    df_dados = pd.DataFrame(dados_relatorio)

    # Cálculo da diferença
    diferenca_total = diferenca_total = (peso_estimado_total + peso_veiculo_vazio) - peso_final_balança
    peso_carga_real = peso_final_balança - peso_veiculo_vazio
    df_dados['% Peso'] = df_dados['Peso Total Líquido'] / peso_base_total_liq
    df_dados['Peso Proporcional Real'] = df_dados['% Peso'] * peso_carga_real
    df_dados['Quantidade Real Estimada'] = (df_dados['Peso Proporcional Real'] / df_dados['Peso Unit. Líquido']).round()
    df_dados['Diferença Estimada (kg)'] = df_dados['% Peso'] * diferenca_total
    df_dados['Unid. Estimada de Divergência'] = df_dados['Quantidade Real Estimada'] - df_dados['Quantidade']

    # Nome do PDF
    nome_pdf = f"Análise quantitativa - {remessa_num}.pdf"
    caminho_pdf = os.path.join(pasta_destino, nome_pdf)

    # Criar gráfico de quantidades esperadas vs ajustadas
    fig, ax = plt.subplots(figsize=(10, 6))
    largura_barra = 0.35
    x = range(len(df_dados))

    qtd_esperada = df_dados['Quantidade']
    qtd_real = df_dados['Quantidade Real Estimada']

    ax.bar([i - largura_barra/2 for i in x], qtd_esperada, width=largura_barra, label='Quantidade Esperada')
    ax.bar([i + largura_barra/2 for i in x], qtd_real, width=largura_barra, label='Quantidade Real Estimada')

    ax.set_ylabel("Quantidade (unidades)")
    ax.set_xlabel("SKU")
    ax.set_title("Comparativo: Quantidade Esperada vs Real Estimada por SKU")
    ax.set_xticks(x)
    ax.set_xticklabels(df_dados['SKU'].astype(str))
    ax.axhline(0, color='gray', linestyle='--')
    ax.legend()

    # Adicionar texto nas barras
    for i in x:
        ax.text(i - largura_barra/2, qtd_esperada.iloc[i], f"{qtd_esperada.iloc[i]:.0f}", ha='center', va='bottom')
        ax.text(i + largura_barra/2, qtd_real.iloc[i], f"{qtd_real.iloc[i]:.0f}", ha='center', va='bottom')

    # Gerar PDF com tabela + gráfico
    with PdfPages(caminho_pdf) as pdf:
        # Página 1 – Tabela
        fig_tabela, ax_tabela = plt.subplots(figsize=(12, len(df_dados) * 0.5 + 3))
        ax_tabela.axis('off')
        table_data = [
            ['SKU', 'Unidade', 'Qtd. Enviada', 'Qtd. Real Estimada', 'Peso Total Líquido', '% do Peso', 'Diferença (kg)', 'Divergência (unid)']
        ] + df_dados[['SKU', 'Unidade', 'Quantidade', 'Quantidade Real Estimada', 'Peso Total Líquido', '% Peso', 'Diferença Estimada (kg)', 'Unid. Estimada de Divergência']].round(2).values.tolist()

        tabela = ax_tabela.table(cellText=table_data, colLabels=None, loc='center', cellLoc='center')
        tabela.auto_set_font_size(False)
        tabela.set_fontsize(10)
        tabela.scale(1, 1.5)

        titulo = f"Relatório Comparativo - Remessa {remessa_num}\nPeso estimado: {peso_estimado_total:.2f} kg | Peso balança: {peso_final_balança:.2f} kg | Peso veículo: {peso_veiculo_vazio:.2f} kg | Diferença: {diferenca_total:.2f} kg"
        fig_tabela.suptitle(titulo, fontsize=12)
        pdf.savefig(fig_tabela, bbox_inches='tight')

        # Página 2 – Gráfico
        pdf.savefig(fig, bbox_inches='tight')

    plt.close('all')
    print(f"Relatório salvo em: {caminho_pdf}")
    return caminho_pdf

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")
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

        ctk.CTkButton(self, text="Calcular", command=self.iniciar_processamento).pack(pady=10)

        ctk.CTkLabel(self, text="Histórico de Logs:").pack(pady=5)
        self.log_display = ctk.CTkLabel(self, text="", wraplength=600, justify="left")
        self.log_display.pack(pady=5)

    def add_log(self, msg):
        print(msg)
        timestamp = datetime.now().strftime("%H:%M:%S")
        entrada = f"[{timestamp}] {msg}"
        self.log_text.append(entrada)
        self.log_display.configure(text="\n".join(self.log_text[-10:]))
    
    def iniciar_processamento(self):
        thread = threading.Thread(target=self.processar)
        thread.start()

    def processar(self):
        path_base_sobrepeso = os.path.join(fonte_dir, "Base_sobrepeso_real.xlsx")
        df_sobrepeso_real = pd.read_excel(path_base_sobrepeso, sheet_name="SOBREPESO")
        df_sobrepeso_real['DataHora'] = pd.to_datetime(df_sobrepeso_real['DataHora'])

        try:
            self.log_text.clear()
            self.log_display.configure(text="")

            file_path = criar_copia_planilha(fonte_dir, "SIMULADOR_BALANÇA_LIMPO_2.xlsx", self.add_log)
            self.add_log(f"Abrindo planilha Excel em: {file_path}")

            xl = pd.ExcelFile(file_path)
            self.add_log("Lendo abas do arquivo...")
            df_exp = xl.parse("dado_exp")
            df_sap = xl.parse("dado_sap")
            df_sku = xl.parse("dado_sku")
            self.add_log("Abas carregadas com sucesso.")

            peso_vazio = float(self.peso_vazio.get())
            peso_balança = float(self.peso_balanca.get())
            qtd_paletes = int(self.qtd_paletes.get())
            remessa = int(self.remessa.get())

            self.add_log(f"Entradas: Remessa={remessa}, Peso Vazio={peso_vazio}, Paletes={qtd_paletes}")

            self.add_log("Iniciando cálculo do peso final...")
            resultado = calcular_peso_final(
                remessa, peso_vazio, qtd_paletes,
                df_exp, df_sku, df_sap, df_sobrepeso_real,
                self.add_log
            )

            if resultado:
                peso_base, sp_total, peso_com_sp, peso_final, media_sp, itens_detalhados = resultado

                dados = {
                    'remessa': remessa,
                    'qtd_skus': df_exp[df_exp['REMESSA'] == remessa]['ITEM'].nunique(),
                    'placa': self.placa.get(),
                    'turno': self.turno.get(),
                    'peso_vazio': peso_vazio,
                    'peso_base': peso_base,
                    'sp_total': sp_total,
                    'peso_com_sp': peso_com_sp,
                    'peso_total_final': peso_final,
                    'media_sp': media_sp,
                    'qtd_paletes': qtd_paletes
                }

                self.add_log("Chamando preenchimento do formulário via COM...")
                preencher_formulario_com_openpyxl(file_path, dados, itens_detalhados, self.add_log)

                self.add_log("Exportando PDF...")
                pdf_path = exportar_pdf_com_comtypes(file_path, "FORMULARIO", nome_remessa=remessa, log_callback=self.add_log)
                self.add_log(f"PDF exportado com sucesso: {pdf_path}")

                self.add_log("Gerando relatório de divergência em PDF...")
                relatorio_path = gerar_relatorio_diferenca(
                    remessa_num=remessa,
                    peso_final_balança=peso_balança,
                    peso_veiculo_vazio=peso_vazio,
                    df_remessa=df_exp[df_exp['REMESSA'] == remessa],
                    df_sku=df_sku,
                    peso_estimado_total=peso_com_sp,
                    pasta_excel=fonte_dir
                )
                self.add_log(f"Relatório adicional salvo em: {relatorio_path}")

                messagebox.showinfo(
                    "Sucesso",
                    f"Formulário exportado: {pdf_path}\n\nRelatório de divergência salvo:\n{relatorio_path}"
                )

            else:
                self.add_log("Falha no cálculo. Verifique os dados inseridos.")
                messagebox.showwarning("Aviso", "Cálculo não pôde ser realizado.")

        except Exception as e:
            self.add_log(f"Erro: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao processar: {str(e)}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
