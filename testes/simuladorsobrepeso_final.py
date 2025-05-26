import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from datetime import datetime
import customtkinter as ctk
from tkinter import messagebox, StringVar
import win32com.client as win32
import comtypes.client
from pathlib import Path
import winreg
from shutil import copyfile
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import numpy as np
import os
import threading
import glob
import subprocess

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

caminho_base_fisica = os.path.join(fonte_dir, "SIMULADOR_BALANÇA_LIMPO_2.xlsx")
df_base_fisica = pd.read_excel(caminho_base_fisica, sheet_name="BASE FISICA")
df_base_familia = pd.read_excel(caminho_base_fisica, 'BASE_FAMILIA')

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

def print_pdf(file_path, impressora="VITLOG01A01", sumatra_path="C:\\Program Files\\SumatraPDF\\SumatraPDF.exe"):
    args = [sumatra_path, "-print-to", impressora, "-silent", file_path]
    try:
        result = subprocess.run(args, check=True, capture_output=True, text=True)
        print(f"Arquivo impresso com sucesso: {file_path}")
        if result.stdout:
            print(f"stdout: {result.stdout}")
        if result.stderr:
            print(f"stderr: {result.stderr}")
    except subprocess.CalledProcessError as e:
        print(f"Erro ao imprimir {file_path}: {e}")
        print(f"Output: {e.output}")
        print(f"Stderr: {e.stderr}")

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
                linha_coluna = "L" + lote[-3:] if isinstance(lote, str) and len(lote) >= 2 else "LB00"
                if linha_coluna in ['LB06', 'LB07']:
                    linha_coluna='LB06/07'
                log_callback(f"Linha produzida ajustada: {linha_coluna}")
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
                        itens_detalhados.append({'sku': sku, 'sp': round(sp, 4), 'origem': 'real'})
                        continue

            except Exception as e:
                log_callback(f"Erro ao calcular SP para pallet {chave_pallet}: {e}")
        else:
            sp_valor, _ = calculo_sobrepeso_fixo(sku, df_base_fisica, 0, log_callback)
            itens_detalhados.append({'sku': sku, 'sp': round(sp_valor, 4)})
            continue

    return itens_detalhados

def calculo_sobrepeso_fixo(sku, df_base_fisica, peso_base_liq, log_callback):
    try:
        sp_row = df_base_fisica[df_base_fisica['CÓDIGO PRODUTO'] == sku]
        if not sp_row.empty:
            sp_fixo = sp_row.iloc[0]['SOBRE PESO'] / 100
            ajuste = peso_base_liq * sp_fixo
            log_callback(f"SOBREPESO FIXO encontrado para SKU {sku}: {sp_fixo:.4f}")
            return sp_fixo, ajuste
        else:
            log_callback(f"Nenhum sobrepeso fixo encontrado para SKU {sku}.")
            return 0, 0
    except Exception as e:
        log_callback(f"Erro ao buscar sobrepeso fixo para SKU {sku}: {e}")
        return 0, 0

def calcular_peso_final(remessa_num, peso_veiculo_vazio, qtd_paletes,
                        df_expedicao, df_sku, df_sap, df_sobrepeso_real,
                        df_base_fisica, df_frac, log_callback):
    try:
        remessa_num = int(remessa_num)
    except ValueError:
        log_callback("Remessa inválida.")
        return None

    df_remessa = df_expedicao[df_expedicao['REMESSA'] == remessa_num]
    if df_remessa.empty:
        log_callback("Remessa não encontrada em data_exp.")
        return None

    peso_base_total = 0
    peso_base_total_liq = 0
    sp_total = 0
    itens_detalhados = []

    for idx, row in df_remessa.iterrows():
        sku = row['ITEM']
        chave_pallet_atual = row['CHAVE_PALETE']
        qtd_caixas = row['QUANTIDADE']

        if pd.isna(chave_pallet_atual):
            df_frac_remessa = df_frac[(df_frac['REMESSA'] == remessa_num) & (df_frac['ITEM'] == sku)]
            if not df_frac_remessa.empty:
                chave_pallet_atual = df_frac_remessa.iloc[0]['CHAVE_PALETE']
                qtd_caixas = df_frac_remessa.iloc[0]['QUANTIDADE'] 
                log_callback(f"Chave pallet e quantidade encontradas na base FRACAO: {chave_pallet_atual} → QTD: {qtd_caixas} para SKU {sku}")

        df_sku_filtrado = df_sku[df_sku['COD_PRODUTO'] == sku]
        if df_sku_filtrado.empty:
            log_callback(f"SKU {sku} não encontrado na base SKU.")
            continue

        peso_por_caixa_bruto = df_sku_filtrado.iloc[0]['QTDE_PESO_BRU']
        peso_por_caixa_liquido = df_sku_filtrado.iloc[0]['QTDE_PESO_LIQ']
        peso_base = qtd_caixas * peso_por_caixa_bruto
        peso_base_liq = qtd_caixas * peso_por_caixa_liquido
        peso_base_total += peso_base
        peso_base_total_liq += peso_base_liq
        sp = 0
        origem_sp = 'não encontrado'
        ajuste_sp = 0

        if pd.notna(chave_pallet_atual) and chave_pallet_atual in df_sap['Chave Pallet'].values:
            pallet_info = df_sap[df_sap['Chave Pallet'] == chave_pallet_atual].iloc[0]
            lote = pallet_info['Lote']
            data_producao = pallet_info['Data de produção']
            hora_inicio = f"{pallet_info['Hora de criação'].hour:02d}:00:00"
            hora_fim = f"{pallet_info['Hora de modificação'].hour:02d}:00:00"
            linha_coluna = "L" + lote[-3:]
            if linha_coluna in ['LB06', 'LB07']:
                linha_coluna = 'LB06/07'
            log_callback(f"Processando pallet {chave_pallet_atual} para SKU {sku}.")
            df_sp_filtro = df_sobrepeso_real[
                (df_sobrepeso_real['DataHora'] >= pd.to_datetime(f"{data_producao} {hora_inicio}")) &
                (df_sobrepeso_real['DataHora'] <= pd.to_datetime(f"{data_producao} {hora_fim}"))
            ]
            if linha_coluna in df_sp_filtro.columns:
                sp_valores = df_sp_filtro[linha_coluna].fillna(0)
                if not sp_valores.empty:
                    media_sp = sp_valores.mean() / 100
                    log_callback(f"[{linha_coluna}] Média SP: {media_sp:.4f}")

                    if media_sp != 0.0:
                        sp = media_sp
                        origem_sp = 'real'
                        ajuste_sp = peso_base_liq * sp
                    else:
                        log_callback(f"Média SP é 0 para SKU {sku}. Indo buscar sobrepeso fixo...")
            else:
                log_callback(f"Coluna {linha_coluna} não encontrada na base de sobrepeso.")

        if sp == 0:
            sp_valor, ajuste_fixo = calculo_sobrepeso_fixo(sku, df_base_fisica, peso_base_liq, log_callback)
            if sp_valor != 0:
                sp = sp_valor
                origem_sp = 'fixo'
                ajuste_sp = ajuste_fixo
                log_callback(f"SOBREPESO FIXO aplicado para SKU {sku}: {sp:.4f} → Ajuste: {ajuste_sp:.2f} kg")
            else:
                log_callback(f"Nenhum sobrepeso encontrado para SKU {sku}.")
        sp_total += ajuste_sp
        itens_detalhados.append({
            'sku': sku,
            'chave_pallet': chave_pallet_atual,
            'sp': round(sp, 4),
            'ajuste_sp': round(ajuste_sp, 2),
            'origem': origem_sp
        })
    peso_com_sobrepeso = peso_base_total + sp_total
    log_callback(f"Peso com sobrepeso: {peso_com_sobrepeso:.2f} kg")
    peso_total_com_paletes = peso_com_sobrepeso + (qtd_paletes * 26) + peso_veiculo_vazio
    log_callback(f"Peso total com paletes ({qtd_paletes} x 26kg): {peso_total_com_paletes:.2f} kg")
    media_sp_geral = (sum(item['sp'] for item in itens_detalhados) / len(itens_detalhados)) if itens_detalhados else 0.0
    log_callback(f"Média geral de sobrepeso (entre {len(itens_detalhados)} itens): {media_sp_geral:.4f}")
    return peso_base_total, sp_total, peso_com_sobrepeso, peso_total_com_paletes, media_sp_geral, itens_detalhados

def calcular_limites_sobrepeso_por_quantidade(dados, itens_detalhados, df_base_familia, df_sobrepeso_tabela, df_sku, df_remessa, log_callback):
    total_quantidade = 0
    quantidade_com_sp_real = 0
    ponderador_pos = 0
    ponderador_neg = 0

    for item in itens_detalhados:
        sku = item['sku']
        sp = item.get('sp', 0)
        origem = item.get('origem', 'fixo')

        qtd = df_remessa[(df_remessa['REMESSA'] == dados['remessa']) & (df_remessa['ITEM'] == sku)]['QUANTIDADE'].sum()
        total_quantidade += qtd

        if origem == 'real':
            quantidade_com_sp_real += qtd
            if sp > 0:
                ponderador_pos += sp * qtd
            elif sp < 0:
                ponderador_neg += abs(sp) * qtd

    if total_quantidade == 0:
        proporcao_sp_real = 0
    else:
        proporcao_sp_real = quantidade_com_sp_real / total_quantidade

    log_callback(f"Total de quantidade: {total_quantidade}, com SP Real: {quantidade_com_sp_real}, proporção: {proporcao_sp_real:.2%}")

    if proporcao_sp_real >= 0.5:
        if quantidade_com_sp_real > 0:
            media_positiva = ponderador_pos / quantidade_com_sp_real if ponderador_pos > 0 else 0.02
            media_negativa = ponderador_neg / quantidade_com_sp_real if ponderador_neg > 0 else 0.01
        else:
            media_positiva = 0.02
            media_negativa = 0.01

        log_callback("Mais de 50% da quantidade com SP Real. Usando médias ponderadas:")
        log_callback(f"Sobrepeso para mais (real): {media_positiva:.4f}")
        log_callback(f"Sobrepeso para menos (real): {media_negativa:.4f}")
    else:
        log_callback("Menos de 50% da quantidade com SP Real. Usando tabela de sobrepeso físico.")
        familias = set()
        for item in itens_detalhados:
            sku = item['sku']
            familia = df_base_familia.loc[df_base_familia['CÓD'] == sku, 'FAMILIA 2']
            if not familia.empty:
                familias.add(familia.values[0])

        if len(familias) == 1:
            familia = list(familias)[0]
            if 'BISCOITO' in familia.upper():
                row = df_sobrepeso_tabela.loc[df_sobrepeso_tabela.index.str.contains("BISCOITO", case=False)]
            elif 'MASSA' in familia.upper():
                row = df_sobrepeso_tabela.loc[df_sobrepeso_tabela.index.str.contains("MASSA", case=False)]
            else:
                row = df_sobrepeso_tabela.loc[df_sobrepeso_tabela.index.str.contains("MIX", case=False)]
        else:
            row = df_sobrepeso_tabela.loc[df_sobrepeso_tabela.index.str.contains("MIX", case=False)]

        media_positiva = row['(+)'].values[0]
        media_negativa = row['(-)'].values[0]

        log_callback(f"Sobrepeso para mais (físico): {media_positiva:.4f}")
        log_callback(f"Sobrepeso para menos (físico): {media_negativa:.4f}")

    return media_positiva, media_negativa, proporcao_sp_real

def preencher_formulario_com_openpyxl(path_copia, dados, itens_detalhados, log_callback,df_sku, df_remessa):
    try:
        dados_tabela = {
        '(+)': [0.02, 0.005, 0.04],
        '(-)': [0.01, 0.01, 0.01]
        }
        index = ['CARGA COM MIX', 'EXCLUSIVO MASSAS', 'EXCLUSIVO BISCOITOS']
        df_sobrepeso_tabela = pd.DataFrame(dados_tabela, index=index)
        df_base_familia = pd.read_excel(caminho_base_fisica, 'BASE_FAMILIA')
        sp_pos, sp_neg, proporcao_sp_real = calcular_limites_sobrepeso_por_quantidade(
            dados, itens_detalhados, df_base_familia, df_sobrepeso_tabela, df_sku, df_remessa, log_callback
        )
        wb = load_workbook(path_copia)
        ws = wb["FORMULARIO"]

        log_callback("Preenchendo cabeçalhos principais com openpyxl...")
        ws["A16"] = f"Sobrepeso para (+): {sp_pos*100:.2f}%"
        ws["A18"] = f"Sobrepeso para (-): {sp_neg*100:.2f}%"
        ws["D7"] = f"{proporcao_sp_real*100:.2f}% x {(1 - proporcao_sp_real)*100:.2f}%"
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
        ws["B16"] = dados['peso_total_final'] * (1 + sp_pos)
        ws["B17"] = dados['peso_total_final']
        ws["B18"] = dados['peso_total_final'] * (1 - sp_neg)
        ws["D4"] = dados['qtd_paletes']
        ws["D9"] = dados['qtd_paletes'] * 26

        linha_inicio = 12
        linha_fim = 46
        max_itens = linha_fim - linha_inicio + 1

        log_callback("Preenchendo SKUs e sobrepesos...")
        itens_real = [item for item in itens_detalhados if item['origem'] == 'real']
        itens_fixo = [item for item in itens_detalhados if item['origem'] == 'fixo']
        itens_nao_encontrado = [item for item in itens_detalhados if item['origem'] == 'não encontrado']

        itens_ordenados = itens_real + itens_fixo + itens_nao_encontrado
        for idx, item in enumerate(itens_ordenados):
            if idx >= max_itens:
                log_callback("Limite máximo de linhas atingido no formulário (C12 até C46). Restante será desconsiderado.")
                break
            linha = linha_inicio + idx
            sku_texto = f"{item['sku']} ({item['origem']})"
            ws[f"C{linha}"] = sku_texto
            ws[f"D{linha}"] = f"{item['sp']*100:.3f}"


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

        pdf_dir = os.path.join(fonte_dir, 'Relatório_Saida')
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

    pasta_destino = os.path.join(pasta_excel, 'Analise_divergencia')
    os.makedirs(pasta_destino, exist_ok=True)

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

    diferenca_total = diferenca_total = (peso_estimado_total + peso_veiculo_vazio) - peso_final_balança
    peso_carga_real = peso_final_balança - peso_veiculo_vazio
    df_dados['% Peso'] = df_dados['Peso Total Líquido'] / peso_base_total_liq
    df_dados['Peso Proporcional Real'] = df_dados['% Peso'] * peso_carga_real
    df_dados['Quantidade Real Estimada'] = (df_dados['Peso Proporcional Real'] / df_dados['Peso Unit. Líquido']).round()
    df_dados['Diferença Estimada (kg)'] = df_dados['% Peso'] * diferenca_total
    df_dados['Unid. Estimada de Divergência'] = df_dados['Quantidade Real Estimada'] - df_dados['Quantidade']

    nome_pdf = f"Análise quantitativa - {remessa_num}.pdf"
    caminho_pdf = os.path.join(pasta_destino, nome_pdf)

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

    for i in x:
        ax.text(i - largura_barra/2, qtd_esperada.iloc[i], f"{qtd_esperada.iloc[i]:.0f}", ha='center', va='bottom')
        ax.text(i + largura_barra/2, qtd_real.iloc[i], f"{qtd_real.iloc[i]:.0f}", ha='center', va='bottom')

    with PdfPages(caminho_pdf) as pdf:
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
        ctk.CTkComboBox(self, values=["A", "B", "C"], variable=self.turno).pack()

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
        path_base_expedicao = os.path.join(fonte_dir, "expedicao.xlsx")
        path_base_sap = os.path.join(fonte_dir, "base_sap.xlsx")
        path_base_frac = os.path.join(fonte_dir, "FRACAO.xlsx")
        df_frac=pd.read_excel(path_base_frac, sheet_name="Sheet1")
        df_sap = pd.read_excel(path_base_sap, sheet_name="Sheet1")
        df_expedicao = pd.read_excel(path_base_expedicao, sheet_name="dado_exp")
        df_sobrepeso_real = pd.read_excel(path_base_sobrepeso, sheet_name="SOBREPESO")
        df_sobrepeso_real['DataHora'] = pd.to_datetime(df_sobrepeso_real['DataHora'])
        df_base_fisica = pd.read_excel(caminho_base_fisica, sheet_name="BASE FISICA")
        try:
            self.log_text.clear()
            self.log_display.configure(text="")

            file_path = criar_copia_planilha(fonte_dir, "SIMULADOR_BALANÇA_LIMPO_2.xlsx", self.add_log)
            self.add_log(f"Abrindo planilha Excel em: {file_path}")

            xl = pd.ExcelFile(file_path)
            self.add_log("Lendo abas do arquivos...")
            df_sku = xl.parse("dado_sku")
            self.add_log("Abas carregadas com sucesso.")

            peso_vazio = float(self.peso_vazio.get())
            peso_balança = float(self.peso_balanca.get())
            qtd_paletes = int(self.qtd_paletes.get())
            remessa = int(self.remessa.get())
            df_remessa = df_expedicao[df_expedicao['REMESSA'] == remessa]

            self.add_log(f"Entradas: Remessa={remessa}, Peso Vazio={peso_vazio}, Paletes={qtd_paletes}")

            self.add_log("Iniciando cálculo do peso final...")
            resultado = calcular_peso_final(
                remessa, peso_vazio, qtd_paletes,
                df_expedicao, df_sku, df_sap, df_sobrepeso_real,df_base_fisica, df_frac,
                self.add_log
            )

            if resultado:
                peso_base, sp_total, peso_com_sp, peso_final, media_sp, itens_detalhados = resultado

                dados = {
                    'remessa': remessa,
                    'qtd_skus': df_expedicao[df_expedicao['REMESSA'] == remessa]['ITEM'].nunique(),
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
                preencher_formulario_com_openpyxl(file_path, dados, itens_detalhados, self.add_log, df_sku, df_remessa)


                self.add_log("Exportando PDF...")
                pdf_path = exportar_pdf_com_comtypes(file_path, "FORMULARIO", nome_remessa=remessa, log_callback=self.add_log)
                self.add_log(f"PDF exportado com sucesso: {pdf_path}")

                self.add_log("Gerando relatório de divergência em PDF...")
                relatorio_path = gerar_relatorio_diferenca(
                    remessa_num=remessa,
                    peso_final_balança=peso_balança,
                    peso_veiculo_vazio=peso_vazio,
                    df_remessa=df_expedicao[df_expedicao['REMESSA'] == remessa],
                    df_sku=df_sku,
                    peso_estimado_total=peso_com_sp,
                    pasta_excel=fonte_dir
                )
                self.add_log(f"Relatório adicional salvo em: {relatorio_path}")

                messagebox.showinfo(
                    "Sucesso",
                    f"Formulário exportado: {pdf_path}\n\nRelatório de divergência salvo:\n{relatorio_path}"
                )
                print_pdf(pdf_path)  
                print_pdf(relatorio_path)

                try:
                    os.remove(file_path)
                    self.add_log(f"Cópia temporária removida: {file_path}")
                except Exception as e:
                    self.add_log(f"Erro ao remover a cópia temporária: {e}")

            else:
                self.add_log("Falha no cálculo. Verifique os dados inseridos.")
                messagebox.showwarning("Aviso", "Cálculo não pôde ser realizado.")


        except Exception as e:
            self.add_log(f"Erro: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao processar: {str(e)}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
