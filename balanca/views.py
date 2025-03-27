from django.shortcuts import render
import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sap_file = os.path.join(BASE_DIR, 'data', 'dados_sap.xlsx')
sku_file = os.path.join(BASE_DIR, 'data', 'dados_sku.xlsx')
sobrepeso_file = os.path.join(BASE_DIR, 'data', 'dados_sobrepeso.xlsx')
expedicao_file = os.path.join(BASE_DIR, 'data', 'dados_expedicao.xlsx')

df_sap = pd.read_excel(sap_file)
df_sku = pd.read_excel(sku_file)
df_sobrepeso = pd.read_excel(sobrepeso_file)
df_expedicao = pd.read_excel(expedicao_file)

import pandas as pd

def calcular_peso_final(remessa_num, peso_veiculo_vazio):
    try:
        remessa_num = int(remessa_num)
    except ValueError:
        return None

    df_exp = df_expedicao[df_expedicao['REMESSA'] == remessa_num]
    if df_exp.empty:
        return None

    sku = df_exp.iloc[0]['ITEM']
    qtd_caixas = df_exp['QUANTIDADE'].sum()  

    df_sku_filtrado = df_sku[(df_sku['COD_PRODUTO'] == sku) & (df_sku['DESC_UNID_MEDID'] == 'Caixa')]
    if df_sku_filtrado.empty:
        print("SKU não encontrado na base de SKU ou unidade diferente de Caixa.")
        return None

    peso_por_caixa = df_sku_filtrado.iloc[0]['QTDE_PESO_LIQ']
    peso_base = qtd_caixas * peso_por_caixa
    chaves_pallet = df_exp['CHAVE_PALETE'].unique()
    
    df_pallets = df_sap[df_sap['Chave Pallet'].isin(chaves_pallet)]
    if df_pallets.empty:
        print("Nenhum pallet encontrado na base do SAP para a remessa.")
        return None

    num_pallets = len(chaves_pallet)
    total_overweight_adjustment=0
    for _, row in df_pallets.iterrows():
        lote = row['Lote']
        data_producao = row['Data de produção']
        last3=lote[-3:]
        linha_produzida = "L" + last3
        df_sobrepeso_filtrado = df_sobrepeso[
            (df_sobrepeso['Linhas'] == linha_produzida) &
            (df_sobrepeso['Dia']==data_producao)
        ]
        if not df_sobrepeso_filtrado.empty:
            sobrepeso_medio = df_sobrepeso_filtrado.iloc[0]['Média de sobrepeso']
            overweight_decimal = sobrepeso_medio
            pallet_weight_share = peso_base / num_pallets
            overweight_adjustment = peso_base * overweight_decimal
            total_overweight_adjustment += overweight_adjustment
        else:
            print(f"Nenhum dado de sobrepeso encontrado para o pallet com data {data_producao.date()} e linha {linha_produzida}.")
    
    
    peso_final = (peso_veiculo_vazio)+(peso_base + overweight_adjustment)

    return {
        'remessa': remessa_num,
        'peso_veiculo_vazio': peso_veiculo_vazio,
        'qtd_caixas': qtd_caixas,
        'peso_por_caixa': peso_por_caixa,
        'peso_base': peso_base,
        'total_overweight_adjustment': overweight_adjustment,
        'peso_final': peso_final,
        'peso pallete':pallet_weight_share,
    }

def analise_ocorrencias(request):
    contexto = {}
    if request.method == 'POST':
        remessa = request.POST.get('remessa')
        try:
            peso_vazio = float(request.POST.get('peso_vazio'))
        except (ValueError, TypeError):
            peso_vazio = 0

        resultado = calcular_peso_final(remessa, peso_vazio)
        if resultado:
            contexto['resultado'] = resultado
        else:
            contexto['erro'] = "Não foi possível calcular o peso final para a remessa informada."
    return render(request, 'balanca/formulario_remessa.html', contexto)


