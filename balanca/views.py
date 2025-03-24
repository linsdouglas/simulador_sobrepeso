from django.shortcuts import render
import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))  # Caminho da pasta onde está o script

sap_file = os.path.join(BASE_DIR, 'data', 'dados_sap.xlsx')
sku_file = os.path.join(BASE_DIR, 'data', 'dados_sku.xlsx')
sobrepeso_file = os.path.join(BASE_DIR, 'data', 'dados_sobrepeso.xlsx')
expedicao_file = os.path.join(BASE_DIR, 'data', 'dados_expedicao.xlsx')

# Carregar os dados de cada planilha
df_sap = pd.read_excel(sap_file)
df_sku = pd.read_excel(sku_file)
df_sobrepeso = pd.read_excel(sobrepeso_file)
df_expedicao = pd.read_excel(expedicao_file)

# Função para calcular o peso final de uma remessa
def calcular_peso_final(remessa_num, peso_veiculo_vazio):
    try:
        remessa_num = int(remessa_num)
    except ValueError:
        return None

    # Filtrar a expedição pela remessa informada
    df_exp = df_expedicao[df_expedicao['REMESSA'] == remessa_num]
    if df_exp.empty:
        return None

    # Supomos que para cada remessa haja um SKU e uma quantidade de caixas (pode haver múltiplos registros, adapte se necessário)
    sku = df_exp.iloc[0]['ITEM']
    qtd_caixas = df_exp['QUANTIDADE'].sum()  # Somando se houver mais de um registro

    # 3. Buscar o peso por caixa do SKU na planilha de SKU (filtrando pela unidade "Caixa")
    df_sku_filtrado = df_sku[(df_sku['COD_PRODUTO'] == sku) & (df_sku['DESC_UNID_MEDID'] == 'Caixa')]
    if df_sku_filtrado.empty:
        print("SKU não encontrado na base de SKU ou unidade diferente de Caixa.")
        return None

    peso_por_caixa = df_sku_filtrado.iloc[0]['QTDE_PESO_LIQ']
    
    # Calcular o peso base (total das caixas)
    peso_base = qtd_caixas * peso_por_caixa

    # 4. Para sobrepeso, precisamos recuperar as chaves pallet associadas à remessa
    # Supondo que a expedição tenha uma coluna 'Chave Pallet'
    chaves_pallet = df_exp['CHAVE_PALETE'].unique()
    
    # Filtrar no SAP para obter as informações dos pallets
    df_pallets = df_sap[df_sap['Chave Pallet'].isin(chaves_pallet)]
    if df_pallets.empty:
        print("Nenhum pallet encontrado na base do SAP para a remessa.")
        return None

    # Suponha que cada pallet possui uma 'linha produzida' e uma 'data de produção'
    # Se houver mais de um pallet, podemos tirar uma média do sobrepeso dos pallets
    num_pallets = len(chaves_pallet)
    total_overweight_adjustment=0
    for _, row in df_pallets.iterrows():
        linha_produzida = row['LINHA PRODUZIDA']
        data_producao = row['Data de produção']  # Pode ser necessário ajustar o formato da data

        # 5. Filtrar a base de sobrepeso pela linha produzida e pela data (ou intervalo de data)
        # Aqui, por simplicidade, filtramos apenas pela linha produzida. Você pode incluir lógica para data.
        df_sobrepeso_filtrado = df_sobrepeso[
            (df_sobrepeso['Linhas'] == linha_produzida) &
            (df_sobrepeso['Dia']==data_producao)
        ]
        if not df_sobrepeso_filtrado.empty:
            # Supondo que a coluna com o sobrepeso médio se chame 'sobrepeso_medio'
            sobrepeso_medio = df_sobrepeso_filtrado.iloc[0]['Média de sobrepeso']
            overweight_decimal = sobrepeso_medio
            pallet_weight_share = peso_base / num_pallets
            overweight_adjustment = peso_base * overweight_decimal
            total_overweight_adjustment += overweight_adjustment
        else:
            print(f"Nenhum dado de sobrepeso encontrado para o pallet com data {data_producao.date()} e linha {linha_produzida}.")
    
    
    # 6. Calcular o peso final: peso base + ajuste de sobrepeso
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


