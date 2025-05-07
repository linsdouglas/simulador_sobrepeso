import os
import pandas as pd
from datetime import datetime
import win32com.client as win32
from pandas.errors import OutOfBoundsDatetime
import win32com.client.gencache
import shutil
import os
folder = win32com.client.gencache.GetGeneratePath()
shutil.rmtree(folder, ignore_errors=True)
win32com.client.gencache.EnsureDispatch("Excel.Application")


print("Script iniciado.")

dados_raw = r"C:\Users\xql80316\Downloads\SOBREPESOPOR HORA2.xlsx"

def encontrar_pasta_onedrive_empresa():
    print("[INFO] Procurando pasta do OneDrive sincronizada com SharePoint...")
    user_dir = os.environ["USERPROFILE"]
    possiveis = os.listdir(user_dir)
    for nome in possiveis:
        if "DIAS BRANCO" in nome.upper():
            caminho_completo = os.path.join(user_dir, nome)
            if os.path.isdir(caminho_completo) and "Gestão de Estoque - Documentos" in os.listdir(caminho_completo):
                print(f"[OK] Pasta encontrada: {caminho_completo}")
                return os.path.join(caminho_completo, "Gestão de Estoque - Documentos")
    print("Pasta não encontrada.")
    return None

fonte_dir = encontrar_pasta_onedrive_empresa()
if not fonte_dir:
    raise FileNotFoundError("Não foi possível localizar a pasta sincronizada do SharePoint via OneDrive.")

base_sobrepeso = os.path.join(fonte_dir, "Base_sobrepeso_real.xlsx")

print("Inicializando Excel via COM...")
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False

print("Abrindo planilha de dados brutos...")
wb = excel.Workbooks.Open(dados_raw)

print("Selecionando planilhas...")
ws_tags = wb.Sheets("Planilha5")        
ws_dados = wb.Sheets("Planilha3")

print("Coletando tags das linhas de produção...")
linhas_tags = []
for i in range(2, 24):
    linha_nome = ws_tags.Cells(i, 1).Value
    tag = ws_tags.Cells(i, 2).Value
    if linha_nome and tag:
        linhas_tags.append((linha_nome, tag))
print(f"[OK] {len(linhas_tags)} tags encontradas.")

dados = {}

for idx, (linha_nome, tag) in enumerate(linhas_tags, 1):
    print(f"[{idx}/{len(linhas_tags)}] Processando linha: {linha_nome} | Tag: {tag}")
    ws_dados.Range("A1").Value = tag
    excel.CalculateUntilAsyncQueriesDone()

    row = 4
    while True:
        data = ws_dados.Cells(row, 3).Value
        valor = ws_dados.Cells(row, 4).Value

        if data is None or not isinstance(data, datetime):
            break  # Fim da leitura

        # Se valor não for numérico, considera como 0
        if not isinstance(valor, (int, float)):
            valor = 0

        if data not in dados:
            dados[data] = {}
        dados[data][linha_nome] = valor

        row += 1
    print(f"    > {row - 4} registros coletados para {linha_nome}.")

print("Finalizando e fechando Excel...")
wb.Close(False)
excel.Quit()

print("Convertendo para DataFrame e ordenando...")
dados_normalizados = {}
for k, v in dados.items():
    try:
        k_convertido = pd.to_datetime(k).replace(tzinfo=None)
        if pd.notna(k_convertido):
            dados_normalizados[k_convertido] = v
    except (ValueError, TypeError, OutOfBoundsDatetime):
        continue  
df_final = pd.DataFrame.from_dict(dados_normalizados, orient='index')
df_final.index = pd.to_datetime(df_final.index).tz_localize(None)
df_final.index.name = "DataHora"
df_final = df_final.sort_index()


print("Salvando dados na planilha destino no SharePoint...")
with pd.ExcelWriter(base_sobrepeso, engine="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
    df_final.to_excel(writer, sheet_name="SOBREPESO", startrow=0, startcol=0)

print("Atualização concluída com sucesso!")
