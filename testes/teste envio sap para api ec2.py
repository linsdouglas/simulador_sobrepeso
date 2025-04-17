import pandas as pd
import requests
import datetime
import certifi
import os
from tqdm import tqdm

def envio_sap_api():
    download_dir = os.path.join(os.environ['USERPROFILE'], 'Downloads', 'base_sap')
    arquivo_excel = os.path.join(download_dir, 'EXPORT.XLSX')

    df = pd.read_excel(arquivo_excel)

    # Normaliza nomes das colunas
    df.columns = [col.strip().lower().replace(" ", "_") for col in df.columns]

    # Renomeia colunas para compatibilidade com modelo Django
    df.rename(columns={
        "doc.material": "doc_material",
        "ano_doc.material": "ano_doc_material",
        "item_doc.material": "item_doc_material",
        "data_de_entrada": "data_entrada",
        "depósito": "deposito",
        "data_do_vencimento": "data_vencimento",
        "data_de_produção": "data_producao",
        "qtd.__um_registro": "qtd_um_registro",
        "nome_do_usuário": "nome_usuario",
        "data_de_criação": "data_criacao",
        "hora_de_criação": "hora_criacao",
        "data_de_modificação": "data_modificacao",
        "hora_de_modificação": "hora_modificacao"
    }, inplace=True)

    print("Colunas normalizadas:", df.columns.tolist())

    # Limpeza rigorosa da coluna chave_pallet
    df['chave_pallet'] = df['chave_pallet'].astype(str).str.strip()

    # Diagnóstico antes do filtro
    print("📊 Registros totais antes do filtro:", len(df))
    print("📊 Linhas com chave_pallet vazia:", (df['chave_pallet'] == '').sum())
    print("📊 Linhas com chave_pallet == 'None':", (df['chave_pallet'] == 'None').sum())
    print("📊 Linhas com chave_pallet == 'nan':", (df['chave_pallet'].str.lower() == 'nan').sum())

    # Remove registros com chave_pallet inválida
    df = df[~df['chave_pallet'].isin([None, '', 'nan', 'NaN', 'None'])]

    print("📊 Registros restantes após filtro:", len(df))

    # Converte datas e horas
    for col in df.columns:
        df[col] = df[col].apply(lambda x:
            x.strftime("%Y-%m-%d") if isinstance(x, (pd.Timestamp, datetime.datetime)) and not pd.isna(x)
            else x.strftime("%H:%M:%S") if isinstance(x, datetime.time) and not pd.isna(x)
            else x
        )

    print("📌 Preview de datas:")
    print(df[["data_criacao", "data_entrada", "hora_criacao"]].head())

    # Substitui NaN por None para o JSON
    df = df.where(pd.notnull(df), None)

    json_data = df.to_dict(orient='records')

    # Diagnóstico do primeiro registro
    print("🧪 Exemplo do primeiro registro a ser enviado:")
    print(json_data[0])
    for i, row in enumerate(json_data):
        if not row["chave_pallet"]:
            print(f"❌ Linha {i} com chave_pallet vazia detectada:", row)

    url = "https://simuladorsobrepesovitarella.com/balanca/api/upload_sap/"
    headers = {"Content-Type": "application/json"}
    batch_size = 10000
    print(f"🚀 Enviando {len(json_data)} registros em blocos de {batch_size}...")

    for i in tqdm(range(0, len(json_data), batch_size), desc="Enviando blocos"):
        bloco = json_data[i:i+batch_size]

        try:
            response = requests.post(
                url,
                json=bloco,
                headers=headers,
                verify=False,
                timeout=60
            )

            if response.status_code in [200, 201]:
                tqdm.write(f"✅ Bloco {i//batch_size + 1} enviado com sucesso.")
            else:
                tqdm.write(f"⚠️ Bloco {i//batch_size + 1} retornou erro {response.status_code}")
                print("Resposta resumida:", response.text[:300])

        except Exception as e:
            tqdm.write(f"❌ Erro ao enviar bloco {i//batch_size + 1}: {str(e)}")

if __name__ == "__main__":
    envio_sap_api()
