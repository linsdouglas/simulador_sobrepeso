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
    df.columns = [col.strip().lower().replace(" ", "_") for col in df.columns]
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


    for col in df.columns:
        df[col] = df[col].apply(lambda x:
            x.strftime("%Y-%m-%d") if isinstance(x, (pd.Timestamp, datetime.datetime)) and not pd.isna(x)
            else x.strftime("%H:%M:%S") if isinstance(x, datetime.time) and not pd.isna(x)
            else x
        )
    print(df[["data_criacao", "data_entrada", "hora_criacao"]].head())
    df = df.where(pd.notnull(df), None)

    json_data = df.to_dict(orient='records')

    url = "https://simuladorsobrepesovitarella.com/balanca/api/upload_sap/"
    headers = {"Content-Type": "application/json"}
    batch_size = 10000
    print(f"Enviando {len(json_data)} registros em blocos de {batch_size}...")

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
                tqdm.write(f"Bloco {i//batch_size + 1} enviado com sucesso.")
            else:
                tqdm.write(f"Bloco {i//batch_size + 1} retornou erro {response.status_code}: {response.text}")

        except Exception as e:
            tqdm.write(f"❌ Bloco {i//batch_size + 1} retornou erro {response.status_code}")
            print("⚠️ Resposta resumida:", response.text[:50])  # imprime só os primeiros 300 caracteres

if __name__ == "__main__":
    envio_sap_api()