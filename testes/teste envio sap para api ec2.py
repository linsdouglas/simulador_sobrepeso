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

    for col in df.columns:
        df[col] = df[col].apply(lambda x:
            x.strftime("%Y-%m-%d") if isinstance(x, (pd.Timestamp, datetime.datetime)) and not pd.isna(x)
            else x.strftime("%H:%M:%S") if isinstance(x, datetime.time) and not pd.isna(x)
            else x
        )

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
            tqdm.write(f"Erro ao enviar bloco {i//batch_size + 1}: {str(e)}")

if __name__ == "__main__":
    envio_sap_api()