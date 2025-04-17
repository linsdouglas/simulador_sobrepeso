import pandas as pd
import requests
import datetime
import certifi
import os

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

    try:
        response = requests.post(url, json=json_data, headers=headers, verify=False)
        print("Status:", response.status_code)
        print("Resposta:", response.text)
    except Exception as e:
        print("Erro ao enviar para a API:", str(e))

if __name__ == "__main__":
    envio_sap_api()
