from sqlalchemy import create_engine
from urllib.parse import quote_plus
import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
expdicao_file = os.path.join(BASE_DIR, 'balanca', 'data', 'dados_expedicao.xlsx')

df = pd.read_excel(expdicao_file)

senha = quote_plus("85838121aA@")
engine = create_engine(f"mysql+mysqlconnector://root:{senha}@localhost:3306/base_expedicao")
df.to_sql('tabela_exped', con=engine, if_exists='replace', index=False)

print("Tabela criada e dados inseridos com sucesso!")
