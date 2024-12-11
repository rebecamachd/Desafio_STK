### 0. Importação de bibliotecas
from pathlib import Path
import requests
import pandas as pd
from openpyxl import load_workbook
import re
import numpy as np
### 1. Definição de funções
def rename_column(col): #Só pras que começam com Q
    match = re.match(r'(\dQ)\s?(\d{4})', col)
    if match:
        return f"{match.group(1)}{match.group(2)[-2:]}"  # Mantém apenas os dois últimos dígitos do ano
    return col
def formata_tabelas (df,sheet_name):

    ''' Função para formatar a planilha bruta importada via API'''

    # ----------- Ajustar o header do dataframe  ---------------------------------------------
    df.columns = df.iloc[0]  # Atribuir a primeira linha como cabeçalho
    df = df[1:]  # Remover a linha do cabeçalho original

    #Tirar colunas inuteis
    #Quero padronizar para que a primeira coluna seja sempre a referencia principal entao é importante verificar se tem alguma coluna inutil no inicio

    df.dropna(axis=1,how='all',inplace=True)

    # -----------   Criar coluna para categoria superior (Quando a primeira linha tem informação e o restante é NaN)  -----------
    #Pra que serve? Ex: Temos a referencia "Receita" em diferentes ca

    # Identificar linhas onde 'Categoria Superior' deve ser atualizada
    mascara_categoria_superior = df.iloc[:, 1:].isna().all(axis=1)

    # Criar uma coluna temporária para armazenar os valores de 'Categoria Superior'
    df['Categoria Superior'] = df.loc[mascara_categoria_superior, df.columns[0]]

    # Preencher os valores ausentes com o último valor válido
    df['Categoria Superior'] = df['Categoria Superior'].fillna(method='ffill')

    #Retirar valores Nan da coluna

    df = df[~df['Categoria Superior'].isna()]
    #--------------------------------------------------------------------------------

    #Retirar Linhas inuteis
    df = df.loc[~mascara_categoria_superior]

    # ---------- ---- Criar coluna de referencia -------------------------------------

    #Defir valores da coluna 
    df['Referencia'] = df['Categoria Superior'] + ' - ' + df[df.columns[0]] #Usa a primeira coluna como base, pois é a que geralmente será a referencia

    # Remover colunas auxiliareS
    df.drop(columns=[df.columns[0], 'Categoria Superior'], inplace=True)

    # ---------- ---- Criar coluna de referencia da sheet -------------------------------------

    ref = df.pop('Referencia')  # Remove a coluna e armazena em uma variável
    df.insert(0, 'Referencia', ref)
    df.insert(1, 'Sheet', sheet_name)

    # ---------- ---- Formatar colunas numericas  -------------------------------------

    #Garantindo tudo numérico antes de transformar
    for coluna in df.columns: 
        try:
            df[coluna]= pd.to_numeric(df[coluna])
        except ValueError:
            pass

    #Substitui separador "." para vírgula para facilitar intepretação no Excel; também formata para 2 casas decimais para facilitar visualização
    colunas_numericas = df.select_dtypes(include='number').columns
    df[colunas_numericas] = df[colunas_numericas].applymap(lambda x: f"{x:.2f}".replace(".", ","))
    df.columns

    #Substitui nan por - para evitar poluição visual
    df = df.replace('nan','-')

    #--------------- Formatar nome e tipo das colunas  ----------------------------------------

    df.columns = df.columns.astype(str)
    df.columns = [col.split('.')[0] if '.' in col else col for col in df.columns]
    df = df.rename(columns=rename_column)

    return df
### 2. Parametros gerais

# URL da API
url = 'https://api.mziq.com/mzfilemanager/v2/d/0afe1b62-e299-4dec-a938-763ebc4e2c11/79c66cca-4d48-0e24-db90-67450e78b597?origin=1'
cod_sucesso = 200
diretorio_export = str(Path.cwd())
nome_arquivo = 'Base_BTG'
### 3. Importação e tratamento
# Baixa o arquivo via API, faz a formatação de cada dataframe e concatena para formar o bando
response = requests.get(url)
if response.status_code == cod_sucesso:
    with open("temp.xlsx", "wb") as f:
        f.write(response.content)
else:
    print(f"Erro ao baixar o arquivo: {response.status_code}") 
    exit()

# Carrega os nomes das sheet_name usando openpyxl
workbook = load_workbook("temp.xlsx")
sheets = workbook.sheetnames
del workbook
print(f"sheet_names disponíveis: {sheets}")

# Lê cada sheet_name como DataFrame
dfs = {} #Dicionario vai acumular dataframes 
base = pd.DataFrame()
for sheet_name in sheets:
    print(sheet_name)
    # try:
    #     dfs[sheet_name] = pd.read_excel("temp.xlsx", sheet_name=sheet_name) #acessar cada workbook
    #     base = pd.concat([base, formata_sheet(dfs[sheet_name])] )
    #     print(f"sheet_name '{sheet_name}' lida com sucesso.")
    # except Exception as e:
    #     print(f"Erro ao ler a sheet_name '{sheet_name}': {e}")
    dfs[sheet_name] = pd.read_excel("temp.xlsx", sheet_name=sheet_name) #acessar cada workbook
    base = pd.concat([base, formata_tabelas(dfs[sheet_name],sheet_name)])
        
    
### 4. Exportação
#Definir quais sao as colunas de interesse
cols_interesse_padrao = ["Referencia","Sheet"]
sufixo_interesse = '23'
cols_interesse = base.columns[base.columns.str.contains(sufixo_interesse)].tolist()
colunas = cols_interesse_padrao + cols_interesse
#tabela completa
base.to_csv(rf"{diretorio_export}\{'Base_BTG_completa'}.csv", index = False, encoding = "utf-8-sig",sep=';',decimal=',')

#tabela filtrada (2023) - entendi que a dinamica era pegar os dados de 2023 
base[colunas].to_csv(rf"{diretorio_export}\{'Base_BTG_23'}.csv", index = False, encoding = "utf-8-sig",sep=';',decimal=',')