import pandas as pd
import numpy as np
import plotly as plt
import scipy.stats as st
from datetime import datetime

def calculo_NS(lista):
    return 0.7 if lista in ['A', 'B', 'C'] else ''

def calcular_media(df, colunas):
    return df.iloc[:, colunas].mean(axis=1).round(2)

def calcular_desvpad(df, colunas):
    return df.iloc[:, colunas].std(axis=1).round(2)

def calcular_dias(estoque, cmd):
    return np.round(estoque / cmd)

# Importar o arquivo
df = pd.read_excel('seu_arquivo.xlsx')

# CMU - ajuste de casas decimais
df['CMU'] = df['CMU'].round(2)

# Calculando as médias
df['Mean_24'] = calcular_media(df, range(1, 25))
df['Mean_12'] = calcular_media(df, range(13, 25))
df['Mean_06'] = calcular_media(df, range(19, 25))

# Calcular o consumo médio diário (CMD)
df['CMD'] = round(df['Mean_24'] / 22, 2)

# Calcular o Desvio Padrão da amostra (24M)
df['DesvPad'] = calcular_desvpad(df, range(1, 25))

# Calcular a demanda durante o lead time
df['DL'] = round(df['CMD'] * (df['LT_Fornec'] + df['SLA_Compras']), 2)

# Calcular o nível de serviço (NS)
df['NS'] = df['Curva ABC'].apply(calculo_NS)

# Calculo do estoque de segurança (ES)
df['ES'] = round(st.norm.ppf(df['NS']) * df['DesvPad'])

# Calculo Ponto de Pedido (PP)
df['PP'] = round(df['CMD'] * df['Freq'] + df['ES'])

# Calculo do Estoque Máximo (Emax)
df['Emax'] = round(df['CMD'] * (df['LT_Fornec'] + df['SLA_Compras']) + df['PP'])

# Convertendo os níveis de estoque em dias
df['ES_dias'] = calcular_dias(df['ES'], df['CMD'])
df['PP_dias'] = calcular_dias(df['PP'], df['CMD'])
df['Emax_dias'] = calcular_dias(df['Emax'], df['CMD'])

# Calcular o valor de estoque máximo
df['Política_Sugerida_($)'] = df['Emax'] * df['CMU']

# Salvar arquivo com data atual
output_file = f"escolha_o_nome_arquivo_{datetime.now().strftime('%m%Y')}.xlsx"
df.to_excel(output_file, index=False)
