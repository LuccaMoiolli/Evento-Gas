import pandas as pd

file_path = 'ACOMPANHAMENTO DOS VOLUMES  PROGRAMADOS.xlsx'

# colocando os nomes das colunas
planilha = open('planilha.csv', 'w')
planilha.write('id_contrato,data,vol_medido,PCS,vol_corrigido,vol_programado,lim_sup,lim_inf,penalidade_maior_apuracao,penalidade_menor_apuracao,penal_maior_QDC\n')
planilha.close()

xls = pd.ExcelFile(file_path)

# itera pelos meses inclusos
for sheet in range(77):
    quinzena_break = 11 if sheet >= 50 else 10

    # coloca uma planilha num dataframe já formatada para pegar somente os dados relevantes (ainda com um pouco de desperdício)
    if sheet >= 50:
        df = pd.read_excel(xls, sheet_name=sheet, usecols='D:T', nrows=23, skiprows=[0,1,2,3,4,7,8,9,10,11,12,23])
    else:
        df = pd.read_excel(xls, sheet_name=sheet, usecols='D:T', nrows=21, skiprows=[0,1,2,3,4,7,8,9,10,11,21])
    
    # repete o id contrato em todas as linhas
    df.iloc[0] = df.iloc[0,0]
    df.iloc[quinzena_break] = df.iloc[0,0]

    df = df.transpose()
    df = df.drop('Unnamed: 3')  # joga fora identificadores de coluna (desperdíco mencionado acima)

    # cast pra inteiro
    # lidando com a separação do mês em 2
    df.iloc[:, 2:quinzena_break] = df.iloc[:, 2:quinzena_break].fillna(0).astype('int64')
    df.iloc[:, 0:quinzena_break].to_csv('planilha.csv',header= False, index= False, mode= 'a')
    df.iloc[:, quinzena_break+2:] = df.iloc[:, quinzena_break+2:].fillna(0).astype('int64')
    df.iloc[:, quinzena_break:].to_csv('planilha.csv',header= False, index= False, mode= 'a')
