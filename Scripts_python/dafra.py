import pandas as pd
import os
import pyodbc


#mostra os tipod de dados de um dataframe
def tipo_dados(df):
    print(f'Coluna \t\t\t Tipo de dados')
    for coluna in df.columns:
        print(f"{coluna} \t\t\t {df[coluna].dtypes}")
   
# converter tipo de dados     
def converter_tipo(df,coluna,tipo):
    if (tipo == str) or (tipo == int) or (tipo == float):
        df[coluna] = df[coluna].astype(tipo)
    else:
        df[coluna] = pd.to_datetime(df[coluna])
        
# Filtra por ano
def filtra_entre_anos(df,coluna,ano1,ano2):
    df_filtrado = df[(df[coluna].dt.year >= ano1) & (df[coluna].dt.year <= ano2)]
    return df_filtrado

# filtra por ano e mês
def filtro_ano_mes(df, coluna, ano1, ano2, mes1, mes2):
    ano_mes = df[(df[coluna].dt.year >= ano1) & 
                 (df[coluna].dt.year <= ano2) & 
                 (df[coluna].dt.month >= mes1) & 
                 (df[coluna].dt.month <= mes2)]
    return ano_mes

# filtra por ano e mês e dia
def filtro_ano_mes_dia(df, coluna, ano1, ano2, mes1, mes2,dia1,dia2):
    ano_mes_dia = df[(df[coluna].dt.year >= ano1) & 
                 (df[coluna].dt.year <= ano2) & 
                 (df[coluna].dt.month >= mes1) & 
                 (df[coluna].dt.month <= mes2) &
                 (df[coluna].dt.day >= dia1) & 
                 (df[coluna].dt.day <= dia2)]
    return ano_mes_dia

#formato brasileiro data sem horas
def mudar_padrao_brasileiro_semHoras(df,coluna):
    df[coluna] = pd.to_datetime(df[coluna])
    df[coluna] = df[coluna].dt.strftime('%d/%m/%Y')
    return df

#formato brasileiro data sem horas
def mudar_padrao_brasileiro_comHoras(df,coluna):
    df[coluna] = pd.to_datetime(df[coluna])
    df[coluna] = df[coluna].dt.strftime('%d/%m/%Y %H:%M:%S')
    return df

# Função para localizar um valor de uma coluna
def filtro_igual_a(df,coluna,valor):
    df_igual = df.loc[df[coluna] == valor]
    return df_igual


# Função para localizar um valor de uma coluna numerica que seja maior que
def filtro_maior_que(df,coluna,valor):
    df_maior = df.loc[df[coluna] > valor]
    return df_maior


# Função para localizar um valor de uma coluna numerica que seja menor ou igual a
def filtro_menor_igual_a(df,coluna,valor):
    df_menor_igual_a = df.loc[df[coluna] <= valor]
    return df_menor_igual_a

# Função para localizar um intervalo de valor de uma coluna numerica 
def filtro_entre(df,coluna,valor1,valor2):
    df_entre = df.loc[(df[coluna] >= valor1) & (df[coluna] <= valor2)]
    return df_entre

# função para encontrar um conjunto de registros de uma coluna 
def varios_valores(df,coluna,lista):
    df_itens = df.loc[df[coluna].isin(lista)]
    return df_itens

#função que captura os parametros da pesquisa por entrada
def pergunta_m():
    coluna = input('Digite o nome da coluna: ')
    n = int(input('Digite quantos registros deseja buscar: '))
    pergunta = int(input('Digite 0 para coluna numerica ou 1 para categorica: '))
    itens = []
    if pergunta == 0:
        for i in range(n):
            itens.append(int(input()))    
    elif pergunta == 1:
        for i in range(n):
            itens.append(input())
    else:
        print("erro")
    return [itens,coluna]


#Verifica se há nulos (base)
def show_null(df):
    null_columns = (df.isnull().sum(axis=0)/len(df)).sort_values(ascending=False).index
    
    null_data = pd.concat([df.isnull().sum(axis=0), 
                           (df.isnull().sum(axis=0)/len(df)).sort_values(ascending=False), 
                           df.loc[:, df.columns.isin(list(null_columns))].dtypes], 
                          axis=1)
    
    null_data = null_data.rename(columns={0: '#', 
                                          1: '% null', 
                                          2: 'type'}).sort_values(ascending=False, 
                                                                  by='% null')
    
    return null_data

#trata os nulos
def tratar_colunas_NAN(df,var_NaN,lista_prenchimento):
    filtered_df = var_NaN[var_NaN['#'] > 0]
    colunas_com_null = filtered_df.index.tolist()
    for i in range(len(colunas_com_null)):
        df[colunas_com_null[i]] = df[colunas_com_null[i]].fillna(lista_prenchimento[i])
    return df

#Função substituir valores de uma determinada coluna
def substituir_valores(df,coluna,valor_atual,valor_substituto):
    df[coluna] = df[coluna].replace(valor_atual,valor_substituto)
    
#Substituindo varios registros de uma só vez - Usando dicionario mapeamento ('Antigo' : 'novo')
def substituir_varios(df,coluna,mapeamento):
    df[coluna] = df[coluna].replace(mapeamento)
    
#consolidando varios arquivos em um unico arquivo
def consolidar_arquivos(diretorio,tipo_arquivo):
    # Diretório dos arquivos Excel
    caminho = (diretorio)

    # Lista para armazenar os dados dos arquivos
    dados = []

    # Iterar sobre os arquivos Excel no diretório
    for arquivo in os.listdir(caminho):
        if arquivo.endswith(tipo_arquivo):
            caminho_arquivo = os.path.join(caminho, arquivo)
            df = pd.read_excel(caminho_arquivo)
            dados.append(df)

    # Consolidar os dados em um único DataFrame
    consolidado = pd.concat(dados)
    return consolidado

# Função que salva o arquivo consolidado em uma pasta especifica
def salvar_consolidado(consolidado,pasta_destino,nome_arquivo,extensao_arquivo):
    # Caminho completo do arquivo Excel
    Caminho_Completo = pasta_destino + nome_arquivo + extensao_arquivo
    if extensao_arquivo == '.xlsx':
        # Salvar o DataFrame como arquivo Excel
        consolidado.to_excel(Caminho_Completo, index=False)
    elif extensao_arquivo == '.csv':
        # Salvar o DataFrame como arquivo CSV
        consolidado.to_csv(Caminho_Completo, index=False,sep=',')
    elif extensao_arquivo == '.txt':
        # Salvar o DataFrame como arquivo CSV
        consolidado.to_csv(Caminho_Completo, index=False,sep=',')
    return consolidado


def sobreescrever_arquivoExcel(pasta_destino,nome_arquivo,consolidado):
    Caminho_Completo = pasta_destino + nome_arquivo 
    # Verificar se o arquivo Excel já existe
    if os.path.exists(Caminho_Completo):
        # Excluir o arquivo existente
        os.remove(Caminho_Completo)
        #print(f"Arquivo {nome_arquivo} existente foi excluído.")

    # Salvar o DataFrame como arquivo Excel
    consolidado.to_excel(Caminho_Completo, index=False)
    
#Cria nova coluna
def nova_coluna(df,nome_coluna,valor_coluna):
    df[nome_coluna] = valor_coluna
    return df

# Adicionar nova coluna com o nome do mês em português - 
def coluna_meses_extenso(df,nova_coluna,lista,coluna_ref):
    df[nova_coluna] = df[coluna_ref].dt.month.map(lambda x: lista[x-1])
    return df

#Adicionar hora e minutos no dataframe
def Adiciona_hora_minuto(df,coluna_ref):
    df['hora'] = df[coluna_ref].dt.hour
    df['minuto'] = df[coluna_ref].dt.minute
    return df


# TRATAMENTO DE DADOS DE PLANILHAS EXCEL SEM PADRAO

# Importando por aba selecionada
def importacao_excel_por_aba(diretorio,aba):
    df = pd.read_excel(diretorio,sheet_name=aba)
    return df

# Pula linhas para capturar apenas os dados
def pular_linhas(diretorio,aba,nlinhas):
    df = pd.read_excel(diretorio,sheet_name=aba,skiprows=nlinhas)
    return df

# seleciona colunas
def selecao_coluna(diretorio,aba,colSelecionadas):
    df = pd.read_excel(diretorio,sheet_name=aba,usecols=colSelecionadas)
    return df

# Pula linha e seleciona colunas
def pular_linhas_selColunas(diretorio,aba,nlinhas,colSelecionadas):
    df = pd.read_excel(diretorio,sheet_name=aba,skiprows=nlinhas,usecols=colSelecionadas)
    return df

# Seleciona colunas e linhas
def seleciona_linhas(diretorio,aba,nlinhas,colSelecionadas):
    df = pd.read_excel(diretorio,sheet_name=aba,nrows=nlinhas,usecols=colSelecionadas)
    return df

# traz as colunas especificas de um dataframe
def trazer_colunas(df, lista_colunas):
    df = df[lista_colunas]
    return df

# Remova diretamente  as duplicatas do DataFrame inplace - df é a base de dados
def remover_duplicadas(df):
    df.drop_duplicates(inplace=True)
    duplicados = df[df.duplicated()]
    return duplicados

# Função que ordena coluna
def ordenar_coluna(dados,coluna,ordenacao):
    if ordenacao == 'c':
        dados=dados.sort_values(by=[coluna],ascending = True )
    else:
        dados=dados.sort_values(by=[coluna],ascending = False )
    return dados



# Gerador de colunas e values para evitar erro
def values_conexao(df):  
    colunas = str(df.columns).replace("'","").split() 
    for i in range(len(colunas)-1):
        if i == 0:
            print(colunas[i][7:],end='')
        elif i == len(colunas)-2:
            print(colunas[i][:-2])
        else:
            print(colunas[i],end='')
    
    valores = (str(df.columns).replace("'","linha.")).replace("linha.,",",").split()
    for i in range(len(valores)-1):
        if i == 0:
            print(valores[i][7:],end='')
        elif i == len(valores)-2:
            print(valores[i][:-8])
        else:
            print(valores[i],end='')
            
    a = ''
    for i in range(df.shape[1]):
        if i == df.shape[1]-1:
                a+='?'
        else:
                a+='?,'
    return a