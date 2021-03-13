import streamlit as st
import pandas as pd 
import tabula
from tabula import read_pdf
import base64
from io import BytesIO


def upload_arquivo():
    data = st.file_uploader('Escolha o arquivo', type = 'pdf')
    if data is not None:
        df = read_pdf(data,pages="all",guess=False)
        df1 = pd.DataFrame(df)
        return df1

def filtra_linhas_nao_vazias(df,serie,index):
  coluna_2 = df[0][index].columns[2]
  selecao = df[0][index][coluna_2].notna()
  return serie[selecao]

def retorna_nome(df,index):

  coluna_2 = df[0][index].columns[2]

  # Extraindo a linha 4 da coluna 2, onde está o nome
  nome_agrupado = df[0][index][coluna_2].iloc[4]
  nome_agrupado = nome_agrupado.split()

  # Laço que seleciona o nome (Quando possui só letras maiúscula) 
  nome_colaborador_lista = []
  for i in nome_agrupado:
    if i == i.upper():
      nome_colaborador_lista.append(i)

  # Laço para alterar o nome de lista para string
  nome_colaborador_string = ''
  for i in nome_colaborador_lista:
      nome_colaborador_string = nome_colaborador_string + ' ' + i
  nome_colaborador_string = nome_colaborador_string.strip() #Exclui o primeiro espaço em branco

  return nome_colaborador_string 

def retorna_lista_normas(df,index):

  # Slecionando a coluna 1 onde estão os dados das normas
  coluna_1 = df[0][index].columns[0]
  lista_normas = df[0][index][coluna_1]
  lista_normas = filtra_linhas_nao_vazias(df,lista_normas,index)

  conhecimento = []
  normas = []
  # Laço para extrair o código do conhecimento e o nome da norma
  for i,j in enumerate(lista_normas):
    if i > 4 and i < len(lista_normas) - 1:
      normas_split = j.split()
      conhecimento.append(normas_split[0])

      normas_string = ''
      # Laço para alterar o nome da norma de lista para string
      for i,j in enumerate(normas_split):
        if i > 0 :
          normas_string = normas_string + ' ' + j
      normas.append(normas_string.strip()) #Exclui o primeiro espaço em branco

  return conhecimento,normas

def retorna_revisoes(df,index):
  coluna_2 = df[0][index].columns[2]
  lista_revisoes = df[0][index][coluna_2]
  lista_revisoes = filtra_linhas_nao_vazias(df,lista_revisoes,index)
  revisao_norma = []
  revisao_tlt = []

  # Laço para extrair o número da revisão_tlt e revisão_norma
  for i,j in enumerate(lista_revisoes):
    if i > 4 and i < len(lista_revisoes) - 1:
      revisao_norma.append(j.split()[-1])
      # Coloca um zero quando o valor está em branco
      if j.split()[-3] ==  'a':
        revisao_tlt.append('0')
      else:
        revisao_tlt.append((j.split()[-3]))

  return revisao_tlt,revisao_norma

def filtra_colaboradores(df):
    colaboradores = ['ADEMIR DOS SANTOS','CARLOS STAHL','DEBORA CRISTINA HINDGES KLAUS',
                 'EDERSON LUIS BRANCO','ENDRIU DE OLIVEIRA MELO','JEAN CARLOS BRUCH', 
                 'JEFFERSON CARDOSO DA SILVA EVANGELISTA','JOAO ALMIR MORSCH','JOCIMAR KOBERNOVICZ',
                 'JOSE MARIVALDO DA CONCEICAO SANTOS','LEANDRO NILSEN', 'LUIZ RICARDO PINHEIRO','MARCO AURELIO BARUFFI',
                 'MAURICIO JOSE GORGES','PAULO SOARES DE OLIVEIRA','RIVELINO STEIN', 'SONIA SARDAGNA NARLOCH',
                 'VALFRIDO MAXIMILIANO MEYER SILVEIRA','VANDERLEI BERTOL']
    lista = []
    for i in df['Colaborador']:
        if (i in colaboradores):
            lista.append(True)
        else:
            lista.append(False)
    return lista

def cria_dataframe(df):
    dataframe = pd.DataFrame()                                                
    for index in range(df.shape[0]):  
        colaborador = retorna_nome(df,index)                                      
        conhecimento,descricao = retorna_lista_normas(df,index)
        rev_tlt,rev_normas = retorna_revisoes(df,index)

        dados = {'Conhecimento':conhecimento,'Descrição_Norma':descricao,'Colaborador': colaborador,
                    'Rev_TLT':rev_tlt,'Rev_Norma':rev_normas}

        dataframe = pd.concat([dataframe,pd.DataFrame(dados)],ignore_index=True) 

    dataframe[['Conhecimento','Rev_TLT','Rev_Norma']] = dataframe[['Conhecimento','Rev_TLT','Rev_Norma']].astype('int')
    selecao = filtra_colaboradores(dataframe)
    dataframe = dataframe[selecao]
    dataframe.drop_duplicates(inplace=True)
    dataframe.set_index('Conhecimento',inplace=True)
    dataframe.rename_axis(['Teste'],axis=1,inplace=True)
    return dataframe

def converter_para_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index = False, sheet_name='Plan1')
    workbook  = writer.book
    worksheet = writer.sheets['Plan1']
    format1 = workbook.add_format({'num_format': '0.00'}) # Tried with '0%' and '#,##0.00' also.
    worksheet.set_column('A:A', None, format1) # Say Data are in column A
    writer.save()
    processed_data = output.getvalue()
    return processed_data

def download_link(df):
    """Generates a link allowing the data in a given panda dataframe to be downloaded
    in:  dataframe
    out: href string
    """
    val = converter_para_excel(df)
    b64 = base64.b64encode(val)  # val looks like b'...'
    return f'<button kind="primary" class="css-2trqyj edgvbvh1"><a href="data:application/octet-stream;base64,{b64.decode()}" style="{css()}" download="saida.xlsx">Download</a></button>' # decode b'abc' => abc

def css():
    css = '''
    color: inherit;
    text-decoration: none;
    font-size: 0.9rem '''
    return css

def main():
    st.markdown('# Converta seu PDF para Excel')
    df_pdf = upload_arquivo()

    if df_pdf is not None:
      if st.button('Converter'):
          df = cria_dataframe(df_pdf)
          st.dataframe(df.head())

          st.markdown(download_link(df), unsafe_allow_html=True)

if __name__ == '__main__':
    main()