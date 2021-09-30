#importanto bibliotecas
import os
import pandas as pd
import easygui as eg

#Abrindo e Salvando na VAR Arquivo Usuario
abrindoexcel = eg.fileopenbox(title='Selecione a Planilha a ser Aberta', default="C:/")

#lendo arquivo excel e salvando na var
planilha_df = pd.read_excel(abrindoexcel)

#filtrando apenas os que geram SPED
gerasped = planilha_df['RAZÃO SOCIAL']

#local onde criar as pastas
choice_dir = eg.diropenbox(title='Onde Criar as Pastas do SPED')

dont_create_list = ["", "nan", "Site Email", "https://webmail.consultsistema.com.br"]
for x in gerasped.items():
    nome_pasta = x[1]
    if (str(nome_pasta) not in dont_create_list):
        if not os.path.isdir(f'{choice_dir}/{nome_pasta}'):
            os.mkdir(f'{choice_dir}/{nome_pasta}')
            print(f'Pasta Criada: {nome_pasta}')
        else:
            print(f'Pasta já existe: {nome_pasta}')
