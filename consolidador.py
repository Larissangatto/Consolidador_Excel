# Importando as bibliotecas a serem usadas

import pandas as pd
import os
import datetime

data = datetime.datetime.now()

#  criando um Dataframe vazio com as colunas
colunas = ['Segmento',
           'País',
           'Produto',
           'Qtde de Unidades Vendidas',
           'Preço Unitário',
           'Valor Total',
           'Desconto',
           'Valor Total c/ Desconto',
           'Custo Total',
           'Lucro',
           'Data',
           'Mês',
           'Ano']

consolidado = pd.DataFrame(columns=colunas)

# busca o nome dos arquivos a serem consolidados

arquivos = os.listdir("planilhas")

# realiza a consolidação dos arquivos e cria lista de erros


for arquivo in arquivos:
    
    if arquivo.endswith('.xlsx'):
        informacao = arquivo.split('-')
        segmento = informacao[0]
        pais = informacao[1].replace('.xlsx','')
        
        try:
            df = pd.read_excel(f"planilhas\\{arquivo}")
            df.insert(0,'Segmento',segmento)
            df.insert(1, 'País', pais)
            consolidado = pd.concat([consolidado,df])
        except:
            
            with open('log_erro.txt','a') as erro:
                erro.write(f"Erro ao tentar consolidar o arquivo: {arquivo}.\n")
    else:
        with open('log_erro.txt','a') as erro:
            erro.write(f"\n O arquivo: {arquivo} não foi consolidados pois não é um arquivo Excel.\n")

#Exporta o DataFrame consolidado para um arquivo Excel
consolidado.to_excel(f"Report-consolidado-{data.strftime('%d.%m.%Y')}.xlsx", index = False,sheet_name='Report consolidado')