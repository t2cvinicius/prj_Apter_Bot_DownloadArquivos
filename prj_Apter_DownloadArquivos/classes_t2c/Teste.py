
from babel.dates import format_date

import pandas as pd

#Imagem de erro caso o botão seja laranja
if not bot.find( "Erro", matching=0.97, waiting_time=10000):
    not_found("Erro")
    
    
#Suspenso ate tal horario 
if not bot.find( "SuspensoAte", matching=0.97, waiting_time=10000):
    not_found("SuspensoAte")


var_CaminhoTabela = r"C:\Users\user\Desktop\Nova pasta\Fiscal\Fiscal.xlsx"
var_CaminhoTabelaNova = "C:\\robo\\BotCity\\Projetos\\prj_Apter_Bot_DownloadArquivos\\prj_Apter_DownloadArquivos\\resources\\UiPath\\Novo(a) Planilha do Microsoft Excel.xlsx"

dados_excel = pd.read_excel(var_CaminhoTabela)

dados_excel["Data Início"] = dados_excel["Data Início"].apply(lambda a: a +'T00:00:00')
dados_excel["Data Fim"] = dados_excel["Data Fim"].apply(lambda a: a +'T00:00:00')


dados_excel["Data Fim"] = pd.to_datetime(dados_excel["Data Fim"],format='%Y-%m-%dT%H:%M:%S').dt.normalize()
dados_excel["Data Início"] = pd.to_datetime(dados_excel["Data Início"],format='%Y-%m-%dT%H:%M:%S').dt.normalize()

dados_excel['Data Fim'] = pd.to_datetime(dados_excel['Data Fim'],errors='raise', format="%d/%m/%Y",dayfirst=True,utc=False )
dados_excel['Data Início'] = pd.to_datetime(dados_excel['Data Início'],errors='raise', format="%d/%m/%Y",dayfirst=True,utc=False )

print ("------------------------------------------DENTRO PROCESS----------------------------------------------------------------")

# arg_dfTable['Data Fim'] = pandas.to_datetime(arg_dfTable['Data Fim'], format='%Y-%m-%dT%H:%M:%S', errors='coerce')
# arg_dfTable['Data Início'] = pandas.to_datetime(arg_dfTable['Data Início'], format='%Y-%m-%dT%H:%M:%S', errors='coerce')


# arg_dfTable['Data Início'] = format_date(arg_dfTable['Data Início'],locale='de_DE')
# arg_dfTable['Data Fim'] = format_date(arg_dfTable['Data Fim'],locale='de_DE')


# arg_dfTable['Data Fim'] = pandas.to_datetime(arg_dfTable['Data Fim'], format='%d/%m/%Y')
# arg_dfTable['Data Início'] = pandas.to_datetime(arg_dfTable['Data Início'], format='%d/%m/%Y')


# arg_dfTable['Data Início'] = arg_dfTable['Data Início'].dt.strftime(date_format='%d/%m/%Y') #Coloca no formato certo mas deixa como string no excel
# arg_dfTable['Data Fim'] =arg_dfTable['Data Fim'].dt.strftime(date_format='%d/%m/%Y') #Coloca no formato certo mas deixa como string no excel

# arg_dfTable['Data Início'] = arg_dfTable['Data Início'].dt.strftime(date_format='%d/%m/%Y')
# arg_dfTable['Data Início'] = pandas.to_datetime(arg_dfTable['Data Início'],errors='raise', format="%d/%m/%Y",dayfirst=True)



# arg_dfTable.style.format({"Data Início": lambda t: t.strftime("%d/%m/%Y")})
# arg_dfTable.style.format({"Data Fim": lambda t: t.strftime("%d/%m/%Y")})

# arg_dfTable['Data Início'] = pandas.to_datetime(arg_dfTable['Data Início'],errors='raise', format="%d/%m/%Y",dayfirst=True).dt.date()
# arg_dfTable['Data Fim'] = pandas.to_datetime(arg_dfTable['Data Fim'],errors='raise', format="%d/%m/%Y",dayfirst=True,utc=False ).dt.date

print ("----------------------------------------------------------------------------------------------------------")



# print(dados_excel.dtypes)

# dados_excel['Data Início'] = format_date(dados_excel['Data Início'],locale= 'en_US')
# dados_excel['Data Fim'] = format_date(dados_excel['Data Fim'],locale='de_DE')

# dados_excel.style = dados_excel.style.format({"Data Início": lambda t: t.strftime("%d/%m/%Y")})
# dados_excel.style = dados_excel.style.format({"Data Fim": lambda t: t.strftime("%d/%m/%Y")})

dados_excel.to_excel(var_CaminhoTabelaNova,index=False)
# dados_excel['Data Fim'].to_excel(var_CaminhoTabelaNova,index=False)



# # Converta a coluna 'Data Início' para datetime
# dados_excel['Data Início'] = pd.to_datetime(dados_excel['Data Início'])

# # Aplique a formatação para cada valor na coluna 'Data Início'
# dados_excel['Data Início'] = dados_excel['Data Início'].apply(lambda x: format_date(x, format='full', locale='de_DE'))

# # Salve o DataFrame de volta no arquivo Excel
# dados_excel.to_excel(var_CaminhoTabelaNova, index=False)


# Converta a coluna 'Data Início' para datetime e aplique a formatação
# dados_excel['Data Fim'] = pd.to_datetime(dados_excel['Data Fim']).dt.strftime('%d/%m/%Y')

# # Converta a coluna 'Data Início' para datetime64
# # dados_excel['Data Início'] = pd.to_datetime(dados_excel['Data Início'], format='%d/%m/%Y')

# # Salve o DataFrame de volta no arquivo Excel
# dados_excel.to_excel(var_CaminhoTabelaNova, index=False)


# # Salve o DataFrame de volta no arquivo Excel com o formato desejado
# with pd.ExcelWriter(var_CaminhoTabelaNova, engine='xlsxwriter') as writer:
#     dados_excel.to_excel(writer, index=False, sheet_name='Sheet1')
#     worksheet = writer.sheets['Sheet1']
    
#     # Adicione a formatação específica para a coluna de data
#     formato_data = writer.book.add_format({'num_format': 'dd/mm/yyyy'})
#     worksheet.set_column('B:B', None, formato_data)

# Converta a coluna 'Data Início' para datetime e aplique a formatação
# dados_excel['Data Início'] = pd.to_datetime(dados_excel['Data Início']).dt.strftime('%d/%m/%Y')

# # Crie um escritor Excel usando openpyxl
# with pd.ExcelWriter(var_CaminhoTabelaNova, engine='openpyxl') as writer:
#     # Salve o DataFrame no arquivo Excel
#     dados_excel.to_excel(writer, index=False, sheet_name='Sheet1')

#     # Acesse a planilha para configurar a formatação
#     workbook = writer.book
#     worksheet = writer.sheets['Sheet1']

#     # Adicione a formatação específica para a coluna de data
#     formato_data = workbook.add_format({'num_format': 'dd/mm/yyyy'})
#     worksheet.set_column('B:B', None, formato_data)