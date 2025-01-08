from itertools import groupby

import pandas as pd #Importando a biblioteca pandas
import win32com.client as win32 #Biblioteca para utilizar o outlook do pc

# importar a base de dados

tabela_vendas = pd.read_excel('Vendas.xlsx')


# Visualizar base de dados

pd.set_option('display.max_columns', None) # Dessa forma, o python irá ler todas as colunas do banco de dados.
# Caso queira filtrar menos que o maximo, basta colocar o numero de colunas no lugar do None
print(tabela_vendas)

# faturamento por loja

faturamento_por_loja = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum() # Agrupa de acordo com a ID e mostra o valor final.

print(faturamento_por_loja)

# Quantidade de produtos vendidos por loja

quantidade_por_loja = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

print(quantidade_por_loja)
print('-' * 50)
# Ticket médio por loja
ticket_medio = (faturamento_por_loja['Valor Final'] / quantidade_por_loja['Quantidade']).to_frame() # to_frame transforme esse item em uma tabela
# ticket_medio = ticket_medio.round(2) # Garante que os valores sejam arredondados para duas casas decimais
ticket_medio = ticket_medio.rename(columns = {0: 'Ticket Médio'})
print(ticket_medio)


# enviar um email com o relatório


outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = "danielcostaroque@gmail.com"
mail.Subject = "E-mail automático do Python - Relatorio de vendas por loja"
mail.HTMLBody =f'''
<p>Prezados,</p>

<p>Segue o relatório de vendas em anexo.</p>

<p>Faturamento:</p>
{faturamento_por_loja.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade vendida:</p>
{quantidade_por_loja.to_html(formatters={'Quantidade': lambda x: f"{x} unidades"})}

<p>Ticket médio dos produtos:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou a disposição.</p>
<p>Att. Daniel</p>

'''

# formatters, gera um dicionario dentro do html
mail.Send()
print('Email enviado com sucesso')