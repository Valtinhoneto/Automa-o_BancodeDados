import pandas as pd
import win32com.client as win32

#Visualização do arquivo
tabela_vendas = pd.read_excel('Tabela_Dados.xlsx')
pd.set_option('display.max_columns', None)

print('='*50)
#Faturamento
Faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum().reset_index()
Faturamento = Faturamento.set_index('ID Loja')
print(Faturamento)


print('='*50)
#quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum().reset_index()
quantidade = quantidade.set_index('ID Loja')
print(quantidade)

print('='*50)
#ticket médio por produto em cada loja
ticket_medio = Faturamento['Valor Final'] / quantidade['Quantidade']
ticket_medio = ticket_medio.to_frame(name='Ticket Médio')
print (ticket_medio)

# enviar email

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'valterneto123456789@outlook.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>
   
<p>Faturamento:</p>
{Faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>Valter Neto</p>
'''

mail.Send()

print("email enviado")
