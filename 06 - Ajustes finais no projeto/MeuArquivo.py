import pandas as pd
import win32com.client as win32

# Importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx') # O pandas vai ler um arquivo em excel e armazenar na variável tabela_vendas

# Visualizar a base de dados
pd.set_option('display.max_columns', None) # Mostra todos os dados da tabela
print(tabela_vendas)

# Faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum() # Filtra as colunas 'ID Loja' e 'Valor Final' da tabela, depois agrupa todas as lojas e soma o faturamento
print(faturamento)

# Quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-' * 50)
# Ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame() # to_frame() - Transforma uma série de dados em uma tabela
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# Enviar um email com o relatório
outlook = win32.Dispatch('outlook.application') # O python se conecta com o email
mail = outlook.CreateItem(0)
mail.To = 'ivoneideduarte28@gmail.com' # Destinatário
mail.Subject = 'Relatório de Vendas por Loja' # Assunto
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja:</p>

<p>Faturamento: </p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Ivoneide</p>
'''

mail.Send() # Envia o email
print('Email Enviado')