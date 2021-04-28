import pandas as pd

# Importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx') # O pandas vai ler um arquivo em excel e armazenar na variável tabela_vendas

# Visualizar a base de dados
pd.set_option('display.max_columns', None) # Mostra todos os dados da tabela
print(tabela_vendas)

# Faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum() # Filtra as colunas 'ID Loja' e 'Valor Final' da tabela, depois agrupa todas as lojas e soma o faturamento
print(faturamento)
# Quantidade de produtos vendidos por loja

# Ticket médio por produto em cada loja

# Enviar um email com o relatório