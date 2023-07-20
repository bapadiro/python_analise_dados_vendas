# ANALISE DE DADOS
# 1: importar a base de dados
# 2: visualizar a base de dados X fazer tratamento
# 3: calcular faturamento por loja
# 4: calcular qnt de produtos vendido por loja
# 5: calcular ticket médio por produto em cada loja: faturamento / quantidade
# 6: enviar e-mail com relatório

import pandas as pd
import win32com.client as win32

table_vendas = pd.read_excel('Vendas.xlsx')
pd.set_option('display.max_columns', None) #visualizar todas as colunas da base de dados: opção x valor
#print(table_vendas)

# 3:
# para filtrar basta chamar a variavel + ['nm_coluna'] [['nm_coluna se for mais de uma coluna']]
# para agrupar alguma coluna basta chamar variavel [['nm_coluna']].groupby('nm_voluna) ons: pode ser mais de uma e para fazer alguma ação com a segunda coluna, basta colocar o .sum() por exemplo
faturamento = table_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

#4:
qnt_produto_loja = table_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(qnt_produto_loja)

print('-' * 50)

#5
ticket_medio = (faturamento['Valor Final'] / qnt_produto_loja['Quantidade']).to_frame() #().to_frame : transforma em table
ticket_medio = ticket_medio.rename(columns={0:'Ticke Medio'}) #renomeando o nome da coluna
print(ticket_medio)

#6 
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'apadiroba@gmail.com'
mail.Subject = 'Relatório Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue relatório de vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade de Vendas por Loja:</p>
{qnt_produto_loja.to_html()}

<p>Ticket Medio dos produtos por Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Medio': 'R${:,.2f}'.format})}

<p>permaneço à disposição</p>

abs, 
Bárbara Diogo
'''

mail.Send()





