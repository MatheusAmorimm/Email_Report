import pandas as pd
import win32com.client as win32

# importar a base
saleChart = pd.read_excel('Vendas (1).xlsx')
pd.set_option('display.max_columns', None)


# (filtrando as colunas necessarias)
profitSaleChart = saleChart[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()


qtdSaleChart = saleChart[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()


medSaleChart = (profitSaleChart['Valor Final'] / qtdSaleChart['Quantidade']).to_frame()
medSaleChart = medSaleChart.rename(columns={0: 'Ticket Médio'})


# # enviar um relatorio por email
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'ottavio@iacontabil.com.br'
mail.Subject = 'Relatório de Vendas - Matheus'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p> Segue o relatório de vendas detalhado de cada loja.</p>

<p style = 'font-weight: bold'>Faturamento:</p>
{profitSaleChart.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}
<br>
<p style = 'font-weight: bold'>Quantidade Vendida:</p>
{qtdSaleChart.to_html()}
<br>
<p style = 'font-weight: bold'>Preço Médio de Venda:</p>
{medSaleChart.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}
<br>
<p>Qualquer dúvida estou a Disposição!</p>
<p>Atenciosamente,</p>

<p>Matheus Amorim.</p>
'''
mail.Send()
