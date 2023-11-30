import pandas as pd
import win32com.client as win32


# 1 - importar base de dados

tabela_vendas = pd.read_excel("Vendas.xlsx")


# 2 - visualizar base de dados

pd.set_option("display.max_columns", None)


# 3 - calcular faturamento por loja

faturamento = tabela_vendas[["ID Loja", "Valor Final"]].groupby("ID Loja").sum()
print(faturamento)
print("-" * 50)


# 4 - quantidade de produtos vendidos por lojas

quantidade_de_produtos = tabela_vendas[["ID Loja", "Quantidade"]].groupby("ID Loja").sum()
print(quantidade_de_produtos)
print("-" * 50)

# 5 - Ticket medio por produto de cada loja 

ticket_medio = (faturamento["Valor Final"] / quantidade_de_produtos["Quantidade"]).to_frame()
ticket_medio = ticket_medio.rename(columns={0: "Ticket Médio"})
print(ticket_medio)

# 5 - fazer o envio de email automatico

outlook = win32.Dispatch("outlook.application")
mail = outlook.CreateItem(0)
mail.To = "gabrielbottan0@gmail.com"
mail.subject = "Relatório de vendas por Loja"
mail.HTMLBody = f''' <p>Prezados,</p>
<p>Segue o relatório de vendas por cada loja.</p>
<p>Faturamento: </p>
{faturamento.to_html(formatters={"Valor Final":"R${:,.2f}".format})}. </

<p>Quantidade Vendida:</p>
{quantidade_de_produtos.to_html()}

<p>Ticket médio dos produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={"Ticket Médio":"R${:,.2f}".format})}


<p>Qualquer dúvida estou a disposição.</p>

  '''

mail.Send()