import pandas as pd
import win32com.client as win3

# importando base de dados
tabela_vendas = pd.read_excel('./Vendas.xlsx')

#colocar quando a tabela n mostrar tds as colunas ai aparece tds
# pd.set_option('display.max_columns',None)
print(tabela_vendas)

#consultando apenas as colunas qu eu quero
# tabela_vendas[['ID Loja', 'Valor Final']]

#agrupando
# tabela_vendas.groupby('ID Loja').sum()

#fazendo a soma
faturamento  = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja',).sum()
print(faturamento)

quantidade =tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-' *50)
ticket_medio = (faturamento['Valor Final']/ quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0:'Tiket medio'})
print(ticket_medio)

outlook =win32.dispatcher('outlok aplication')
mail = outlook.CreateItem(0)
mail.to = 'renata0cristiane@gmail.con'
mail.subject = 'relatorio de vendas'
mail.HTMLBoody =f'''
 <p>Prezado</p>

 <psegue as tabelas></p>

 <p>FAruramento</P>
{faturamento.to_html(formatters={'Valor Final': 'R$ (:,.2f)'.format})}

<p>Quantidade de vendas</P>
{quantidade.to_html()}

 <p>Ticked medio de produtos em cada loja</P>
{ticket_medio.to_html(formatters={'Ticket Medio': 'R$ (:,.2f)'.format})}

<p> qualquer duvida estou a disposicao, 
abs </P>
'''

mail.Sand()
