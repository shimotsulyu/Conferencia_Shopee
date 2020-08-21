import pandas as pd
import glob
endereco = glob.glob('Arquivos/*.csv')
target = ['ID do pedido','Nome de usuário (comprador)','Número de produtos pedidos', 'Valor Total', 'Cupom do vendedor', 'Taxa de envio pagas pelo comprador']
def calc(data1,data2,text):
        data2 = data2.append(data1[target], ignore_index=True)
        data2.drop_duplicates()
        VTotal = data2['Valor Total'].sum()
        CupomV = data2['Cupom do vendedor'].sum()
        Frete = data2['Taxa de envio pagas pelo comprador'].sum()
        total = VTotal - CupomV - Frete
        return data2, total
completo, enviando, enviar = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
for end in endereco:
    data = pd.read_csv(end,sep=';')
    status = end[15]
    if status == 'c':
        completo, Carteira = calc(data,completo,'Carteira:')
    elif status == 's':
        enviando, Liberar1 = calc(data,enviando,'Enviando:')
    elif status == 't':
        enviar, Liberar2 = calc(data,enviar,'Enviar:')
print('*****************************************************')
print('Vendas completadas:',completo['ID do pedido'].count(),'/ Itens:',completo['Número de produtos pedidos'].sum())
print('Vendas enviando:',enviando['ID do pedido'].count(),'/ Itens:',enviando['Número de produtos pedidos'].sum())
print('Vendas enviar:',enviar['ID do pedido'].count(),'/ Itens:',enviar['Número de produtos pedidos'].sum())
print('*****************************************************')
print('Para Liberar:', Liberar1+Liberar2)
print('Carteira Total:', Carteira)
print('TOTAL:',Liberar1+Liberar2+Carteira)