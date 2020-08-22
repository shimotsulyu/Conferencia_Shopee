#importa as bibliotecas
import pandas as pd
import glob
#cria uma lista com o nome de todos os arquivos csv na pasta "Arquivos"
endereco = glob.glob('Arquivos/_Mes__completo/*.csv')+glob.glob('Arquivos/*.csv')
print(len(endereco),endereco)
#lista com as colunas a serem utilizadas
target = ['ID do pedido','Hora completa do pedido','Status do pedido','Nome de usuário (comprador)','Número de produtos pedidos', 'Valor Total', 'Cupom do vendedor', 'Taxa de envio pagas pelo comprador']
#função de consulta de rendimento pelo ID
def consultaV(endereco,ID):
    ID = ID.upper()
    print('Procurando:',ID)
    empty = True
    for end in endereco:
        #cria um dataframe com do arquivo lido
        data = pd.read_csv(end,sep=';')
        #verifica se o arquivo contem o ID
        if len(data.loc[data['ID do pedido']==ID])==1:
            empty = False
            #Valor total do pedido
            VTotal = float(data.loc[data['ID do pedido']==ID]['Valor Total'])
            #Valor do cupom de desconto do vendendor
            CupomV = float(data.loc[data['ID do pedido']==ID]['Cupom do vendedor'])
            #Valor da taxa de envio paga pelo comprador
            Frete = float(data.loc[data['ID do pedido']==ID]['Taxa de envio pagas pelo comprador'])
            #Rendimento total
            total = round(VTotal - CupomV - Frete,2)
            #informações da consulta
            print('*****************************************************')
            print(end)
            print('ID:',ID)
            print(data.loc[data['ID do pedido']==ID]['Nome de usuário (comprador)'])
            print('Itens:',int(data.loc[data['ID do pedido']==ID]['Número de produtos pedidos']),'\n')
            print('_______________________________\nTotal do pedido:',VTotal)
            print('Cupon do vendedor:',-1*CupomV)
            print('Taxa de envio',-1*Frete)
            print('_______________________________\nRendimento de pedido:',total)
            print('\nFIM rendimento pedido\n*****************************************************\n\n')
    if empty is True:
        print('*****************************************************')
        print('ID não encontrado!')
        print('\nFIM rendimento pedido\n*****************************************************\n\n')
#função de soma de rendimentos
def somaV(data1,data2,text,target):
    data2 = data2.append(data1[target], ignore_index=True)
    data2.drop_duplicates()
    data2.dropna(axis='rows')
    VTotal = data2['Valor Total'].sum()
    CupomV = data2['Cupom do vendedor'].sum()
    Frete = data2['Taxa de envio pagas pelo comprador'].sum()
    total = VTotal - CupomV - Frete
    return data2, total
#função gerar rendimento
def gerarR(endereco,target):
    #cria dataframes pandas
    completo, enviando, enviar = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    #percorre os arquivos csv
    for end in endereco:
        print(end)
        #cria um dataframe com do arquivo lido
        data = pd.read_csv(end,sep=';')
        #variavel para determinar o tipo de venda
        status = end[15:21]
        #c para completo
        if status == 'comple':
            completo, Carteira = somaV(data,completo,'Carteira:',target)
        #s para enviando
        elif status == 'shippi':
            enviando, Liberar1 = somaV(data,enviando,'Enviando:',target)
        #t para enviar
        elif status == 'toship':
            enviar, Liberar2 = somaV(data,enviar,'Enviar:',target)
    #informações das vendas
    print('*****************************************************')
    print('Vendas completadas:',completo['ID do pedido'].count(),'/ Itens:',completo['Número de produtos pedidos'].sum())
    print('Vendas enviando:',enviando['ID do pedido'].count(),'/ Itens:',enviando['Número de produtos pedidos'].sum())
    print('Vendas enviar:',enviar['ID do pedido'].count(),'/ Itens:',enviar['Número de produtos pedidos'].sum())
    print('*****************************************************')
    print('Para Liberar:', Liberar1+Liberar2)
    print('Carteira Total:', round(Carteira))
    print('TOTAL:',Liberar1+Liberar2+Carteira)
    print('\nFIM rendimento total\n*****************************************************\n\n')
def gerarTabela(endereco,target):
    #cria dataframes pandas
    Main = pd.DataFrame()
    #percorre os arquivos csv
    for end in endereco:
        print(end)
        #cria um dataframe com do arquivo lido
        data = pd.read_csv(end,sep=';')
        print(len(data))
        Main = Main.append(data[target], ignore_index=True)
        Main.drop_duplicates()
        Main.dropna(axis='rows')
        print(len(Main))
    Main['Rendimento'] = Main['Valor Total'] - Main['Cupom do vendedor'] - Main['Taxa de envio pagas pelo comprador'] 
    return Main
def salvarTabela(endereco,target,nome_arquivo):
    out = gerarTabela(endereco,target)
    df1 = out[['ID do pedido','Status do pedido','Nome de usuário (comprador)','Rendimento']].dropna(axis='rows').reset_index(drop=True)
    writer = pd.ExcelWriter(nome_arquivo)
    df1.to_excel(writer, sheet_name = "Rendimento")
    writer.save()
    writer.close()
    print('Arquivo',nome_arquivo,'salvo com sucesso')
    return df1