#importa as bibliotecas
import pandas as pd
import glob
#cria uma lista com o nome de todos os arquivos csv na pasta "Arquivos"
endereco = glob.glob('Arquivos/_Mes__completo/*.xls')+glob.glob('Arquivos/*.xls')
print(len(endereco),endereco,'\n')
#lista com as colunas a serem utilizadas
target = ['ID do pedido','Data de criação do pedido','Status do pedido','Nome de usuário (comprador)','Número de produtos pedidos', 'Valor Total', 'Cupom do vendedor', 'Taxa de envio pagas pelo comprador']
#função ler arquivos
def lerArquivos():
    #cria dataframes pandas
    Main = pd.DataFrame()
    #percorre os arquivos csv
    for end in endereco:
        #cria um dataframe com do arquivo lido
        data = pd.read_excel(end,sep=';')
        Main = Main.append(data, ignore_index=True)
    return Main
#função de consulta de rendimento pelo ID
def consultaV(ID,saida=False):
    ID = str(ID.upper())
    print('Procurando:',ID)
    empty = True
    data = lerArquivos()
    data = data.loc[data['ID do pedido']==ID]
    if len(data)>0:
        print(data['Nome de usuário (comprador)'].to_string(index=False))
        dados = data[['Nome de usuário (comprador)', 'Nome do destinatário', 'Telefone','Endereço de entrega', 'Cidade', 'Bairro', 'Cidade.1', 'UF', 'País','CEP']]
        dados['Telefone'] = dados['Telefone'].astype(int)
        total = data['Valor Total'] - data['Cupom do vendedor'] - data['Taxa de envio pagas pelo comprador']
        status = data[['Status do pedido', 'Status da Devolução / Reembolso']].to_numpy()[0]
        if saida is True:
            return total, status, dados, data
        else:
            print('Status:',status)
            print('Rendimento:',total.to_string(index=False))
    else:
        print('ID não encontrado!')
#função de soma de rendimentos
def somaR(data,tipo):
    data = data[target].loc[data['Status do pedido']==tipo]
    vendas = data['ID do pedido'].count()
    VTotal = data['Valor Total'].sum()
    CupomV = data['Cupom do vendedor'].sum()
    Frete = data['Taxa de envio pagas pelo comprador'].sum()
    total = VTotal - CupomV - Frete
    return total, vendas
#função para verificar mÊs
def mesR(end):
    if end[-8:-6] == '01':
        return 'Janeiro'
    elif end[-8:-6] == '02':
        return 'Fevereiro'
    elif end[-8:-6] == '03':
        return 'Março'
    elif end[-8:-6] == '04':
        return 'Abril'
    elif end[-8:-6] == '05':
        return 'Maio'
    elif end[-8:-6] == '06':
        return 'Junho'
    elif end[-8:-6] == '07':
        return 'Julho'
    elif end[-8:-6] == '08':
        return 'Agosto'
    elif end[-8:-6] == '09':
        return 'Stembro'
    elif end[-8:-6] == '10':
        return 'Outubro'
    elif end[-8:-6] == '11':
        return 'Novembro'
    elif end[-8:-6] == '12':
        return 'Dezembro'
    else:
        return 'Mes não identificado'
import numpy as np
#função gerar rendimento
def gerarR():
    #percorre os arquivos xls
    Total = 0
    for end in endereco:
        data = pd.read_excel(end,sep=';')
        data.dropna(subset=['Status do pedido'])
        mes = mesR(end).upper()
        print('__________________________#\n',mes)
        tipo = data['Status do pedido'].unique()
        TotalMes, Npago = 0, 0
        for t in tipo:
            if str(t)!='nan':
                total, vendas = somaR(data,t)
                print('    ',vendas,t,':',total)
                if t!='Cancelado':
                    if t!='Não pago':
                        TotalMes = TotalMes+total
        Total = Total+TotalMes
        print('__________________________\nRendimento',mes,TotalMes)
        print('__________________________#')
    print('____________________________________________________\nRendimento Total:',Total) 
#função para salvar tabela xls
def salvarTab(data,nome_arquivo,planilha):
    nome_arquivo = nome_arquivo+'.xls'
    writer = pd.ExcelWriter(nome_arquivo)
    data.to_excel(writer, sheet_name = planilha)
    writer.save()
    writer.close()
    print('Arquivo',nome_arquivo,'salvo com sucesso')
#função para gerar tabela xls
def gerarTab(nome_arquivo,saida=False):
    #cria dataframes pandas
    Main = pd.DataFrame()
    #percorre os arquivos csv
    Main = lerArquivos()
    Main['Rendimento'] = Main['Valor Total'] - Main['Cupom do vendedor'] - Main['Taxa de envio pagas pelo comprador'] 
    Main = Main[['Data de criação do pedido','ID do pedido','Status do pedido','Nome de usuário (comprador)','Rendimento']]
    Main = Main.loc[(Main['Status do pedido']=='Completo') | (Main['Status do pedido']=='Cancelado') | (Main['Status do pedido']=='Frete') | (Main['Status do pedido']=='Não pago') | (Main['Status do pedido']=='A Enviar')]
    print('\n____________________________________________________')
    print('DADOS DE VENDAS\n')
    for status in Main['Status do pedido'].unique():
        print(status,round(Main.loc[Main['Status do pedido']==status,'Rendimento'].sum(),2))
    print('____________________________________________________')
    salvarTab(Main,nome_arquivo,'Rendimento')
    if saida is True:
        return Main
#função para visualizar rank
def rank(tipo):
    Main = lerArquivos()
    Main = Main.loc[(Main['Status do pedido']==tipo)]
    compradores = Main['Nome de usuário (comprador)'].value_counts()
    print(compradores)
    salvarTab(compradores,tipo,'planilha')