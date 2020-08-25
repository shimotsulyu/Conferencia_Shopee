#importa as bibliotecas
import tkinter as tk
import pandas as pd
import glob

class Application(tk.Frame):
    #cria uma lista com o nome de todos os arquivos csv na pasta "Arquivos"
    endereco = glob.glob('Arquivos/_Mes__completo/*.xls')+glob.glob('Arquivos/*.xls')
    #lista com as colunas a serem utilizadas
    target = ['ID do pedido','Data de criação do pedido',
              'Status do pedido','Nome de usuário (comprador)',
              'Número de produtos pedidos', 'Valor Total', 
              'Cupom do vendedor', 'Taxa de envio pagas pelo comprador']
    #função ler arquivos
    def lerArquivos(self):
        #cria dataframes pandas
        Main = pd.DataFrame()
        #percorre os arquivos csv
        for end in self.endereco:
            #cria um dataframe com do arquivo lido
            data = pd.read_excel(end,sep=';')
            Main = Main.append(data, ignore_index=True)
        return Main
    #função de consulta de rendimento pelo ID
    def consultaV(self,saida=False):
        ID = self.EntryBox.get()
        self.clearEntry()
        self.janela('\n______________________________#\nProdurando por '+str(ID))
        empty = True
        data = self.lerArquivos()
        data = data.loc[data['ID do pedido']==ID]
        if len(data)>0:
            self.janela('\nNome: '+data['Nome de usuário (comprador)'].to_string(index=False))
            dados = data[['Nome de usuário (comprador)', 'Nome do destinatário', 'Telefone','Endereço de entrega', 'Cidade', 'Bairro', 'Cidade.1', 'UF', 'País','CEP']]
            dados['Telefone'] = dados['Telefone'].astype(int)
            total = data['Valor Total'] - data['Cupom do vendedor'] - data['Taxa de envio pagas pelo comprador']
            status = data[['Status do pedido', 'Status da Devolução / Reembolso']].to_numpy()[0]
            if saida is True:
                return total, status, dados, data
            else:
                self.janela('\nStatus: '+str(status))
                self.janela('\nRendimento: '+total.to_string(index=False)+'\n______________________________#')
        else:
            self.janela('\n**ERRO! ID não encontrado!\n______________________________#')
    #função de soma de rendimentos
    def somaR(self,data,tipo):
        data = data[self.target].loc[data['Status do pedido']==tipo]
        vendas = data['ID do pedido'].count()
        VTotal = data['Valor Total'].sum()
        CupomV = data['Cupom do vendedor'].sum()
        Frete = data['Taxa de envio pagas pelo comprador'].sum()
        total = VTotal - CupomV - Frete
        return total, vendas
    #função para verificar mÊs
    def mesR(self,end):
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
    #função gerar rendimento
    def gerarR(self):
        #percorre os arquivos xls
        Total = 0
        for end in self.endereco:
            data = pd.read_excel(end,sep=';')
            data.dropna(subset=['Status do pedido'])
            mes = self.mesR(end).upper()
            self.janela('\n______________________________#\n'+str(mes))
            tipo = data['Status do pedido'].unique()
            TotalMes, Npago = 0, 0
            for t in tipo:
                if str(t)!='nan':
                    total, vendas = self.somaR(data,t)
                    self.janela('\n    '+str(vendas)+' '+str(t)+':'+str(total))
                    if t!='Cancelado':
                        if t!='Não pago':
                            TotalMes = TotalMes+total
            Total = Total+TotalMes
            self.janela(Total)
            self.janela('\n__________________________\nRendimento '+str(mes)+' '+str(TotalMes))
            self.janela('\n______________________________#')
        self.janela('\n______________________________\nRendimento Total: '+str(Total)) 
    #função para salvar tabela xls
    def salvarTab(self,data,nome_arquivo,planilha):
        nome_arquivo = nome_arquivo+'.xls'
        writer = pd.ExcelWriter(nome_arquivo)
        data.to_excel(writer, sheet_name = planilha)
        writer.save()
        writer.close()
        self.janela('\n\n'+str(nome_arquivo)+' salvo com sucesso\n______________________________#')
    #função para gerar tabela xls
    def gerarTab(self,saida=False):
        nome_arquivo = 'Rendimento_Total'
        #cria dataframes pandas
        Main = pd.DataFrame()
        #percorre os arquivos csv
        Main = self.lerArquivos()
        Main['Rendimento'] = Main['Valor Total'] - Main['Cupom do vendedor'] - Main['Taxa de envio pagas pelo comprador'] 
        Main = Main[['Data de criação do pedido','ID do pedido','Status do pedido','Nome de usuário (comprador)','Rendimento']]
        Main = Main.loc[(Main['Status do pedido']=='Completo') | (Main['Status do pedido']=='Cancelado') | (Main['Status do pedido']=='Frete') | (Main['Status do pedido']=='Não pago') | (Main['Status do pedido']=='A Enviar')]
        self.janela('\n______________________________#')
        self.janela('\nDADOS DE VENDAS')
        for status in Main['Status do pedido'].unique():
            self.janela('\n'+str(status)+' '+str(round(Main.loc[Main['Status do pedido']==status,'Rendimento'].sum(),2)))
        self.salvarTab(Main,nome_arquivo,'Rendimento')
        if saida is True:
            return Main
    #função para visualizar rank
    def rank(self):
        tipo = self.EntryBox.get()
        self.clearEntry()
        n = 10
        self.janela('\n______________________________#')
        self.janela('\nTop '+str(n)+' clientes '+ tipo)
        Main = self.lerArquivos()
        Main = Main.loc[(Main['Status do pedido']==tipo)]
        vendas = Main['Nome de usuário (comprador)'].value_counts()
        if len(vendas)==0:
            self.janela('\n**ERRO - Opções aceitas (Completo, Frete, A Enviar, Não pago, Cancelado)')
        else:
            self.janela('\n\n'+str(vendas.head(n)))
            self.salvarTab(vendas,tipo,'planilha')

    def botoesmenu(self):
        self.labelEntry = tk.Label(self, text='COMANDO:', 
                                   font=self.fonte, bg='#87CEFA',width=10)
        self.labelEntry.grid(row=65,column=1)
        
        self.EntryBox = tk.Entry(self, 
                                 font=self.fonte, width=100)
        self.EntryBox.grid(row=65,column=2)
        
        self.CP = tk.Button(self, text='CONSULTA PEDIDO', 
                            font=self.fonte, width=25,
                            command=self.consultaV)
        self.CP.grid(row=1,column=0)
        
        self.GRM = tk.Button(self, text='RENDIMENTO MENSAL', 
                             font=self.fonte, width=25, 
                             command=self.gerarR)
        self.GRM.grid(row=2,column=0)

        self.GRT = tk.Button(self, text='RENDIMENTO TOTAL', 
                             font=self.fonte, width=25, 
                             command=self.gerarTab)
        self.GRT.grid(row=3,column=0)

        self.RC = tk.Button(self, text='RANK CLIENTES',
                            font=self.fonte, width=25, 
                            command=self.rank)
        self.RC.grid(row=4,column=0)
        
        self.Clear = tk.Button(self, text='APAGAR DISPLAY', fg='#FF6347',
                            font=self.fonte, width=25, 
                            command=self.clearD)
        self.Clear.grid(row=5,column=0)
        
        self.quit = tk.Button(self, text='QUIT', bg='#FF6347', 
                              font=self.fonte, width=25, 
                              command=self.master.destroy)
        self.quit.grid(row=64,column=0)
        
    def janela(self,texto):
        self.display.insert('end',texto, 'justified')
        #self.display.tag_config('start',fg='red')
    def clearD(self):
        self.clearEntry()
        self.display.delete('5.0','end')
    def clearEntry(self):
        self.EntryBox.delete('0','end')
        
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        #self.fonte = ('Verdana', '14', 'normal', 'bold')
        self.fonte = ('Verdana', '14')
        self.pack()
        self.botoesmenu()
        self.display = tk.Text(self,font=self.fonte, width=110, height=50)
        self.display.grid(row=0,column=1,rowspan=65, columnspan=2)
        self.janela('Carregando arquivos')
        if len(self.endereco)==0:
            self.janela('\n**ERRO - A pasta está vazia!')
        else:
            for end in self.endereco:
                self.janela('\n'+end)
        self.janela('\n____________________________________________________________')
    
root = tk.Tk()
root.title("Conferencia Shopee Alpha test 1.0")
root.geometry("{0}x{1}+0+0".format(root.winfo_screenwidth(), root.winfo_screenheight()))
app = Application(master=root)
app.mainloop()