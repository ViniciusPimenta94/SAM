import sys
from PyQt5 import uic, QtWidgets
from SAM_2 import verificação


class load_arquivos:
        def __init__(self):
            self.arquivo_fornecedor = []
            self.palavra_chave = []
            self.list_fornecedor = []

        def ler_arquivo_fornecedor(self):
            
            fornecedor = QtWidgets.QFileDialog.getOpenFileName()[0]
            with open (fornecedor, 'r') as a:
                self.arquivo_fornecedor = a.name
            
            self.list_fornecedor.append(fornecedor)
            
            print('Load file fornecedor complete!')
            print('list_fornecedor:', self.list_fornecedor)

        def Identificar_Cliente(self):
            if self.palavra_chave[:3] == 'PUC':
                 self.Cliente = 'PUC'
                 print('Cliente = ',self.Cliente)
                
            elif self.palavra_chave[:3] == 'LOC':
                self.Cliente = 'LOCALIZA'
                print('Cliente = ',self.Cliente)
            
            elif self.palavra_chave[:3] == 'RIA':
                self.Cliente = 'RIACHUELO'
                print('Cliente = ',self.Cliente)
            
            elif self.palavra_chave[:3] == 'FLE':
                self.Cliente = 'FLEURY'
                print('Cliente = ',self.Cliente)
                
            elif self.palavra_chave[:3] == 'DAK':
                 self.Cliente = 'DAKI'
                 print('Cliente = ',self.Cliente)
        
            else: 
                print('Cliente não identificado!!')
        
        def executar_SAM2(self):
            print('Executando SAM_2...')
            palavra_chave = tela.lineEdit.text()
            self.palavra_chave = palavra_chave
            
            load_active.Identificar_Cliente()
            
            fornecedor_x = 0 
            while fornecedor_x < len(self.list_fornecedor):
                print("Fazendo....>> ", self.list_fornecedor[fornecedor_x])
            #try:
                chama_SAM2 = verificação(self.list_fornecedor[fornecedor_x], self.palavra_chave, self.Cliente)
                chama_SAM2.criar_new_sheet()
                chama_SAM2.count_ws()
                chama_SAM2.File_name()      
                chama_SAM2.count_new_ws()
                chama_SAM2.tabela_dinâmica_new_sheet()
                chama_SAM2.comparar_valores()
            #except:
                print("Erro ao calcular >> ", self.list_fornecedor[fornecedor_x])                   
                
                fornecedor_x+=1
            

app = QtWidgets.QApplication([])
tela = uic.loadUi("Interface_SAM.ui")

tela.show()

load_active = load_arquivos()
tela.pushButton.clicked.connect(load_active.ler_arquivo_fornecedor)
tela.pushButton_4.clicked.connect(load_active.executar_SAM2)

sys.exit(app.exec_())