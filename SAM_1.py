import sys
from PyQt5 import uic, QtWidgets
from PyQt5.QtWidgets import QCompleter
from SAM_2 import verificação
from SAM_4_docparser import Parseamento_docparser
from pathlib import Path
import openpyxl, zipfile, os, shutil

class load_arquivos:
        def __init__(self):
            self.arquivo_fornecedor = []
            self.palavra_chave = []
            self.list_fornecedor = []
            self.nome_fornecedor = []
            
        def extrair_arquivo_fornecedor(self):
            rar = QtWidgets.QFileDialog.getOpenFileNames()[0]
            path = Path('./')
            path.mkdir(parents=True, exist_ok=True)
            for i in rar:
                with zipfile.ZipFile(i, 'r') as zip_ref:
                    zip_ref.extractall('./PDF')
                    
            print('Arquivos extraídos para a pasta PDF')
                    
        def ler_arquivo_fornecedor(self):
            self.fornecedor = QtWidgets.QFileDialog.getOpenFileNames()[0]
            print('Arquivos importados:', len(self.fornecedor))
            
            for arquivo in self.fornecedor:
                self.list_fornecedor.append(arquivo)
            print('Load file fornecedor complete!')
        
        def Identificar_Cliente(self):
            
            lista_clientes = ['LOCALIZA', 'FLEURY', 'RIACHUELO', 'PUC', 'DAKI', 'PAGUE_MENOS', 'DPSP', 'GRPCOM']
            self.Cliente = None
            
            for i, v in enumerate(lista_clientes):
                if lista_clientes[i][:3] == self.palavra_chave[:3]:
                    self.Cliente = v
                    print('Cliente = ',self.Cliente)

            if self.Cliente == None:
                print('Cliente não identificado!!')
                
            path = Path(f'Consolidados\{self.Cliente}')
            path.mkdir(parents=True, exist_ok=True)
            path2 = Path(f'Verticais\{self.Cliente}')
            path2.mkdir(parents=True, exist_ok=True)
      
        def executar_SAM4(self):
            self.nova_list_fornecedor = []
            for nome_arquivo in self.list_fornecedor:
                file_oldname = os.path.join(nome_arquivo)
                
                novo_nome = nome_arquivo[:-4] + self.palavra_chave + '.pdf'
                self.nova_list_fornecedor.append(novo_nome)
                
                file_newname_newfile = os.path.join(novo_nome)
                os.rename(file_oldname, file_newname_newfile)
                
            self.list_fornecedor = self.nova_list_fornecedor
            
            nome_fornecedor = self.nome_fornecedor
            arquivo_pdf = self.list_fornecedor
            
            print(f'{nome_fornecedor}\n')
            
            chama_SAM4 = Parseamento_docparser(nome_fornecedor, arquivo_pdf)
            self.list_json_parseado = chama_SAM4.importar_pdf()
         
            load_active.json_for_excel()
            
        def json_for_excel(self):
            # print(self.nome_fornecedor.upper())
            
            if 'AGUA' in self.nome_fornecedor.upper():
                vertical = 'AGUA'
            elif 'ENERGIA' in self.nome_fornecedor.upper():
                vertical = 'ENERGIA'
            
            if vertical == 'AGUA':  
                self.wb_vertical = openpyxl.load_workbook('./arquivo_cliente/vertical_agua.xlsx')
                self.ws_vertical = self.wb_vertical['Worksheet']
                
                coluna_faturados = 'valores_faturados'
                
            elif vertical == 'ENERGIA': 
                self.wb_vertical = openpyxl.load_workbook('./arquivo_cliente/vertical_energia.xlsx')
                self.ws_vertical = self.wb_vertical['Worksheet']
                
                coluna_faturados = 'valores_faturados_auditoria'
            
            is_data = True
            count_col_vertical = 0
            while is_data:
                count_col_vertical += 1
                data =  self.ws_vertical.cell (row = 1, column = count_col_vertical).value
                if data == None:
                    is_data = False
            
            row_vertical = 2
            faturas = []
           
            try: 
                for json_parseado in self.list_json_parseado:
                    json_parseado = json_parseado[0]
                    print(json_parseado['file_name'])
                    faturas.append(json_parseado['file_name'])
                   
                    if type(json_parseado[coluna_faturados]) == list:
                        quant_pdf = len(json_parseado[coluna_faturados])
                    else:
                        quant_pdf = 1
                    
                    for num_descricao in range(quant_pdf):
                        json_valores_faturados = json_parseado[coluna_faturados]
            
                        if type(json_parseado[coluna_faturados]) == list:
                            json_valores_faturados = json_valores_faturados[num_descricao]
                        
                        for j in range(count_col_vertical-1):
                            j+=1
                            column_vertical = self.ws_vertical.cell (row = 1, column = j).value
                            
                            if column_vertical.lower() == 'valores_faturados valor':
                               column_vertical = 'valor'
                            
                            if column_vertical.lower() == 'valores_faturados descricao':
                               column_vertical = 'descricao'
                             
                            if column_vertical.lower() == 'valores_faturados_auditoria descricao':
                               column_vertical = 'descricao'
                            
                            if column_vertical.lower() == 'valores_faturados_auditoria quantidade':
                               column_vertical = 'quantidade'
                               
                            if column_vertical.lower() == 'valores_faturados_auditoria tarifa preco':
                               column_vertical = 'tarifa_preco'
                            
                            if column_vertical.lower() == 'valores_faturados_auditoria valor':
                               column_vertical = 'valor'
                              
                            if json_parseado == None:
                                json_parseado == {}
                            
                            if json_valores_faturados == None:
                                json_valores_faturados == {}
                            
                            if type(json_parseado) != dict:
                                json_parseado == {}
                                print('aviso: json com problema')
                                
                            if type(json_valores_faturados) != dict:  
                                json_valores_faturados == {}
                                print('aviso: json com problema valores faturados')
                            
                            if column_vertical.lower() in json_parseado.keys():
                                    valor_parseado = json_parseado[column_vertical.lower()]
                        
                                    if vertical == 'AGUA':
                                        if column_vertical.lower() == 'valores_consumo':
                                           json_valor = json_parseado[column_vertical.lower()]
                                           
                                           if type(json_valor) == list:
                                               json_valor = json_valor[0]
                                          
                                           if json_valor == None:
                                               valor_parseado = 0
                                           
                                           elif type(json_valor) == int or type(json_valor) == float or type(json_valor) == str:
                                               valor_parseado = json_valor
                                           
                                           else:
                                               valor_parseado = json_valor['valor']
                                    
                                    if type(valor_parseado) == list:
                                        valor_parseado = valor_parseado[0]
                                    
                                    self.ws_vertical.cell (row = row_vertical, column = j).value = valor_parseado
    
                            elif column_vertical.lower() in json_valores_faturados.keys():
                                    
                                    valor_faturado = json_valores_faturados[column_vertical.lower()]
                                    
                                    self.ws_vertical.cell (row = row_vertical, column = j).value = valor_faturado
                        
                        row_vertical += 1
                        
            except:
                    print('Erro: Ou a fatura foi deletada no Parser ou ainda está em fila de processamento.')
                                
            if len(self.fornecedor) != len(faturas):
                print('\n\nCONSOLIDADO INCOMPLETO\n'*15)
            
            self.wb_vertical.save(f'Verticais\{self.Cliente}/______vertical_' + self.Cliente + '__' + self.nome_fornecedor+'.xlsx')
            # print('Save file >>>>>', '______vertical_' + self.Cliente + '__' + self.nome_fornecedor+'.xlsx')
                        
            self.list_fornecedor = [(f'Verticais\{self.Cliente}/______vertical_' + self.Cliente + '__' + self.nome_fornecedor+'.xlsx')]
            
            try:
                shutil.rmtree('./PDF')
            except OSError as e:
                print(e)
        
        def executar_SAM2(self):
            print('Executando SAM_2...')
            palavra_chave = tela.lineEdit.text()
            self.palavra_chave = palavra_chave
            
            nome_fornecedor = tela.lineEdit_2.text()
            self.nome_fornecedor = nome_fornecedor
            
            load_active.Identificar_Cliente()
            load_active.executar_SAM4()
            
            fornecedor_x = 0 
            while fornecedor_x < len(self.list_fornecedor):
                print("Fazendo....>> ", self.list_fornecedor[fornecedor_x])
            
                chama_SAM2 = verificação(self.list_fornecedor[fornecedor_x], self.palavra_chave, self.Cliente)
                chama_SAM2.criar_new_sheet()
                chama_SAM2.count_ws()
                chama_SAM2.File_name()      
                chama_SAM2.count_new_ws()
                chama_SAM2.tabela_dinâmica_new_sheet()
                chama_SAM2.comparar_valores()
            
                fornecedor_x+=1

app = QtWidgets.QApplication([])
tela = uic.loadUi("./assets/Interface_SAM.ui")

tela.show()

lista_parser = ['Agua_Aegea', 'Agua_Aegea_Seta_Vermelha', 'Agua_Aegea_Seta_Vermelha_testes', 'Agua_Aguas_Arenapolis', 'Agua_AguasDoBrasil', 'Agua_BRK', 'Agua_CAERR', 'Agua_CAESA', 'Agua_Caesb', 'Agua_Cagece', 'Agua_Casan', 'Agua_CEDAE', 'Agua_Cesama', 'Agua_Cesama_teste', 'Agua_Cesan', 'Agua_CIS', 'Agua_Codau', 'Agua_COGERH', 'Agua_Cohasb', 'Agua_Compesa', 'Agua_Comusa', 'Agua_Copasa', 'Agua_Corsan', 'Agua_Cosanpa', 'Agua_DAAE_Araraquara', 'Agua_DAAE_Rio_Claro', 'Agua_DAE_Americana', 'Agua_DAE_Bauru', 'Agua_DAE_Jundiai', 'Agua_DAE_Santa_Barbara', 'Agua_DAE_Santana_do_Livramento', 'Agua_DAEM', 'Agua_DAEP', 'Agua_DAES_Juina', 'Agua_DAEV', 'Agua_Damae_SaoJoaoDelRei', 'Agua_DEMAE_Campo_Belo', 'Agua_Departamento_Divisao_Servico', 'Agua_Depasa', 'Agua_DMAE_Poços_de_Caldas', 'Agua_DMAE_Poços_de_Caldas_copy1', 'Agua_DMAE_Porto_Alegre', 'Agua_DMAE_Uberlandia', 'Agua_Elipse', 'Agua_Emasa', 'Agua_Embasa', 'Agua_Embasa_testes', 'Agua_EMDAEP', 'Agua_Iguasa', 'Agua_Linhas', 'Agua_Prefeitura_Itirapina', 'Agua_Prefeituras', 'Agua_Prolagos', 'Agua_SAAE_Amparo', 'Agua_SAAE_Atibaia', 'Agua_SAAE_Bacabal', 'Agua_SAAE_Balsas', 'Agua_SAAE_Barra_Mansa', 'Agua_SAAE_Bebedouro', 'Agua_SAAE_Campo_Maior', 'Agua_SAAE_Canaa_Dos_Acarajas', 'Agua_SAAE_Capivari', 'Agua_SAAE_Catu', 'Agua_SAAE_Caxias', 'Agua_SAAE_Ceara_Mirim', 'Agua_SAAE_Codo', 'Agua_SAAE_Estancia', 'Agua_SAAE_Garca', 'Agua_SAAE_Governador_Valadares', 'Agua_SAAE_Grajau', 'Agua_SAAE_Ibitinga', 'Agua_SAAE_Iguatu', 'Agua_SAAE_Indaiatuba', 'Agua_SAAE_Itapetinga', 'Agua_SAAE_Itapira', 'Agua_SAAE_Jacareí', 'Agua_SAAE_Juazeiro', 'Agua_SAAE_LençoisPaulistas', 'Agua_SAAE_Limoeiro', 'Agua_SAAE_Linhares', 'Agua_SAAE_Mogi_Mirim', 'Agua_SAAE_Morada_Nova', 'Agua_SAAE_Parauapebas', 'Agua_SAAE_Parintins', 'Agua_SAAE_PDFSimples', 'Agua_SAAE_Penedo', 'Agua_SAAE_Quadrado', 'Agua_SAAE_Quadradoteste', 'Agua_SAAE_Quixeramobim', 'Agua_SAAE_Sao_Carlos', 'Agua_SAAE_Sobral', 'Agua_SAAE_Volta_Redonda', 'Agua_SAAEJ_Jaboticabal', 'Agua_SAAEP', 'Agua_SAAETRI', 'Agua_Sabesp', 'Agua_SAE_Araguari', 'Agua_SAE_Ituiutaba', 'Agua_Saec', 'Agua_Saecil', 'Agua_SaemaAraras', 'Agua_SAEP', 'Agua_SAESA', 'Agua_SAEV', 'Agua_SAMAE_Blumenau', 'Agua_SAMAE_Brusque', 'Agua_SAMAE_Caxias', 'Agua_SAMAE_Jaragua_do_Sul', 'Agua_Samae_Mogi_Guaçu', 'Agua_Sanasa', 'Agua_Saneago', 'Agua_Saneamento', 'Agua_Sanebavi', 'Agua_SANEP', 'Agua_Sanepar', 'Agua_Saneparteste', 'Agua_Sanesalto', 'Agua_SANESUL', 'Agua_SEMAE_Ararangua', 'Agua_Semae_Mogi_das_Cruzes', 'Agua_SEMAE_Piracicaba', 'Agua_SEMAE_São_José_do_Rio_Preto', 'Agua_SEMAE_Sao_Leopoldo', 'Agua_Semasa_Itajaí', 'Agua_Semasa_Santo_Andre', 'Agua_SESAN', 'Agua_Setae', 'Energia_2W', 'Energia_Aliança', 'Energia_AmazonasEnergia', 'Energia_Banco_BTG', 'Energia_CEA', 'Energia_CEB', 'Energia_CEEE', 'Energia_CELESC', 'Energia_CELESC_copy', 'Energia_Cemig', 'Energia_CemigSim', 'Energia_Cerbanorte', 'Energia_Cercar', 'Energia_CERCI', 'Energia_Cergal', 'Energia_CERMOFUL', 'Energia_CERRP', 'Energia_CHESP', 'Energia_Copel', 'Energia_COCEL', 'Energia_Coprel', 'Energia_CPFL', 'Energia_Demei', 'Energia_DME', 'Energia_Ecom', 'Energia_EDP', 'Energia_Elektro', 'Energia_EMGD', 'Energia_Enel', 'Energia_Energisa', 'Energia_Energisa_testes', 'Energia_Engie', 'Energia_EquatorialEnergia', 'Energia_EquatorialEnergia_testes', 'Energia_Lemon', 'Energia_Light', 'Energia_Merito', 'Energia_NeoEnergia', 'Energia_Nova', 'Energia_Roraima', 'Energia_Safira', 'Energia_SantaMaria', 'Energia_SULGIPE', 'Energia_Tereos', 'Energia_Test']

completer = QCompleter(lista_parser)
tela.lineEdit_2.setCompleter(completer)

load_active = load_arquivos()
tela.pushButton.clicked.connect(load_active.ler_arquivo_fornecedor)
tela.pushButton_2.clicked.connect(load_active.extrair_arquivo_fornecedor)
tela.pushButton_4.clicked.connect(load_active.executar_SAM2)

sys.exit(app.exec_())