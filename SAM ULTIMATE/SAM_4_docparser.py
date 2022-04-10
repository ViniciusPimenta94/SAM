import pydocparser
from time import sleep
from tqdm import tqdm

class Parseamento_docparser():
    def __init__(self, fornecedor, list_arquivo_pdf):
        '''
            @arquivo_pdf: arquivo gerado pela saída do GED no RavenDB
            @tabela_twm: contém @fornecedor e @identificador_twm
            @fornecedor: Razão social do fornecedor que vem do SQL do twm
            @identificador_twm: identificador da fatura para acompanhar seu trajeto
            @parser_login: Usa-se a chave API da guiando para enviar arquivos ao Parser
            @result: ping.pong (pong: significa que tudo ocorreu corretamente)
        '''
        parser = pydocparser.Parser()
        parser.login('7a5a9bde8daf2b40e2282e7002a5bc2b689770eb')
        result = parser.ping()
        print(result)   
        self.parser = parser
        self.fornecedor = fornecedor
        self.list_arquivo_pdf = list_arquivo_pdf
        
        # parsers = self.parser.list_parsers()
        # lista = [i['label'] for i in parsers]
        # lista_agua = []
        # lista_energia = []
        # for i in lista:
        #     if 'Agua' in i:
        #         lista_agua.append(i)
        # for i in lista:
        #     if 'Energia' in i:
        #         lista_energia.append(i)
        
        #-----------------------------------------------Envia PDF para ser parseado -------------------------------------------------------------
    def importar_pdf(self):
        list_data = []
        list_id = []
        fornecedor = self.fornecedor #'Energia_Cemig'
        list_path = self.list_arquivo_pdf # '123131413123.pdf'
        #path = 'teste_cemig_3010594770_20220222.pdf'
        
        for path in list_path:
            id = self.parser.upload_file_by_path(path, fornecedor) #args: file to upload, the name of the parser
            list_id.append(id)

        print('\nAguarda processamento no Parser...\n')
        
        for _ in tqdm(range(300)):
            sleep(1)

        for id in list_id:
            #Note that "fileone.pdf" was in the current working directory
            data = self.parser.get_one_result(fornecedor, id) # The id is the doc id that was returned by `parser.upload()`
            list_data.append(data)
                     
        return list_data
