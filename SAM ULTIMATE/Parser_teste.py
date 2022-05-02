from pathlib import Path
import sys, time
from PyQt5 import uic, QtWidgets
from PyQt5.QtWidgets import QCompleter
from SAM_4_docparser import Parseamento_docparser

class load_arquivos:
        def __init__(self):
            self.arquivo_fornecedor = []
            self.palavra_chave = []
            self.list_fornecedor = []
            self.nome_fornecedor = []
            
        def ler_arquivo_fornecedor(self):
            self.fornecedor = QtWidgets.QFileDialog.getOpenFileNames()[0]
            print('Arquivos importados:', len(self.fornecedor))
            
            for arquivo in self.fornecedor:
                self.list_fornecedor.append(arquivo)
            print('Load file fornecedor complete!')
            
        def executar_SAM4(self):
            self.inicio = time.time()
            nome_fornecedor = tela.lineEdit_2.text()           
            self.nome_fornecedor = nome_fornecedor
            arquivo_pdf = self.list_fornecedor
            
            print(f'{nome_fornecedor}\n')
            
            chama_SAM4 = Parseamento_docparser(nome_fornecedor, arquivo_pdf)
            self.list_json_parseado = chama_SAM4.importar_pdf()
            
            load_active.teste_parser()
            
        def teste_parser(self):           
            # JSON PARSEADO
            # lista_data = [{'id': 'c5248ddb7bffa0eb2743ce717d7ded61', 'document_id': '6c4be26902e6186ff5f6afd3a0d9d9eb', 'remote_id': '', 'file_name': '15728964_20220422teste.pdf', 'media_link': 'https://api.docparser.com/v1/document/media/DfaXevMuZ5v7KTwL3wwRtTLebGjIijB8C6K0hgX0Ib-hu_HC8XapGZkq1Yq5IMoAT1G2FKVbFsu1Pb5yOoRcu5bVvmdkb-87Rp46S48GSqs', 'media_link_original': 'https://api.docparser.com/v1/document/media/DfaXevMuZ5v7KTwL3wwRtTLebGjIijB8C6K0hgX0Ib-hu_HC8XapGZkq1Yq5IMoAT1G2FKVbFsu1Pb5yOoRcu5bVvmdkb-87Rp46S48GSqs/original', 'media_link_data': 'https://api.docparser.com/v1/document/media/DfaXevMuZ5v7KTwL3wwRtTLebGjIijB8C6K0hgX0Ib-hu_HC8XapGZkq1Yq5IMoAT1G2FKVbFsu1Pb5yOoRcu5bVvmdkb-87Rp46S48GSqs/data', 'page_count': 1, 'uploaded_at': '2022-04-19T14:58:36+00:00', 'processed_at': '2022-04-19T14:59:29+00:00', 'uploaded_at_utc': '2022-04-19T14:58:36+00:00', 'uploaded_at_user': '2022-04-19T11:58:36+00:00', 'processed_at_utc': '2022-04-19T14:59:29+00:00', 'processed_at_user': '2022-04-19T11:59:29+00:00', 'dc_identificador_conta': '15728964', 'dt_vencimento': '22/04/2022', 'vl_total': '534,42', 'dc_razao_social_cliente': 'SOC MINEIRA DE CULTURA', 'dc_razao_social': 'DME Distribuição S/A - DMED', 'dc_identificador_pessoa_juridica': '2366430300010', 'dc_identificador_pessoa_juridica_cliente': '17178195004588', 'dc_endereco_cliente': 'CAM UM, 2010 / 1 RURAL 37701-970 Poços de Caldas - MG', 'dt_leitura_anterior': '04/02/2022', 'dt_leitura_atual': '07/03/2022', 'unidade_medida': 'kWh', 'dt_mes_referencia': '03/2022', 'vl_base_calculo_icms': '494,49', 'vl_base_calculo_pis_pasep': '370,87', 'vl_base_calculo_cofins': '370,87', 'vl_aliquota_icms': '25,00a', 'vl_aliquota_pis_pasep': '1,05%', 'vl_aliquota_cofins': '4,86', 'vl_valor_cofins': '18,02', 'vl_valor_icms': '123,62', 'vl_valor_pis_pasep': '3,89', 'dc_classe': 'COMERCIAL-Outros', 'nome_fornecedor': 'DME', 'dc_subclasse': 'serviços e outras atividades', 'marcador_padrao': 'Consumo', 'dc_grupo_tensao': '3', 'dc_subgrupo_tensao': 'B3', 'dc_modalidade_tarifaria': 'Convencional', 'vl_tensao_nominal': '220/127', 'vl_limites_tensao': '201 a 231', 'vl_demanda_contratada': None, 'nr_cliente': '105182', 'dc_limites_tensao': None, 'nr_instalacao': '15728964', 'dt_emissao': '21/03/2022', 'ic_classificador_medidor': 'Trifásico', 'valores_consumo': {'valor': '475 kWh'}, 'nr_medidor': '300212656', 'nr_nota_fiscal': '014823111', 'dc_codigo_boleto_1': '8369000000573442004100011001684517380172894032', 'dc_codigo_boleto_2': None, 'dc_codigo_boleto': '836900000057344200410001100168451738017289403226', 'nu_fatura_fornecedor': '014823111', 'dt_prev_prox_leitura': None, 'dc_endereco_fornecedor': 'RUA AMAZONAS, 65 - CENTRO - Poços de Caldas - MG - CEP 37701-008', 'nu_serie_fiscal': 'B', 'vl_leitura_anterior': '66', 'vl_leitura_atual': '3 66768', 'valores_faturados': [{'descricao': 'CONSUMO', 'unidade_medida': '', 'quantidade': '475', 'tarifa_preco': '0,83983', 'valor': '398,92'}, {'descricao': 'BANDEIRA TARIFARIA', 'unidade_medida': '', 'quantidade': '475', 'tarifa_preco': '0,20120', 'valor': '95,57'}, {'descricao': 'Subtotal', 'unidade_medida': '', 'quantidade': '', 'tarifa_preco': '', 'valor': '494,49'}, {'descricao': 'C I P - MUNICIPAL', 'unidade_medida': '', 'quantidade': '', 'tarifa_preco': '', 'valor': '39,93'}, {'descricao': 'Subtotal', 'unidade_medida': '', 'quantidade': '', 'tarifa_preco': '', 'valor': '39,93'}], 'valores_faturados_auditoria': [{'descricao': 'CONSUMO', 'unidade_medida': '', 'quantidade': '475', 'tarifa_preco': '0,83983', 'valor': ''}, {'descricao': 'BANDEIRA TARIFARIA', 'unidade_medida': '', 'quantidade': '475 W', 'tarifa_preco': '0,20120 57', 'valor': '95,57'}, {'descricao': 'C I P - MUNICIPAL', 'unidade_medida': '', 'quantidade': '', 'tarifa_preco': '', 'valor': '39,93'}], 'dc_tipo': 'Nota Fiscal/Conta de Energia Elétrica', 'valores_faturados_auditoria_1': None, 'valores_faturados_auditoria_2': None, 'dc_serie_nota_fiscal': 'B', 'nr_cfop': '5.253', 'valores_faturados_volumetria': None, 'dc_identificador_layout': 'Padrao'}, {'id': 'c5248ddb7bffa0eb2743ce717d7ded61', 'document_id': '6c4be26902e6186ff5f6afd3a0d9d9eb', 'remote_id': '', 'file_name': '15728964_20220422teste.pdf', 'media_link': 'https://api.docparser.com/v1/document/media/DfaXevMuZ5v7KTwL3wwRtTLebGjIijB8C6K0hgX0Ib-hu_HC8XapGZkq1Yq5IMoAT1G2FKVbFsu1Pb5yOoRcu5bVvmdkb-87Rp46S48GSqs', 'media_link_original': 'https://api.docparser.com/v1/document/media/DfaXevMuZ5v7KTwL3wwRtTLebGjIijB8C6K0hgX0Ib-hu_HC8XapGZkq1Yq5IMoAT1G2FKVbFsu1Pb5yOoRcu5bVvmdkb-87Rp46S48GSqs/original', 'media_link_data': 'https://api.docparser.com/v1/document/media/DfaXevMuZ5v7KTwL3wwRtTLebGjIijB8C6K0hgX0Ib-hu_HC8XapGZkq1Yq5IMoAT1G2FKVbFsu1Pb5yOoRcu5bVvmdkb-87Rp46S48GSqs/data', 'page_count': 1, 'uploaded_at': '2022-04-19T14:58:36+00:00', 'processed_at': '2022-04-19T14:59:29+00:00', 'uploaded_at_utc': '2022-04-19T14:58:36+00:00', 'uploaded_at_user': '2022-04-19T11:58:36+00:00', 'processed_at_utc': '2022-04-19T14:59:29+00:00', 'processed_at_user': '2022-04-19T11:59:29+00:00', 'dc_identificador_conta': '15728964', 'dt_vencimento': '22/04/2022', 'vl_total': '534,42', 'dc_razao_social_cliente': 'SOC MINEIRA DE CULTURA', 'dc_razao_social': 'DME Distribuição S/A - DMED', 'dc_identificador_pessoa_juridica': '23664303000104', 'dc_identificador_pessoa_juridica_cliente': '1', 'dc_endereco_cliente': 'CAM UM, 2010 / 1 RURAL 37701-970 Poços de Caldas - MG', 'dt_leitura_anterior': '04/02/2022', 'dt_leitura_atual': '07/03/2022', 'unidade_medida': 'kWh', 'dt_mes_referencia': '03/2022', 'vl_base_calculo_icms': '494,49', 'vl_base_calculo_pis_pasep': '370,87', 'vl_base_calculo_cofins': '370,87', 'vl_aliquota_icms': '25,00', 'vl_aliquota_pis_pasep': '1,05', 'vl_aliquota_cofins': '4,86', 'vl_valor_cofins': '18,02', 'vl_valor_icms': '123,62', 'vl_valor_pis_pasep': '3,89', 'dc_classe': 'COMERCIAL-Outros', 'nome_fornecedor': 'DME', 'dc_subclasse': 'serviços e outras atividades', 'marcador_padrao': 'Consumo', 'dc_grupo_tensao': '3', 'dc_subgrupo_tensao': 'B3', 'dc_modalidade_tarifaria': 'Convencional', 'vl_tensao_nominal': '220/127', 'vl_limites_tensao': '201 a 231', 'vl_demanda_contratada': None, 'nr_cliente': '105182', 'dc_limites_tensao': None, 'nr_instalacao': '15728964', 'dt_emissao': '21/03/2022a', 'ic_classificador_medidor': 'Trifásico', 'valores_consumo': {'valor': ''}, 'nr_medidor': '300212656', 'nr_nota_fiscal': '014823111', 'dc_codigo_boleto_1': '836900000057344200410001100168451738017289403226', 'dc_codigo_boleto_2': None, 'dc_codigo_boleto': '836900000057344200410001100168451738017289403226', 'nu_fatura_fornecedor': '014823111', 'dt_prev_prox_leitura': None, 'dc_endereco_fornecedor': 'RUA AMAZONAS, 65 - CENTRO - Poços de Caldas - MG - CEP 37701-008', 'nu_serie_fiscal': 'B', 'vl_leitura_anterior': '66', 'vl_leitura_atual': '3 66768', 'valores_faturados': [{'descricao': '', 'unidade_medida': '', 'quantidade': '', 'tarifa_preco': '0,83983', 'valor': '398,92'}, {'descricao': 'BANDEIRA TARIFARIA', 'unidade_medida': '', 'quantidade': '475', 'tarifa_preco': '0,20120', 'valor': '95,57'}, {'descricao': 'Subtotal', 'unidade_medida': '', 'quantidade': '', 'tarifa_preco': '', 'valor': '494,49'}, {'descricao': 'C I P - MUNICIPAL', 'unidade_medida': '', 'quantidade': '', 'tarifa_preco': '', 'valor': '39,93'}, {'descricao': 'Subtotal', 'unidade_medida': '', 'quantidade': '', 'tarifa_preco': '', 'valor': '39,93'}], 'valores_faturados_auditoria': [{'descricao': 'CONSUMO', 'unidade_medida': '', 'quantidade': '475 kWh', 'tarifa_preco': '0,83983 45', 'valor': '398,92'}, {'descricao': 'BANDEIRA TARIFARIA', 'unidade_medida': '', 'quantidade': '475', 'tarifa_preco': '0,20120', 'valor': '95,57'}, {'descricao': 'C I P - MUNICIPAL', 'unidade_medida': '', 'quantidade': '', 'tarifa_preco': '', 'valor': '39,93'}], 'dc_tipo': 'Nota Fiscal/Conta de Energia Elétrica', 'valores_faturados_auditoria_1': None, 'valores_faturados_auditoria_2': None, 'dc_serie_nota_fiscal': 'B', 'nr_cfop': '5.253', 'valores_faturados_volumetria': None, 'dc_identificador_layout': 'Padrao'}]
            
            # SEPARANDO REGRAS PARSER DOS VALORES PARSEADOS
            for indice, parseado in enumerate(self.list_json_parseado):
                for i, v in enumerate(parseado):
                    myfile = Path(f".\{self.nome_fornecedor}.txt")
                    myfile.touch(exist_ok=True)
                    f = open(f"{self.nome_fornecedor}.txt", 'a+')
    
                    print('-'*75)
                    print(f"{v['file_name']} - {v['dc_identificador_layout']}\n")
                    f.write('-'*75)
                    f.write(f"\n{v['nome_fornecedor']}\n")
                    f.write(f"\n{v['file_name']} - {v['dc_identificador_layout']}\n\n")
                    chave_data = []
                    valor_data = []
                
                    for key in parseado[i].keys():
                        chave_data.append(key)
                    
                    for values in parseado[i].values():
                        valor_data.append(values)
                    
                    for i, v in enumerate(chave_data):
                        #f.writelines(f'{chave_data[i]} - {valor_data[i]}\n')
                        pass
                    #f.write('\n')
                # MOSTRANDO CAMPOS COM VALORES VAZIOS    
                    for i, v in enumerate(valor_data):
                        if v == '' or v == None:
                            if chave_data[i] == 'remote_id' or chave_data[i][:8] == 'marcador':
                                pass
                            else:
                                print(f'{chave_data[i]} - CAMPO VAZIO')
                                f.writelines(f'{chave_data[i]} - CAMPO VAZIO\n')
                            #pass
                    print()    
                # VALIDANDO CNPJ
                    for i, v in enumerate(chave_data):
                        if 'dc_identificador_pessoa_juridica' in v:
                            cnpj = valor_data[i]
                            
                            lista = []
                            lista2 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
                            soma1 = 0
                            soma2 = 0
                            
                            cnpjn = ''.join(filter(str.isnumeric, cnpj))
                            
                            for i, v in enumerate(cnpjn):
                                if i < 12:
                                    lista.append(cnpjn[i])
                            
                            for i, v in enumerate(lista):
                                soma1 += int(lista[i])*(lista2[i])
                            
                            valor = 11 - (soma1 % 11)
                            
                            if valor > 9:
                                valor = '0'
                            else:
                                valor = str(valor)
                            
                            lista += valor
                            
                            x = [6]
                            x.extend(lista2)
                            
                            for i, v in enumerate(lista):
                                soma2 += int(lista[i])*(x[i])
                            
                            valor2 = 11 - (soma2 % 11)
                            
                            if valor2 > 9:
                                valor2 = '0'
                            else:
                                valor2 = str(valor2)
                            
                            lista += valor2
                            
                            sequencia = ''.join(cnpjn) == (str(lista[0])) * len(cnpjn)    # Elimina as sequências
                            
                            validar = ''.join(lista)
                            
                            if validar == cnpjn and not sequencia:
                                pass
                            else:
                                if cnpj == '' or cnpj == None:
                                    pass
                                else:
                                    for i, v in enumerate(valor_data):
                                        if cnpj == v:
                                            print(f'{chave_data[i]} - {cnpj} - CNPJ INVÁLIDO')
                                            f.writelines(f'\n{chave_data[i]} - {cnpj} - CNPJ INVÁLIDO')
                    print()
                    f.write('\n')
                    
                # VERIFICANDO OS CAMPOS DE VALORES (Regras Vl_)
                    for i, v in enumerate(chave_data):
                        if v[:3] == 'vl_':
                            try:
                                if type(valor_data[i]) == str:
                                    string = '!.#%'
                                    char_rep = {k: '' for k in string}
                                    num = valor_data[i].translate(str.maketrans(char_rep))
                                    num = float(num.replace(',', '.'))
                            except:
                                if valor_data[i] == '' or valor_data[i] == None or valor_data[i] == 'Não Informado':
                                    pass
                                else:
                                    print(f'{chave_data[i]} - {valor_data[i]}')
                                    f.writelines(f'\n{chave_data[i]} - {valor_data[i]}')
                
                    print()
                    f.write('\n')
                
                # VERIFICANDO OS CAMPOS DE DATA (Regras Dt_)
                    for i, v in enumerate(chave_data):
                        if v[:3] == 'dt_':
                            try:
                                float(valor_data[i].replace('/',''))
                            except:
                                if valor_data[i] == '' or valor_data[i] == None:
                                    pass
                                else:
                                    print(f'{chave_data[i]} - {valor_data[i]}')
                                    f.writelines(f'\n{chave_data[i]} - {valor_data[i]}')
                
                    print()
                    f.write('\n')
                    
                # VERIFICANDO O CÓDIGO DO BOLETO
                    for i, v in enumerate(chave_data):
                        if 'dc_codigo' in v:
                            boleto = valor_data[i]
                            
                            if boleto == '' or boleto == None:
                                pass            
                            elif len(boleto) == 47 or len(boleto) == 48:
                                pass
                            else:
                                print(f'{chave_data[i]} - CAMPO INCOMPLETO - {valor_data[i]}')
                                f.writelines(f'{chave_data[i]} - CAMPO INCOMPLETO - {valor_data[i]}')
                
                    print()
                    f.write('\n')
                
                # MOSTRANDO SE O CAMPO DE CONSUMO DO TWM ESTÁ VAZIO
                    for i, v in enumerate(valor_data):                       
                        if type(v) == dict:
                            chave_dict = []
                            valor_dict = []
                            campo_dict = chave_data[i]
                            
                            for key in v.keys():
                                chave_dict.append(key)
                    
                            for values in v.values():
                                valor_dict.append(values)
                    
                            for i, v in enumerate(valor_dict):
                                try:
                                    if type(valor_dict[i]) == str:
                                        string = '!.#%'
                                        char_rep = {k: '' for k in string}
                                        num = valor_dict[i].translate(str.maketrans(char_rep))
                                        num = float(num.replace(',', '.'))
                                except:
                                    if v == '':
                                        print(f'{campo_dict}_{chave_dict[i]} - CAMPO VAZIO')
                                        f.writelines(f'\n{campo_dict}_{chave_dict[i]} - CAMPO VAZIO')
                                        #pass                    
                                    elif v == None:
                                        print(f'{campo_dict}_{chave_dict[i]} - CAMPO VAZIO')
                                        f.writelines(f'\n{campo_dict}_{chave_dict[i]} - CAMPO VAZIO')
                                        #pass
                                    else:
                                        print(f'{campo_dict}_{chave_dict[i]} - {valor_dict[i]} - CAMPO INCORRETO')
                                        f.writelines(f'\n{campo_dict}_{chave_dict[i]} - {valor_dict[i]} - CAMPO INCORRETO')
                            f.write('\n')
                    
                # VERIFICANDO SE A PARTE DE VALORES NA AUDITORIA ESTÁ CORRETA                    
                        elif type(v) == list:
                            if chave_data[i] == 'valores_consumo':
                                for ind, val in enumerate(v):
                                    chave_dict = []
                                    valor_dict = []
                                    campo_dict = chave_data[i]
                                    
                                    for key in v[ind].keys():
                                        chave_dict.append(key)
                            
                                    for values in v[ind].values():
                                        valor_dict.append(values)
                            
                                    for i, v in enumerate(valor_dict):
                                        try:
                                            if type(valor_dict[i]) == str:
                                                string = '!.#%'
                                                char_rep = {k: '' for k in string}
                                                num = valor_dict[i].translate(str.maketrans(char_rep))
                                                num = float(num.replace(',', '.'))
                                        except:
                                            if v == '':
                                                print(f'{campo_dict}_{chave_dict[i]} - CAMPO VAZIO')
                                                f.writelines(f'\n{campo_dict}_{chave_dict[i]} - CAMPO VAZIO')
                                                #pass                    
                                            elif v == None:
                                                print(f'{campo_dict}_{chave_dict[i]} - CAMPO VAZIO')
                                                f.writelines(f'\n{campo_dict}_{chave_dict[i]} - CAMPO VAZIO')
                                                #pass
                                            else:
                                                print(f'{campo_dict}_{chave_dict[i]} - {valor_dict[i]} - CAMPO INCORRETO')
                                                f.writelines(f'\n{campo_dict}_{chave_dict[i]} - {valor_dict[i]} - CAMPO INCORRETO')
                                    f.write('\n')

                            if chave_data[i] == 'valores_faturados_auditoria':
                                soma = 0    
                                for ind, val in enumerate(v):
                                    chave_aud = []
                                    valor_aud = []
                                    count = 0
                                    
                                    for key in v[ind].keys():
                                        chave_aud.append(key)  
                                        if key == 'valor':
                                            valor = count
                                        elif key == 'quantidade':
                                            qtd = count
                                        elif key == 'tarifa_preco':
                                            tarifa = count
                                        count += 1
                                        
                                    for values in v[ind].values():
                                        valor_aud.append(values)
                                    
                                    try:
                                        if type(valor_aud[valor]) == str:
                                            string = '!.#%'
                                            char_rep = {k: '' for k in string}
                                            num = valor_aud[valor].translate(str.maketrans(char_rep))
                                            num = float(num.replace(',', '.'))
                                            soma += num
                                    except:
                                        if valor_aud[valor] == None or valor_aud[valor] == '':
                                            print(f'{chave_data[i]}_valor - CAMPO VAZIO')
                                            f.writelines(f'\n{chave_data[i]}_valor - CAMPO VAZIO')
                                        else:
                                            print(f'{chave_data[i]}_valor - {valor_aud[valor]} - CAMPO INCORRETO')
                                            f.writelines(f'\n{chave_data[i]}_valor - {valor_aud[valor]} - CAMPO INCORRETO')
                                    
                                    try:
                                        if type(valor_aud[qtd]) == str:
                                            string = '!.#%'
                                            char_rep = {k: '' for k in string}
                                            num = valor_aud[qtd].translate(str.maketrans(char_rep))
                                            num = float(num.replace(',', '.'))
                                    except:
                                        if valor_aud[qtd] == None or valor_aud[qtd] == '':
                                            pass
                                        else:
                                            print(f'{chave_data[i]}_quantidade - {valor_aud[qtd]} - CAMPO INCORRETO')
                                            f.writelines(f'\n{chave_data[i]}_quantidade - {valor_aud[qtd]} - CAMPO INCORRETO')
                                    
                                    try:
                                        if type(valor_aud[tarifa]) == str:
                                            string = '!.#%'
                                            char_rep = {k: '' for k in string}
                                            num = valor_aud[tarifa].translate(str.maketrans(char_rep))
                                            num = float(num.replace(',', '.'))
                                    except:
                                        if valor_aud[tarifa] == None or valor_aud[tarifa] == '':
                                            pass
                                        else:
                                            print(f'{chave_data[i]}_tarifa_preco - {valor_aud[tarifa]} - CAMPO INCORRETO')
                                            f.writelines(f'\n{chave_data[i]}_tarifa_preco - {valor_aud[tarifa]} - CAMPO INCORRETO')
                                f.write(f'\nValor auditoria: {soma:.2f}\n')
                    print()
                    f.write('\n')
                    
                # CAMPOS NÃO VERIFICADOS
                    for i, v in enumerate(chave_data):
                        if v[:3] == 'vl_' or v[:3] == 'dt_' or v == '' or v == None:
                            pass  
                        elif 'dc_identificador_pessoa_juridica' in v:
                            pass
                        elif 'dc_codigo' in v:
                            pass
                        else:
                            if valor_data[i] == '' or valor_data[i] == None:
                                pass
                            elif type(valor_data[i]) == dict or type(valor_data[i]) == list:
                                pass
                            else:
                                #print(f'{chave_data[i]} - {valor_data[i]}')
                                pass
                            
                ########################################################################################################################
                    
                    f.write('\n          CAMPOS DE VALORES\n')
                    for i, v in enumerate(chave_data):
                        if v[:3] == 'vl_':
                            f.writelines(f'\n          {chave_data[i]} - {valor_data[i]}')
                    
                    f.write('\n\n          CAMPOS DE DATA\n')
                    for i, v in enumerate(chave_data):
                        if v[:3] == 'dt_':
                            f.writelines(f'\n          {chave_data[i]} - {valor_data[i]}')
                    
                    f.write('\n\n          CONSUMO\n')
                    for i, v in enumerate(chave_data):
                        if 'valores_consumo' in v:
                            f.writelines(f'\n          {chave_data[i]} - {valor_data[i]}')
                            
                    f.write('\n\n          VALORES AUDITORIA\n')
                    for i, v in enumerate(chave_data):
                        if v[-9:] == 'auditoria':
                            f.writelines(f'\n          {chave_data[i]} - {valor_data[i]}')
                        
                    f.write('\n\n          OUTROS CAMPOS IMPORTANTES\n')
                    for i, v in enumerate(chave_data):
                        if v[-5:] == 'conta':
                            f.writelines(f'\n          {chave_data[i]} - {valor_data[i]}')
                    f.write('\n')
                    for i, v in enumerate(chave_data):
                        if 'dc_endereco' in v:
                            f.writelines(f'\n          {chave_data[i]} - {valor_data[i]}')
                    f.write('\n')
                    for i, v in enumerate(chave_data):
                        if 'dc_razao_social' in v:
                            f.writelines(f'\n          {chave_data[i]} - {valor_data[i]}')
                    f.write('\n')
                    for i, v in enumerate(chave_data):
                        if 'dc_identificador_pessoa_juridica' in v:
                            f.writelines(f'\n          {chave_data[i]} - {valor_data[i]}')
                    f.write('\n')
                    for i, v in enumerate(chave_data):
                        if 'dc_codigo' in v:
                            f.writelines(f'\n          {chave_data[i]} - {valor_data[i]}')            
                    f.write('\n')
            fim = time.time()
            total = (fim - self.inicio)
            tempo = time.strftime('%H:%M:%S', time.gmtime(total))
            f.write(f'\nO tempo de execução foi de {tempo}\n')
            f.close()
            print('\nFIM DO PROCESSO')

app = QtWidgets.QApplication([])
tela = uic.loadUi("./assets/Interface_Parser.ui")

tela.show()

lista_parser = ['Agua_Aegea', 'Agua_Aegea_Seta_Vermelha', 'Agua_Aegea_Seta_Vermelha_testes', 'Agua_Aguas_Arenapolis', 'Agua_AguasDoBrasil', 'Agua_BRK', 'Agua_CAERR', 'Agua_CAESA', 'Agua_Caesb', 'Agua_Cagece', 'Agua_Casan', 'Agua_CEDAE', 'Agua_Cesama', 'Agua_Cesama_teste', 'Agua_Cesan', 'Agua_CIS', 'Agua_Codau', 'Agua_COGERH', 'Agua_Cohasb', 'Agua_Compesa', 'Agua_Comusa', 'Agua_Copasa', 'Agua_Corsan', 'Agua_Cosanpa', 'Agua_DAAE_Araraquara', 'Agua_DAAE_Rio_Claro', 'Agua_DAE_Americana', 'Agua_DAE_Bauru', 'Agua_DAE_Jundiai', 'Agua_DAE_Santa_Barbara', 'Agua_DAE_Santana_do_Livramento', 'Agua_DAEM', 'Agua_DAEP', 'Agua_DAES_Juina', 'Agua_DAEV', 'Agua_Damae_SaoJoaoDelRei', 'Agua_DEMAE_Campo_Belo', 'Agua_Departamento_Divisao_Servico', 'Agua_Depasa', 'Agua_DMAE_Poços_de_Caldas', 'Agua_DMAE_Poços_de_Caldas_copy1', 'Agua_DMAE_Porto_Alegre', 'Agua_DMAE_Uberlandia', 'Agua_Elipse', 'Agua_Emasa', 'Agua_Embasa', 'Agua_Embasa_testes', 'Agua_EMDAEP', 'Agua_Iguasa', 'Agua_Linhas', 'Agua_Prefeitura_Itirapina', 'Agua_Prefeituras', 'Agua_Prolagos', 'Agua_SAAE_Amparo', 'Agua_SAAE_Atibaia', 'Agua_SAAE_Bacabal', 'Agua_SAAE_Balsas', 'Agua_SAAE_Barra_Mansa', 'Agua_SAAE_Bebedouro', 'Agua_SAAE_Campo_Maior', 'Agua_SAAE_Canaa_Dos_Acarajas', 'Agua_SAAE_Capivari', 'Agua_SAAE_Catu', 'Agua_SAAE_Caxias', 'Agua_SAAE_Ceara_Mirim', 'Agua_SAAE_Codo', 'Agua_SAAE_Estancia', 'Agua_SAAE_Garca', 'Agua_SAAE_Governador_Valadares', 'Agua_SAAE_Grajau', 'Agua_SAAE_Ibitinga', 'Agua_SAAE_Iguatu', 'Agua_SAAE_Indaiatuba', 'Agua_SAAE_Itapetinga', 'Agua_SAAE_Itapira', 'Agua_SAAE_Jacareí', 'Agua_SAAE_Juazeiro', 'Agua_SAAE_LençoisPaulistas', 'Agua_SAAE_Limoeiro', 'Agua_SAAE_Linhares', 'Agua_SAAE_Mogi_Mirim', 'Agua_SAAE_Morada_Nova', 'Agua_SAAE_Parauapebas', 'Agua_SAAE_Parintins', 'Agua_SAAE_PDFSimples', 'Agua_SAAE_Penedo', 'Agua_SAAE_Quadrado', 'Agua_SAAE_Quadradoteste', 'Agua_SAAE_Quixeramobim', 'Agua_SAAE_Sao_Carlos', 'Agua_SAAE_Sobral', 'Agua_SAAE_Volta_Redonda', 'Agua_SAAEJ_Jaboticabal', 'Agua_SAAEP', 'Agua_SAAETRI', 'Agua_Sabesp', 'Agua_SAE_Araguari', 'Agua_SAE_Ituiutaba', 'Agua_Saec', 'Agua_Saecil', 'Agua_SaemaAraras', 'Agua_SAEP', 'Agua_SAESA', 'Agua_SAEV', 'Agua_SAMAE_Blumenau', 'Agua_SAMAE_Brusque', 'Agua_SAMAE_Caxias', 'Agua_SAMAE_Jaragua_do_Sul', 'Agua_Samae_Mogi_Guaçu', 'Agua_Sanasa', 'Agua_Saneago', 'Agua_Saneamento', 'Agua_Sanebavi', 'Agua_SANEP', 'Agua_Sanepar', 'Agua_Saneparteste', 'Agua_Sanesalto', 'Agua_SANESUL', 'Agua_SEMAE_Ararangua', 'Agua_Semae_Mogi_das_Cruzes', 'Agua_SEMAE_Piracicaba', 'Agua_SEMAE_São_José_do_Rio_Preto', 'Agua_SEMAE_Sao_Leopoldo', 'Agua_Semasa_Itajaí', 'Agua_Semasa_Santo_Andre', 'Agua_SESAN', 'Agua_Setae', 'Energia_2W', 'Energia_Aliança', 'Energia_AmazonasEnergia', 'Energia_Banco_BTG', 'Energia_CEA', 'Energia_CEB', 'Energia_CEEE', 'Energia_CELESC', 'Energia_CELESC_copy', 'Energia_Cemig', 'Energia_CemigSim', 'Energia_Cerbanorte', 'Energia_Cercar', 'Energia_CERCI', 'Energia_Cergal', 'Energia_CERMOFUL', 'Energia_CERRP', 'Energia_CHESP', 'Energia_Copel', 'Energia_COCEL', 'Energia_Coprel', 'Energia_CPFL', 'Energia_Demei', 'Energia_DME', 'Energia_Ecom', 'Energia_EDP', 'Energia_Elektro', 'Energia_EMGD', 'Energia_Enel', 'Energia_Energisa', 'Energia_Energisa_testes', 'Energia_Engie', 'Energia_EquatorialEnergia', 'Energia_EquatorialEnergia_testes', 'Energia_Lemon', 'Energia_Light', 'Energia_Merito', 'Energia_NeoEnergia', 'Energia_Nova', 'Energia_Roraima', 'Energia_Safira', 'Energia_SantaMaria', 'Energia_SULGIPE', 'Energia_Tereos', 'Energia_Test']

completer = QCompleter(lista_parser)
tela.lineEdit_2.setCompleter(completer)

load_active = load_arquivos()
tela.pushButton.clicked.connect(load_active.ler_arquivo_fornecedor)
tela.pushButton_4.clicked.connect(load_active.executar_SAM4)

sys.exit(app.exec_())     

# Vl_ -> Valor, Impostos, Valores faturados, Valor da leitura atual e anterior
# Dt_ -> Vencimento, emissão, data da leitura anterior e atual, data da próxima leitura
# Consumo, Código de barras, CNPJ.

# Falta verificar: Conta, Razão Social, Endereço.

# 4 pilares (conta, vencimento, emissão e valor), Razão social e CNPJ (tanto do cliente quanto do fornecedor), Impostos, Endereço, Valores faturados (serviços), Consumo e dados de leitura (Data da leitura anterior, Data da próxima leitura, Data da leitura atual, valor da leitura atual e valor da leitura anterior) e código de barras
