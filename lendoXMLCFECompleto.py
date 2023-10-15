import xml.etree.ElementTree as ET
import pandas as pd
import csv
import os


produtos = {}

def inserir_produto(nome, codigo, data, undMedida, vrUnitario,qtdItem,fornecedor,cnpj,custoTotalItem,NFE,CHAVE):
    if nome not in produtos:
        produtos[nome] = []
    produtos[nome].append({ "nome":nome,
                            "codigo": codigo,
                            "data": data,
                            "undMedida":undMedida,
                            "valorUnitario":vrUnitario,
                            "quantidade":qtdItem,
                            "fornecedor":fornecedor,
                            "cnpj":cnpj,
                            "custo": custoTotalItem,
                            "NFE":NFE,
                            "CHAVE":CHAVE                        
                            })
    
    '''
    # Exemplo de inserção
    inserir_produto("Produto1", "01-01-2023", 12345, 100.50)
    inserir_produto("Produto1", "05-01-2023", 12346, 110.50)
    inserir_produto("Produto2", "02-01-2023", 12347, 150.75)

    print(produtos)
    '''


def leituraXML(caminho):

    # Carregue o arquivo XML
    tree = ET.parse(caminho)
    print("ok1")
    # Obtenha o elemento raiz
    root = tree.getroot()
    print("ok2")
     # Encontre o elemento infCFe
    infCFe = root.find(".//infCFe")
    print("ok3")
    # Extraia o valor do atributo Id e remova o prefixo 'CFe'
    chave_eletronica = "'"+ infCFe.attrib['Id'][3:]
    print("ok4")
    # Encontre o valor de dEmi
    data_emissao = root.find(".//dEmi").text
    print("ok5")
    # Formate o valor para o formato desejado
    data_emissao_formatada = "{}/{}/{}".format(data_emissao[6:8], data_emissao[4:6], data_emissao[0:4])
    #print(data_emissao_formatada)


    meio_de_pagamento=root.find(".//cMP").text[1]
    if meio_de_pagamento == "1":
        meio_de_pagamento="Dinheiro"
    elif meio_de_pagamento == "2":
        meio_de_pagamento="Cheque"
    elif meio_de_pagamento == "3":
        meio_de_pagamento="Cartão de Crédito"
    elif meio_de_pagamento == "4":
        meio_de_pagamento="Cartão de Débito"
    elif meio_de_pagamento == "5":
        meio_de_pagamento="Cartão Refeição/Alimentação"
    elif meio_de_pagamento == "6":
        meio_de_pagamento="Vale Refeição/Alimentação (em papel)"
    elif meio_de_pagamento == "7":
        meio_de_pagamento="Outros"
    else:
        #print("Entrada inválida!")
        meio_de_pagamento="Outros"

    #print(meio_de_pagamento)

    # Lista para armazenar todos os produtos
    produtos = []
    try:
            
        for det in root.findall(".//det"):
            prod = det.find("./prod")
            print("ok1")
            codigo_produto = prod.find("cProd").text
            #cEAN = prod.find("cEAN").text if prod.find("cEAN") is not None else None  # Considerando que cEAN pode não existir em todos os registros
            print("ok2")            
            nome_produto = prod.find("xProd").text
            print("ok3")            
            NCM = prod.find("NCM").text
            print("ok4")            
            CFOP = prod.find("CFOP").text
            print("ok5")            
            unidade_medida = prod.find("uCom").text
            print("ok6")            
            qtd_item = prod.find("qCom").text.replace(".",",")
            print("ok7")            
            valor_unitario = prod.find("vUnCom").text.replace(".",",")
            print("ok8")            
            vProd = prod.find("vProd").text.replace(".",",")
            #indRegra = prod.find("indRegra").text if prod.find("indRegra") is not None else None  # Considerando que indRegra pode não existir em todos os registros
            print("ok9")            
            valor_total_item = prod.find("vItem").text.replace(".",",") if prod.find("vItem") is not None else None  # Considerando que vItem pode não existir em todos os registros
            
            writer.writerow([codigo_produto,nome_produto,NCM,CFOP,unidade_medida,qtd_item,valor_unitario,valor_total_item,data_emissao_formatada,meio_de_pagamento,chave_eletronica])
    except Exception as e:  # Captura qualquer exceção
        print(f"Ocorreu um erro!")  # Exibe a mensagem da exceção
        print(f"Descrição do erro: {e}")
        print("pressione Enter para continuar...")
        y=input()       

def extensao_arquivo(caminho):
    return os.path.splitext(caminho)[1]


def ler_todos_arquivos_xml(diretorio):
    # Percorra todos os arquivos no diretório
    for raiz, subdiretorios, arquivos in os.walk(diretorio):
        print("ok1")
        for nome_arquivo in arquivos:
            print("ok2")            
            caminho_completo = os.path.join(raiz, nome_arquivo)
            print("ok3")            
            caminho_atualizado = caminho_completo.replace('\\', '\\\\')
            print("ok4")                        
            if ('.XML' in extensao_arquivo(caminho_completo).upper()):
                try:
           
                    leituraXML(caminho_completo)
                    print("ok6")            
                except Exception as e:  # Captura qualquer exceção
                    print(f"Erro XML {caminho_completo}")
                    print(f"Descrição do erro: {e}")
                    print("pressione Enter para continuar...")
                    y=input()



diretorio = "C:\\Users\\Gabriel\\Desktop\\chromedriver_win32\\testexml\\Tamires\\TAmires\\092023\\"       #CAMINHO ONDE CONTÉM OS ARQUIVOS XML CF-E
diretorio_relatorio_cfe='C:\\Users\\Gabriel\\Desktop\\Reuniao Python\\Materiais\\VENDAS_NAO_AGRUPADAS092023.csv'
diretorio_relatorio_agrupamento_vendas_cfe='C:\\Users\\Gabriel\\Desktop\\Reuniao Python\\Materiais\\AGRUPAMENTO_VENDAS092023.xlsx'

############ INICIO LEITURA XML ############


with open(diretorio_relatorio_cfe, mode='w', newline='') as file:
    # Cria um objeto writer para escrever no arquivo CSV
    writer = csv.writer(file, delimiter=';')
    writer.writerow(["Codigo","Nome","NCM.", "CFOP", "UNID. MEDIDA", "QTD","VALOR UNITARIO","VALOR TOTAL","DATA EMISSAO","MEIO DE PAGAMENTO","CHAVE ELETRONICA"])
    ler_todos_arquivos_xml(diretorio)

############ FIM LEITURA XML ############




############# INICIO AGRUPAMENTO DE VENDAS ########################

import pandas as pd

# Ler o arquivo CSV
df = pd.read_csv(diretorio_relatorio_cfe, sep=';', encoding='ISO-8859-1')

# Substitua as vírgulas por pontos e converta a coluna para float
df['VALOR TOTAL'] = df['VALOR TOTAL'].str.replace(',', '.').astype(float)

# Totalizar valores diariamente e classificar por meio de pagamento
resultado = df.groupby(['DATA EMISSAO', 'MEIO DE PAGAMENTO'])['VALOR TOTAL'].sum().reset_index()

# Gravar resultado em um arquivo Excel
resultado.to_excel(diretorio_relatorio_agrupamento_vendas_cfe, index=False, engine='openpyxl')


############# FIM AGRUPAMENTO DE VENDAS ########################
