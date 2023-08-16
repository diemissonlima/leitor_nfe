import xmltodict
import os
import pandas as pd
from time import sleep
import json

print("Lendo os XML's, aguarde por gentileza..")
sleep(2)


def pegar_infos(nome_arquivo, valores):
    # print(f'Pegou informações {nome_arquivo}')
    # print('-=' * 20)
    with open(f'nfe/{nome_arquivo}', 'rb') as arquivo_xml:
        dic_arquivo = xmltodict.parse(arquivo_xml)
        # print(json.dumps(dic_arquivo, indent=4))
        if "NFe" in dic_arquivo:
            infos_nf = dic_arquivo["NFe"]["infNFe"]
        else:
            infos_nf = dic_arquivo["nfeProc"]["NFe"]["infNFe"]
        natureza_op = infos_nf["ide"]["natOp"]
        data_nf = infos_nf["ide"]["dhEmi"]
        numero_nota = infos_nf["ide"]["nNF"]
        modelo_nf = infos_nf["ide"]["mod"]
        valor_nf = infos_nf["total"]["ICMSTot"]["vNF"]
        if "CNPJ" in infos_nf["emit"]:
            cnpj_emitente = infos_nf["emit"]["CNPJ"]
        else:
            cnpj_emitente = infos_nf["emit"]["CPF"]
        emitente = infos_nf["emit"]["xNome"]
        logradouro = infos_nf["emit"]["enderEmit"]["xLgr"]
        bairro = infos_nf["emit"]["enderEmit"]["xBairro"]
        numero = infos_nf["emit"]["enderEmit"]["nro"]
        if "dest" not in infos_nf:
            destinatario = 'Não Informado'
            cnpj_destinatario = 'Não Informado'
            dest_logradouro = 'Não Informado'
            dest_bairro = 'Não Informado'
            dest_numero = 'Não Informado'
        else:
            if "xNome" not in infos_nf["dest"]:
                destinatario = 'Não Informado'
            else:
                destinatario = infos_nf["dest"]["xNome"]
            if "CNPJ" in infos_nf["dest"]:
                cnpj_destinatario = infos_nf["dest"]["CNPJ"]
            elif "CPF" in infos_nf["dest"]:
                cnpj_destinatario = infos_nf["dest"]["CPF"]
            if "enderDest" not in infos_nf["dest"]:
                dest_logradouro = 'Não Informado'
                dest_bairro = 'Não Informado'
                dest_numero = 'Não Informado'
            else:
                dest_logradouro = infos_nf["dest"]["enderDest"]["xLgr"]
                dest_bairro = infos_nf["dest"]["enderDest"]["xBairro"]
                dest_numero = infos_nf["dest"]["enderDest"]["nro"]
        valores.append([numero_nota, modelo_nf, valor_nf, natureza_op, data_nf[:10], cnpj_emitente, emitente,
                        logradouro, bairro, numero, cnpj_destinatario, destinatario, dest_logradouro, dest_bairro,
                       dest_numero])


lista_arquivos = os.listdir("nfe")

colunas = [
    "numero_nota", "modelo", "valor_nota", "natureza_op", "data_nota", "cnpj_emitente", "razao_emitente",
    "logradouro_emitente", "bairro_emitente", "numero_emitente", "cnpj_destinatario", "razao_destinatario",
    "logradouro_destinatario", "bairro_destinatario", "numero_destinatario"
]

valores = []

cont = 0
for arquivo in lista_arquivos:
    pegar_infos(arquivo, valores)
    cont += 1

tabela = pd.DataFrame(columns=colunas, data=valores)
tabela.to_excel("Notas Fiscais.xlsx", index=False)

print('-=' * 20)
print('Leitura finalizada com sucesso!'.center(40))
print(f'Foram lidas {cont} NFs.'.center(40))
print('-=' * 20)
input('Pressione ENTER para encerrar.')
