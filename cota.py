#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# para verificar se os pacotes necessários estão instalados
import sys
import subprocess
import pkg_resources

required = {'beautifulsoup4', 'lxml', 'openpyxl', 'xlcalculator', 'requests', 'urllib3'}
installed = {pkg.key for pkg in pkg_resources.working_set}
missing = required - installed

if missing:
        python = sys.executable
        subprocess.check_call([python, '-m', 'pip', 'install', *missing], stdout=subprocess.DEVNULL)



# para navegar no html
from bs4 import BeautifulSoup
# para fazer a solicitação
import requests

# para ignorar certificado
from urllib3.exceptions import InsecureRequestWarning
# Suppress only the single warning from urllib3 needed.
requests.packages.urllib3.disable_warnings(category=InsecureRequestWarning)


# para abrir o xlsx
import openpyxl

import locale
locale.setlocale( locale.LC_ALL, 'pt_BR.UTF-8' )


import json

# para atualizar a data de atualização de cotações
import datetime

# para testar a data
from datetime import date

# para avaliar funções dentro do excel -- necessário para os trechos do moneylog

from xlcalculator import ModelCompiler
from xlcalculator import Model
from xlcalculator import Evaluator


def call_URL(codAtivo):
    headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}
    eh_acao = 0
    eh_acao = requests.head('https://statusinvest.com.br/acoes/' + codAtivo , verify=False, headers=headers)
    if eh_acao.status_code == 200:
        req = requests.get('https://statusinvest.com.br/acoes/' + codAtivo , verify=False, headers=headers)
    else:
        req = requests.get('https://statusinvest.com.br/fundos-imobiliarios/' + codAtivo, verify=False, headers=headers)

    if req.status_code == 200:
        content = req.content
    else:
        print('Problemas na Requisição a %s! ' , req)
        print(req.content.decode())

    return content


def retorna_valor_rendimento(URL):
    soup = BeautifulSoup(URL, 'lxml')
    rendimentos=soup.find('input', attrs={'id':'results'})
    # pega o json dos rendimentos
    rend_json=rendimentos['value']
    # mostra o json: json.dumps(rend_json)
    data = json.loads(rend_json)
    # ultimo mes: data[0] -- {u'etd': u'Rendimento', u'd': 0, u'ed': u'31/08/2020', u'ov': None, u'm': 0, u'sv': u'1,09000000', u'sov': u'-', u'pd': u'08/09/2020', u'v': 1.09, u'y': 0, u'et': u'Rendimento', u'adj': False, u'ad': None}
    try:
        data_pagamento = data[0]['pd']
        ultimorend = data[0]['sv']
    except:
        ultimorend = str("0")
    print(ultimorend, end=" ")
    return ultimorend, data_pagamento


def retorna_valor_cota(URL):

    soup = BeautifulSoup(URL, 'lxml')
    divvalor=soup.find('div', attrs={'title':'Valor atual do ativo'})
    cota = str(divvalor.strong.string)
    print("{0:<6}".format(cota), end=" "), 
    return cota

# f= raw_input("Digite o nome do arquivo a ser lido/gravado: ")
f = sys.argv[1];
excel_document = openpyxl.load_workbook(f)

excel_document.sheetnames
sheet = excel_document['Cotacoes']
#print sheet['A2'].value

# atualiza a data da atualização de cotações
current_datetime = datetime.datetime.now()
sheet['C6'].value=current_datetime.strftime('%x')


# a definição de multiple_cells tem que ser feita várias vezes mesmo por conta do cursor e de onde ele está!

print("Valores gravados")
multiple_cells = sheet['B8':'C24']
for row in multiple_cells:
    for cell in row:
        print("{0:<10}".format(cell.value), end=" " ),
    print('', end='\n')


print("Atualizando Valores")
multiple_cells = sheet['B8':'B24']
for row in multiple_cells:
    for cell in row:
        print("{0:<10}".format(cell.value) , end=" " ),
        URL = call_URL(cell.value)
        valoratual = retorna_valor_cota( URL )
        cell.offset(row=0,column=1).value=locale.atof(valoratual)
#        ultimorend = retorna_valor_rendimento( cell.value )[0]
#        data_pagamento = retorna_valor_rendimento( cell.value )[1]
        ultimorend, data_pagamento = retorna_valor_rendimento( URL )
        cell.offset(row=0,column=2).value=locale.atof(ultimorend)
        cell.offset(row=0,column=3).value=data_pagamento
    print('', end='\n')


print("Valores novos")
multiple_cells = sheet['B8':'E24']
for row in multiple_cells:
    for cell in row:
        print("{0:<10}".format(cell.value), end=" " ),
    print('', end='\n')


excel_document.save(f)


print("Valores de Rendimentos para Moneylog")

filename = f
compiler = ModelCompiler()
new_model = compiler.read_and_parse_archive(filename)
evaluator = Evaluator(new_model)

# 2022-12-01      0000.00 rendimento,fii,$FII| Rendimentos de $MES do FII
valor_mes= 0
saidamlog=[]
mes_atual= date.today().month
multiple_cells = sheet['B8':'B24']
for row in multiple_cells:
    for cell in row:
        fii= cell.value
        valoratual = cell.offset(row=0,column=1).value
        ultimorend = cell.offset(row=0,column=2).value
        numcotas = cell.offset(row=0,column=4).value
        data_pagamento =  cell.offset(row=0,column=3).value
        celula_cotastenho=cell.offset(row=0,column=5).coordinate
        celula_totrend=cell.offset(row=0,column=6).coordinate
        #print(celula_totrend, celula_cotastenho)
        val_cotastenho = evaluator.evaluate('Cotacoes!' + celula_cotastenho)
        total_rendimentos = evaluator.evaluate('Cotacoes!' + celula_totrend)
        tr=float(str(round(total_rendimentos,2)))
#        print(f'{tr:.2f} fii {fii} para # {val_cotastenho} cotas em {data_pagamento}')
#        print("value 'evaluated' for Cotacoes!Hx:", val1)
        dia = data_pagamento[0:2]
        mes = data_pagamento[3:5]
        ano = data_pagamento[6:10]
        data_mlog = ano + '-' + mes  + '-' + dia
        if int(mes_atual) == int(mes) or ( int(mes_atual)  + 1) == int(mes) :
#            print(f'{data_mlog}\t{tr:.2f}\trendimentos,{fii}| Rendimentos do fii {fii} para {val_cotastenho} cotas')
            valor_mes += tr
            saidamlog.append(f'{data_mlog}\t{tr:.2f}\trendimentos,{fii}| Rendimentos do fii {fii} para {val_cotastenho} cotas  {ultimorend:.2f}/cota')

print(f'Valor dos Rendimentos no Mês: {valor_mes:.2f}')
# ordena a lista de saída e mostra a saída ordenada por data
saidamlog.sort()

# grava a saída no moneylog-entries.txt para só importar ele no moneylog ;-)
with open("moneylog-entries.txt", "w") as external_file:
    for x in saidamlog:
        # mostra na tela
        print(x)
        # mostra no arquivo
        print(x, file=external_file)
    external_file.close()
    
#        print(f'{data_mlog}\t {numcotas:2}\t rendimentos,{fii}| Rendimentos do fii {fii}')
#        print(f'{data_mlog}\t {ultimorend.replace(",","."):2}\t rendimentos,{fii}| Rendimentos do fii {fii}')
#    print('', end='\n')

# https://stackoverflow.com/a/61585431/1663112
# >>> val1 = evaluator.evaluate('Cotacoes!G8')
# >>> print("value 'evaluated' for Cotacoes!G8:", val1)
# value 'evaluated' for Cotacoes!G8: 12

# https://openpyxl.readthedocs.io/en/stable/api/openpyxl.cell.cell.html
# referência para documentação da biblioteca openpyxl
