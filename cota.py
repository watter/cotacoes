#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# para verificar se os pacotes necessários estão instalados
import sys
import subprocess
import pkg_resources

required = {'bs4', 'lxml', 'openpyxl'}
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

import sys

import json

def retorna_valor_rendimento(codAtivo):
    if (codAtivo == "BIDI11") or (codAtivo == "ITSA4"):
        req = requests.get('https://statusinvest.com.br/acoes/' + codAtivo , verify=False)
    else:
        req = requests.get('https://statusinvest.com.br/fundos-imobiliarios/' + codAtivo, verify=False)

    if req.status_code == 200:
        content = req.content
    else:
        print('Problemas na Requisição a %s! ' , req)
    soup = BeautifulSoup(content, 'lxml')
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
    print(ultimorend)
    return ultimorend, data_pagamento


def retorna_valor_cota(codAtivo):
    if (codAtivo == "BIDI11") or (codAtivo == "ITSA4"):
        req = requests.get('https://statusinvest.com.br/acoes/' + codAtivo, verify=False)
    else:
        req = requests.get('https://statusinvest.com.br/fundos-imobiliarios/' + codAtivo , verify=False)
    if req.status_code == 200:
        content = req.content
    else:
        print('Problemas na Requisição a %s! ' , req)
    soup = BeautifulSoup(content, 'lxml')
    divvalor=soup.find('div', attrs={'title':'Valor atual do ativo'})
    cota = str(divvalor.strong.string)
    print("{0:<6}".format(cota)), 
    return cota

# f= raw_input("Digite o nome do arquivo a ser lido/gravado: ")
f = sys.argv[1];
excel_document = openpyxl.load_workbook(f)

excel_document.get_sheet_names()
sheet = excel_document.get_sheet_by_name('Cotacoes')
#print sheet['A2'].value



print("Valores gravados")
multiple_cells = sheet['B8':'C21']
for row in multiple_cells:
    for cell in row:
        print("{0:<10}".format(cell.value) ),
    print


print("Atualizando Valores")
multiple_cells = sheet['B8':'B21']
for row in multiple_cells:
    for cell in row:
        print("{0:<10}".format(cell.value) ),
        valoratual = retorna_valor_cota( cell.value )
        cell.offset(row=0,column=1).value=locale.atof(valoratual)
#        ultimorend = retorna_valor_rendimento( cell.value )[0]
#        data_pagamento = retorna_valor_rendimento( cell.value )[1]
        ultimorend, data_pagamento = retorna_valor_rendimento( cell.value )
        cell.offset(row=0,column=2).value=locale.atof(ultimorend)
        cell.offset(row=0,column=3).value=data_pagamento


print("Valores novos")
multiple_cells = sheet['B8':'E21']
for row in multiple_cells:
    for cell in row:
        print("{0:<10}".format(cell.value) ),
    print


excel_document.save(f)
