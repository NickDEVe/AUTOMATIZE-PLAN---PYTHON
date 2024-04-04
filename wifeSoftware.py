import openpyxl
import os
from openpyxl import workbook
from openpyxl import load_workbook
from PySimpleGUI import PySimpleGUI as te
from datetime import timedelta
from datetime import datetime







# lista de empresas para o programa colocar no arquivo para salvar

empresas = ['APOLO.xlsx', 'ABUZZO.xlsx', 'ABUZZO 2.xlsx','AMELIA.xlsx','CATAMARA.xlsx','CATAMARA ARI.xlsx', 'CIPLART AVANT.xlsx','CIPLART GREN PARK.xlsx','COELMA.xlsx' ,
'CONCRETO.xlsx','COSTA RICA.xlsx','CURITIBA GPM.xlsx','DESING.xlsx','DESING 2.xlsx','DAMA.xlsx','ENGECARB.xlsx','ENCANADOR PRC4.xlsx','ELETRICISTA.xlsx','FCA BOTA FOGO.xlsx',
'FCA OG3.xlsx','FCA MARLUC.xlsx','FCA ITICON.xlsx','FCA ALOJAMENTO.xlsx','FCA CASA.xlsx','FORMATE.xlsx','GRAN LONDON.xlsx','GREVILHA.xlsx','GRP 3 ILUMINE.xlsx',
'GRP MORADA DO PARK.xlsx','GRP SEDE LILIOS.xlsx','GRP FERNANDO.xlsx','HITECH.xlsx','INCORP. QUATRO.xlsx','JUSTE BARROSO.xlsx','JUSTE BALANK.xlsx','JUSTE CRECHE.xlsx',
'JUSTE CIPRIANO.xlsx','JUSTE TRAVEZA.xlsx','JUSTE NEO.xlsx','JPS EL CIELO.xlsx','JPS LE PARK.xlsx','JPS GRESCON.xlsx','JPS FOR SISON.xlsx','JPS SEDE.xlsx',
'JPS SEVEN.xlsx','JPS AQUALINA.xlsx','LBX.xlsx','MAFIP.xlsx','MGM.xlsx','MONDEO PIX NOTA.xlsx','COMUNIDADE TRIFOLD.xlsx','OG3.xlsx','OURO COLA.xlsx','PATIO MOUREIRA.xlsx',
'PATIO MOUREIRA 2.xlsx','PATIO FILIPINAS.xlsx','P3 CARMEN MIRANDA.xlsx','p3 PAULISTA.xlsx','P3.xlsx','PATIO.xlsx','PATIO DIRETORIA.xlsx','PEQUENO.xlsx',
'RENOVAR OG3.xlsx','RENOVAR PEDRO GRANADO.xlsx','RENOVAR PALM.xlsx','RICARDO(EDIF. PORTAL).xlsx','CALEFI.xlsx','SISA 3.xlsx',
'SISA.xlsx','SERATI.xlsx','ARGUS.xlsx','TERRA CASA.xlsx','TAEC.xlsx','TERRAÇO CRISTAL.xlsx','TK.xlsx','TRANSAMERICA.xlsx',
'UEM.xlsx','UNIMED CARL B..xlsx','VMX BRAVO.xlsx', 'VILLAGIO ITALIA.xlsx' , 'ZE DA EGUA.xlsx']

#numeros correspondentes ao mes selecionado

meses  = {
    1: "JANEIRO",
    2: "FEVEREIRO",
    3: "MARÇO",
    4: "ABRIL",
    5: "MAIO",
    6: "JUNHO",
    7: "JULHO",
    8: "AGOSTO",
    9: "SETEMBRO",
    10: "OUTUBRO",
    11: "NOVEMBRO",
    12: "DEZEMBRO"
}

# layout interface

te.theme("DarkPurple1")
te.theme_text_color("orange")
layout = [
    [te.Text('SELECIONE O ARQUIVO MODELO')],

    [te.Input(enable_events= True, key='Arquivo', font=('Arial Bold', 12), background_color='white', expand_x=True), te.FileBrowse(key='arquivo')],
    [te.Text('COLOQUE O NUMERO DO MES CORESPONDENTE')],
    [te.Input('', key=('numMes'), background_color='white')],
    [te.Text('COLOQUE O NOME DA PASTA QUE DESEJA SALVAR OS ARQUIVOS')],
    [te.Input('', key='sele', background_color='white')],
    


    [te.Button('Salvar')]

]

# janela

janela = te.Window('ROGERS - ATALHO PARA COPIAR PLANILHAS', layout, size=(700,300))




while True:
    eventos, valores = janela.read()
    print( eventos, valores)
    if eventos == te.WIN_CLOSED:

        break

    if eventos == 'adicionarData':
        txt= valores['data']
        janela['input'].update(value=txt)
        

    if eventos == 'Salvar':
         
        local = valores['sele']
            
        os.mkdir(f"{local}")
            
        for empresa in empresas:

            #pegando a planilha modelo para formatar

            plan = openpyxl.load_workbook(valores['arquivo'])
            sheet= plan['Plan1']
            nomoE= empresa.replace(".xlsx", "")

            #formatando a data  e o titulo da celula A1 da planilha

            nomeMes = meses.get(int(valores['numMes']))
            sheet['A1'] = f'FECHAMENTO {nomoE} {nomeMes} 2024'


            #salvando os arquivo na pasta especifica
            
           
            plan.save(empresa)

            
        

        
        arquivos = os.listdir()
        for arquivo in arquivos :
            
            arquivos = os.listdir()
            local = valores['sele']

            if 'xlsx' in arquivo :
                os.rename (arquivo, f'{local}\{arquivo}')