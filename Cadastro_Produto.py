from csv import excel
import time
from openpyxl import load_workbook
import pyautogui

wb = load_workbook(filename='./Planilhas/rosstamp.xlsx')

def sleepI():
    time.sleep(0.6)

# LACO DE REPETICAO LISTA ACESSA PLANILHA
i = 2
sheet_ranges = wb['Planilha1']
for description in sheet_ranges:
    description = [
    sheet_ranges['B{}'.format(i)].value,     # 0    CODIGO PRODUTO
    sheet_ranges['C{}'.format(i)].value,     # 1    DESCRIÇÃO PRODUTO
    sheet_ranges['AI{}'.format(i)].value,    # 2    CONDIÇÃO IF LINHA 
    sheet_ranges['AJ{}'.format(i)].value,    # 3    FATOR DE COMPRA SE SELECIONADO CAIXA
    sheet_ranges['F{}'.format(i)].value,     # 4    VERIFICA UNIDADES CAIXAS
    sheet_ranges['AK{}'.format(i)].value,    # 5    INSERE VALOR PESO 
    sheet_ranges['D{}'.format(i)].value,     # 6    INSERE NCM
    sheet_ranges['AC{}'.format(i)].value,    # 7    INSERE EAN 13
    ]

    #CONVERTE NCM EM FORMATO DE NCM 0000.00.00
    lenNcm = str(description[6]) #RECEBE VALOR DO NCM FORMATADO EM STRING
    if len(lenNcm) < 8: #ENUMERA CADA INDEX NA STRING
        lenNcm = lenNcm.zfill(8)
    formatedNcm = '{}.{}.{}'.format(lenNcm[:4], lenNcm[4:6], lenNcm[6:8])#FORMATA A STRING COM BASE NA ENUMERACAO 0000.00.00

    # INICIANDO CARACTERISTICAS

    #CLILCA INCLUIR
    pyautogui.moveTo(x=334, y=683)
    pyautogui.click()
    pyautogui.click()

    time.sleep(2.5)

    # ADICIONANDO DESCRIÇÕES
    pyautogui.moveTo(609, 91,  0.3)
    sleepI()
    pyautogui.click()
    pyautogui.typewrite(description[1]) #DESCRIÇÃO
    sleepI()
    pyautogui.press('TAB',2)
    pyautogui.typewrite(description[1]) #DESCRIÇÃO RESUMIDA
    sleepI()
    #cloca caracteristicas  
    pyautogui.moveTo(343, 166, 0.3)
    pyautogui.click()
    sleepI()
    # PREENChE FORNECEDOR
    pyautogui.moveTo(1048, 207, 0.3)
    pyautogui.click()
    pyautogui.typewrite("ROSSTAMP")
    pyautogui.press('ENTER')
    pyautogui.moveTo(662, 279, 0.3)
    sleepI()
    pyautogui.click()
    pyautogui.moveTo(662, 502, 0.3)
    sleepI()
    pyautogui.click()

    #VERIFICA LINHA
    if description[2] =='MIX2PET MEDICAMENTOS':
        pyautogui.moveTo(738, 246, 0.3)
        pyautogui.click()
        sleepI()
        pyautogui.moveTo(736, 277, 0.3)
        pyautogui.mouseDown()
        pyautogui.moveTo(736, 388, 0.3)
        pyautogui.click()
        sleepI()
        pyautogui.moveTo(623, 349, 0.3)
        pyautogui.click()
    elif description[2] == 'MIX2PET':
        pyautogui.moveTo(738, 246, 0.3)
        pyautogui.click()
        sleepI()
        pyautogui.moveTo(736, 277, 0.3)
        pyautogui.mouseDown()
        pyautogui.moveTo(736, 388, 0.3)
        pyautogui.click()
        sleepI()
        pyautogui.moveTo(598, 319, 0.3)
        pyautogui.click()
    elif description[2] == 'MERCHANDISING':
        pyautogui.moveTo(738, 246, 0.3)
        pyautogui.click()
        sleepI()
        pyautogui.moveTo(736, 277, 0.3)
        pyautogui.mouseDown()
        pyautogui.moveTo(736, 384, 0.3)
        pyautogui.click()
        sleepI()
        pyautogui.moveTo(602, 321, 0.3)
        pyautogui.click()
    elif description[2] == 'GOURMET':
        pyautogui.moveTo(738, 246, 0.3)
        pyautogui.click()
        sleepI()
        pyautogui.moveTo(736, 277, 0.3)
        pyautogui.mouseDown()
        pyautogui.moveTo(736, 448, 0.3)
        pyautogui.click()
        sleepI()
        pyautogui.moveTo(611, 320, 0.3)
        pyautogui.click()
#    elif description[2] == 'TUTELO':
#        pyautogui.moveTo(738, 246, 0.3)
#        pyautogui.click()
#        sleepI()
#        pyautogui.moveTo(736, 277, 0.3)
#        pyautogui.mouseDown()
#        pyautogui.moveTo(736, 462, 0.3)
#        pyautogui.click()
#        sleepI()
#        pyautogui.moveTo(612, 388, 0.3)
#        pyautogui.click()
#    elif description[2] == 'HIGIENE4':
#        pyautogui.moveTo(738, 246, 0.3)
#        pyautogui.click()
#        sleepI()
#        pyautogui.moveTo(736, 277, 0.3)
#        pyautogui.mouseDown()
#        pyautogui.moveTo(736, 376, 0.3)
#        pyautogui.click()
#        sleepI()
#        pyautogui.moveTo(606, 364, 0.3)
#        pyautogui.click()
#    #FIM LINHA 

    # PREENCHE CODIGO DO FABRICANTE
    pyautogui.press('TAB')
    sleepI()
    pyautogui.typewrite(str(description[0])) #CODIGO FABRICANTE
    pyautogui.press('TAB')

    # UNIDADE DE ESTOQUE
    # SELECIONA A 1º UNIDADE
    pyautogui.press('DOWN', 18, 0.2)
    # UNIDADE DE COMPRA
    pyautogui.press('TAB')
    # CONDIÇÃO SE FOR CAIXA
    if description[4] == 'CX':
        pyautogui.press('DOWN', 3)
        pyautogui.press('TAB')
        pyautogui.typewrite(str(description[3]))# FATOR

    # QUANTIDADE MIN VENDA
    pyautogui.press('TAB')
    pyautogui.typewrite("1")
    # QUANTIDADE MULT VENDA
    pyautogui.press('TAB')
    pyautogui.typewrite("1")

    # FLAG COMPRA
    pyautogui.moveTo(349, 315, 0.3)
    pyautogui.click()    
    sleepI()
    pyautogui.click()
    pyautogui.press('TAB', 4)
    pyautogui.press('SPACE')

    # ADICIONAR PESO LIQUIDO
    pyautogui.moveTo(642, 329, 0.3)
    pyautogui.click()
    pyautogui.typewrite(str(description[5])) #PESO

    # ADICIONAR PESO BRUTO
    pyautogui.press('TAB')
    pyautogui.typewrite(str(description[5])) #PESO
    # ADICIONAR DIAS ESTIMADO
    pyautogui.press('TAB',2)
    pyautogui.typewrite("60")

    #====================================================
    #           FIM CARACTERISTICAS
    #====================================================
    # SELECIONAR DADOS FISCAIS
    sleepI()
    pyautogui.moveTo(437, 165, 0.3)
    sleepI()
    pyautogui.click()
    sleepI()

    # SELECIONA TIPO DE PRODUTO
    pyautogui.moveTo(458, 201, 0.3)
    sleepI()
    pyautogui.click()
    sleepI()
    pyautogui.moveTo(358, 225, 0.3)
    sleepI()
    pyautogui.click()

    # ADICIONA O NCM
    pyautogui.press('TAB',2)
    sleepI()
    pyautogui.typewrite(formatedNcm) #NCM

    # CLASSIFICAÇÃO FISCAL
    pyautogui.moveTo(955,393, 0.3)
    pyautogui.click()
    sleepI()
    pyautogui.moveTo(869, 425, 0.3)
    pyautogui.click()
    #====================================================
    #           FIM FISCAIS
    #====================================================
    # INICIO COMPLEMENTARES
    sleepI()
    pyautogui.moveTo(530, 161, 0.3)
    pyautogui.click()
    sleepI()

    # PREENCHE DESCRIÇÃO COMPLEMENTAR
    pyautogui.moveTo(812, 205, 0.3)
    pyautogui.click()
    sleepI()
    pyautogui.typewrite(description[1]) # DESCRIÇÃO
    sleepI()

   #INSERE CODIGO DE BARRAS
    pyautogui.press('TAB')
    pyautogui.press('DOWN', 4)
    pyautogui.press('TAB')
    pyautogui.typewrite(str(description[7]))
    pyautogui.press('TAB')
    pyautogui.press('DOWN', 4)
    pyautogui.press('TAB')
    pyautogui.typewrite(str(description[7]))

    # PREENCHE DESCRIÇÃO COMPLETAR NF
    pyautogui.press('TAB')
    sleepI()
    pyautogui.typewrite(description[1]) # DESCRIÇÃO
    sleepI()
    # PREENCHE DESCRIÇÃO PRODUTO FABRICANTE
    pyautogui.press('TAB')
    sleepI()
    pyautogui.typewrite(description[1]) # DESCRIÇÃO
    sleepI()

    # DESFLAG "ATIVAR CONTR VERBA"
    pyautogui.moveTo(503, 387, 0.3)
    sleepI()
    pyautogui.click()
    # FLAG "BLOQUEIO ECOMMERCE"
    # DESGFLAG "ATIVAR CONTR VERBA"
    pyautogui.moveTo(503, 410, 0.3)
    sleepI()
    pyautogui.click()
    sleepI()

    # EFETIVA
    pyautogui.moveTo(1038, 687, 0.3)
    pyautogui.click()
    sleepI()
    time.sleep(2.5)
    # CLICA OK WMS
    pyautogui.moveTo(683, 407, 0.3)
    sleepI()
    pyautogui.click()
    sleepI()
    time.sleep(3)
    # CLICA FECHAR WMS
    pyautogui.moveTo(1111, 44, 0.3)
    sleepI()
    pyautogui.click()
    sleepI()
    time.sleep(3.5)

    #ASSOCIAÇÃO DE COMISSAO PADRAO
    #SELECIONA TABELA 2022 NOVA
    pyautogui.moveTo(735, 151, 0.3)
    pyautogui.click()
    sleepI()
    pyautogui.moveTo(587, 226, 0.3)
    pyautogui.click()
    # BUSCAR
    sleepI()
    pyautogui.moveTo(977, 283, 0.3)
    pyautogui.click()
    # SELECIONA ITENS BUSCADOS 
    sleepI()
    pyautogui.moveTo(449, 408, 0.3)
    pyautogui.mouseDown()
    pyautogui.moveTo(482, 507, 0.3)
    sleepI()
    pyautogui.click()
    # ALTERA GRUPO DE COMISSAO PARA 0%
    sleepI()
    pyautogui.moveTo(955, 351, 0.3)
    pyautogui.click()
    sleepI()
    pyautogui.moveTo(809, 430, 0.3)
    pyautogui.click()
    sleepI()
    # CONFIRMA
    pyautogui.moveTo(1007, 354, 0.3)
    pyautogui.click()
    # EFETIVA
    pyautogui.moveTo(988, 624, 0.3)
    sleepI()
    pyautogui.click()
    # CLICA OK DEPOIS DE EFETIVAR
    sleepI()
    pyautogui.moveTo(681, 407, 0.3)
    sleepI()
    pyautogui.click()
    # FECHA COMISSAO
    sleepI()
    pyautogui.moveTo(1021, 95, 0.3)
    sleepI()
    pyautogui.click()
    time.sleep(4)

#    #ADICIONA EAN EM UNIDADES
#    #SELECIONA CARACTERISTICAS
#    pyautogui.moveTo(349, 162, 0.3)
#    pyautogui.click()
#    sleepI()
#    #CLIA EM UNIDADES
#    pyautogui.moveTo(996, 375, 0.3)
#    pyautogui.click()
#    time.sleep(1.5)
#    #CLICA NO BRANCO PARA ADICIONAR NOVA LINHA DE EAN
#    pyautogui.moveTo(690, 577, 0.3)
#    pyautogui.click()
#    sleepI()
#    pyautogui.press('TAB')
#    pyautogui.press('SPACE')
#    pyautogui.moveTo(459, 470, 0.3)
#    pyautogui.click()
#    sleepI
#    pyautogui.click()
#    sleepI()
#    pyautogui.press('DOWN', 16)
#    pyautogui.press('TAB')
#    pyautogui.press('DOWN')
#    pyautogui.press('TAB')
#    pyautogui.typewrite(description[7])
#    #efetiva
#    pyautogui.moveTo(1132, 609, 0.3)
#    pyautogui.click()
#    #fecha tela de unidades
#    pyautogui.moveTo(1155, 119, 0.3)
#    pyautogui.click()

    print(description)

    i += 1