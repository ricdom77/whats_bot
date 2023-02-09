from docx2txt import process
from PySimpleGUI import popup_ok, popup_yes_no, popup_ok_cancel
from pyautogui import position, doubleClick, tripleClick, hotkey, press
from clipboard import copy
from time import sleep
import sys


def capturaCoordenadas(msg):
    alerta = f"Posicione o mouse sobre o campo de\n{msg} do What's App\n e aguarde 3 segundos..."
    resp = popup_ok_cancel(alerta)
    if resp != 'OK':
        popup_ok('Rotina cancelada')
        sys.exit()
    sleep(3)
    return position()


def executarBot(local, coord, dados):
    doubleClick(coord) if local == 'pessoa' else tripleClick(coord)
    copy(dados)
    hotkey('ctrl', 'v')
    press('enter')


# puxa dados do arquivo word 'doc_base.docx'
dados = process('doc_base.docx')
lista = dados.split('\n\n')

# confirma coordenadas da tela
coord_pessoa = capturaCoordenadas('pesquisa de pessoas')
coord_msg = capturaCoordenadas('mensagens')

tudo_ok = popup_yes_no(
    f'Coordenadas capturadas.\n\n - Pesquisa: {coord_pessoa};\n - Mensagens: {coord_msg};\n\nPodemos dar início?')

if tudo_ok == 'Yes':

    contador = 0
    qtde_msgs = int(len(lista)/2)

    for i in range(qtde_msgs):

        # definição dos endereçamentos e respectivas msgs
        para = lista[contador+1]
        msg = lista[contador]

        # seleciona a pessoa
        executarBot(local='pessoa', coord=coord_pessoa, dados=para)

        # encaminha a msg
        executarBot(local='msg', coord=coord_msg, dados=msg)

        contador += 2
        sleep(0.3)

    msg_final = '1 mensagem foi encaminhada'
    if qtde_msgs > 1:
        msg_final = f'{qtde_msgs} mensagens foram encaminhadas'

    popup_ok(msg_final)

else:
    popup_ok('Rotina cancelada')
