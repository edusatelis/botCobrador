import openpyxl
from urllib.parse import quote
import webbrowser
import pyautogui
from time import sleep


workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Página1']


for linha in pagina_clientes.iter_rows(min_row=2):
  nome = linha[0].value
  telefone = linha[1].value
  vencimento = linha[2].value

  texto = f'Olá {nome} seu boleto vence no dia {vencimento.strftime("%d/%m/%Y")}. Favor pagar no link https://www.link_do_pagamento.com'

  link_whats = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(texto)}'
  webbrowser.open(link_whats)
  sleep(15)
  try:

    pyautogui.hotkey('enter')
    sleep(5)
    pyautogui.hotkey('ctrl','w')
    sleep(2)
  except:
    print(f'Não foi possivel enviar mensagem {nome}')
    with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
        arquivo.write(f'{nome},{telefone}')