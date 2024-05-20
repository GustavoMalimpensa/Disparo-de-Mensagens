"""AUTOMAÇÃO DE ENVIO DE MENSAGENS PARA WHATSAPP"""

import webbrowser
import openpyxl
from urllib.parse import quote
from time import sleep
import pyautogui 

workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['basedados']


chrome_path = 'C:\Program Files\Google\Chrome\Application\chrome.exe'  
webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))


webbrowser.get('chrome').open('https://web.whatsapp.com/')
sleep(10)

for linha in pagina_clientes.iter_rows(min_row=2) :
    
    nome = linha[0].value
    telefone = linha[1].value

    
    mensagem = (
        "Olá {nome}, boa tarde!\n\n"
        "Meu nome é Gustavo, sou especialista tecnológico da Solution Tech! 🧑🏻‍💻\n\n"
        "A Empresa que mais impulsiona em Rio Claro trás uma incrível PROMOÇÃO! "
        "E-commerce personalizado com entrada de R$ 379,00 e as melhores condições do mercado, "
        "já com todos os benefícios inclusos! 🤩\n\n"
        "Gostaria de saber mais?\n\n"
        "🌐 Confira como ficaria: finaestampafashion.com/"
    ).format(nome=nome)

    
    sleep(8)
   
    try:
        
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
       
        webbrowser.get('chrome').open(link_mensagem_whatsapp)
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(4)
        pyautogui.click(seta[0], seta[1])
        sleep(4)
       
        pyautogui.hotkey('ctrl', 'w')
        sleep(4)
    except: 
        print(f'Não foi possivel mandar mensagem para {nome}')
        with open('erros.csv', 'a', newline='', encoding='utf=8') as arquivo:
            arquivo.write(f'{nome},{telefone}')

    