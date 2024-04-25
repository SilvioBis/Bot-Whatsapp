import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

webbrowser.open('https://web.whatsapp.com/')
sleep(30)

#Ler Planilha e guardar informações sobre nome, telefone e data
workbook = openpyxl.load_workbook('contatos_chat_bot.xlsx')
pagina_clientes = workbook['Planilha1']

for linha in pagina_clientes.iter_rows(min_row=2):
    #variaveis nome,telefone, data
    nome = linha[0].value
    telefone = linha[1].value
    data = linha[2].value

    mensagem = f'Olá {nome} Sou um chatbot desenvolvido em python(linguagem de programação) Voce tem até o dia {data.strftime("%d/%m/%Y")}. Para me responder se recebeu a mensagem e se gostou e deixar o like na minha publicação que vou deixar o link https://link_do_post_.com Obs. é apenas um teste'
    
    #Criar links personalizados do whatsapp e enviar mensagens para cada cliente com base nos dados da planilha
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(10)
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(5)
        pyautogui.click(seta[0],seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl','w')
        sleep(5)
    except:
        print(f'Não foi possivel enviar mensagem para {nome}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}')