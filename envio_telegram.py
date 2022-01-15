import pandas as pd, time, os, glob, sys
from datetime import datetime, timedelta
import pandas as pd, time, os, glob
import telegram
from datetime import datetime, timedelta
import win32com.client
import pyautogui
from PIL import ImageOps, Image

# token bot Telegram
token_correios = ''

# Preparar msg
def send(msg, chat_id, token=token_correios):

    #Bot para envio da mensagem    
    bot = telegram.Bot(token=token)
    bot.sendMessage(chat_id=chat_id, text="Olá! Segue ultima atualização do Report.")
    bot.sendDocument(chat_id=chat_id, document=open(file_envio, 'rb'))
    
# Informaões fixas
dirRaiz = ''
filename= dirRaiz +''
file_envio = ''

try:
    # Gerar arquivo
    o = win32com.client.Dispatch("Excel.Application")
    o.Visible = True
    # Ler Excel
    sheets = o.Workbooks.Open(filename)
    work_sheets = sheets.Worksheets[0]
    print ("Arquivo aberto")
    time.sleep(160)

    # Printando o EXCEL
    pyautogui.screenshot('')
    print ("Print realizado")

    # Abrindo e cortando a imagem
    img = Image.open('')
    border = (30, 292, 600, 110)
    cropped_img = ImageOps.crop(img, border).save('')
    print ("Imagem formatada")

    # Enviar msg
    message = ""
    send(message, <-group_id_telegram>)
    print ("Mensagem enviada")

    # Removendo arquivos gerados
    os.remove('')
    os.remove('')
    o.Quit()
    print ("arquivos deletados")
    
except:
    the_type, the_value, the_traceback = sys.exc_info()
    print('falha ao realizar o script')
    print(the_type, ',' ,the_value,',', the_traceback)
    
