import pandas as pd, time, os, glob, sys
from datetime import datetime, timedelta
import pandas as pd, time, os, glob
import telegram
from datetime import datetime, timedelta
import win32com.client
import pyautogui
from PIL import ImageOps, Image

now = datetime.now() - timedelta()
horaAtual = now.strftime('%H')
dayWeek = now.date().strftime("%A")
dataLog = time.strftime('%Y-%m-%d_%H-%M-%S')

logDir = ''
logFile = 'log_'+dataLog+'.txt'
log = logDir + logFile
fileDoMetodoLog = logDir + 'detailed'
padraoLog = '\nlog||Report||automação Telegram||'+dataLog+'||'

# token bot correios gross
token_correios = ''


# Função do log
def write_log(file,texto):
    detailed_log = datetime.now().strftime("-%Y-%m-%d-")+ 'log.txt'
    file = file +  detailed_log
    print(texto)
    if os.path.exists(file):
        logs = open(file, 'a+')
        logs.write(texto)
        logs.close()
    else: 
        logs = open(file, 'w+')
        logs.write(texto)
        logs.close()
    return texto

# Preparar msg
def send(msg, chat_id, token=token_):

    #Bot para envio da mensagem    
    bot = telegram.Bot(token=token)
    bot.sendMessage(chat_id=chat_id, text="Olá! Segue ultima atualização do Report.")
    bot.sendDocument(chat_id=chat_id, document=open(file_envio, 'rb'))
    
    
if (dayWeek =="Sunday"):
    print('Domingo, não executaremos o resto do código!!!')
    mensagem = 'Código não executado por ser domingo||day_exception'
    write_log(fileDoMetodoLog, padraoLog + mensagem)
    sys.exit()
else:
    print('Não é domingo, executaremos o resto do código!!!')
    mensagem = 'Código proseguindo para ser executado por não ser domingo||execution_register'
    write_log(fileDoMetodoLog, padraoLog + mensagem)
    
    if horaAtual =='22' or horaAtual =='00' or horaAtual =='01' or horaAtual =='02' or horaAtual =='03' or horaAtual =='04' or horaAtual =='05' or horaAtual =='06':
        print('Já são mais de 21h30, não executaremos o resto do código')
        mensagem = 'Código não executado por ser mais que 21h30||hour_exception'
        write_log(fileDoMetodoLog, padraoLog + mensagem)
        sys.exit()
    else:
        print('Ainda não é 22h, executaremos o resto do código')
        mensagem = 'Código executado por ser ainda não mais que 21h30||execution_register'
        write_log(fileDoMetodoLog, padraoLog + mensagem)
        
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
            mensagem = 'Arquivo aberto com sucesso||execution_register'
            write_log(fileDoMetodoLog, padraoLog + mensagem)
            
            time.sleep(180)
            # Printando o EXCEL
            pyautogui.screenshot('')
            print ("Print realizado")
            mensagem = 'Print retirado com sucesso||execution_register'
            write_log(fileDoMetodoLog, padraoLog + mensagem)

            # Abrindo e cortando a imagem
            img = Image.open('')
            border = (30, 220, 500, 88)
            cropped_img = ImageOps.crop(img, border).save('')
            print ("Imagem formatada")
            mensagem = 'imagem formatada||execution_register'
            write_log(fileDoMetodoLog, padraoLog + mensagem)
            o.Quit()

            # Enviar msg
            message = ""
            send(message, -chat_id)
            print ("Mensagem enviada")
            mensagem = 'Mensagem enviada||execution_register'
            write_log(fileDoMetodoLog, padraoLog + mensagem)

            # Removendo arquivos gerados
            os.remove('')
            os.remove('')

            print ("arquivos deletados")

        except:
            the_type, the_value, the_traceback = sys.exc_info()
            print('falha ao realizar o script')
            print(the_type, ',' ,the_value,',', the_traceback)
            mensagem = 'falha ao realizar o script||failure_register'
            write_log(fileDoMetodoLog, padraoLog + mensagem)