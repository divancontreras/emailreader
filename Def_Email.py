import win32com.client
import openpyxl
from openpyxl.styles import Font
import datetime
import time
from openpyxl.styles import Alignment
from datetime import timedelta
from ClassEmailBox import *

excelfile = '\\\\WSMX02402FP\\Shared\\IMP-EXP\INTERNATIONAL TRADE\\RK & Activos\\EXPORT P2\\2017\HORARIOS EXP P2.xlsx'
excelbackup = '\\\\WSMX02402FP\\Shared\\IMP-EXP\INTERNATIONAL TRADE\\RK & Activos\\EXPORT P2\\2017\HORARIOS EXP P2 - BACKUP.xlsx'


def SaveDoc(book):
    status = True
    while status:
        try:
            book.save(excelfile)
            status = False
        except:
            print("ERROR, EXCEL FILE OPEN, TRYING AGAIN IN 10 SECS!")
            status = True
            time.sleep(10)
    print("DOCUMENTO GUARDADO")


def InitEnv():
    print("Buscando cajas nuevas..." + "\n")
    book = openpyxl.load_workbook(excelfile)
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace(
        "MAPI")
    ws = book.worksheets[5]
    inbox = outlook.GetDefaultFolder(
        6)  # "6" refers to the index of a folder - in this case,
    # the inbox. You can change that number to reference
    # # any other folder
    messages = inbox.Items
    cajas_folder = inbox.Folders.Item("Sistema de Cajas")
    print(book.worksheets[5])
    IterInbox(ws, messages, cajas_folder, book)


def getRow(ws):
    for row in range(900, ws.max_row):
        if ws.cell(row=row, column=1).value == None:
            return row


def CheckRepeat(ws, row, fecha, hora, limit=30):
    # This function will check if the box already exists.
    for xrow in range(row - 1, limit, -1):
        localdate = ws.cell(row=xrow, column=14).value
        localhour = ws.cell(row=xrow, column=15).value
        if fecha == localdate and localhour == hora[:8]:
            return True
    return False


def FormatCells(ws, row):
    # This functions is only used to keep the same format in the document.
    for x in range(1, 21):
        cell = ws.cell(row=row, column=x)
        cell.font = Font(name="Calibri", size=9)
        cell.alignment = Alignment(horizontal='center', vertical='bottom')
        if (x == 3 or x == 4 or x == 14):
            cell.font = Font(name="Calibri", size=9, bold=True)


def IterInbox(ws, messages, cajas_folder, book):
    email = "AdminCajasMty"
    subject = 'P2- Reporte de Caja'
    subjectAd = 'NACIONAL'
    row = getRow(ws)
    rowAUx = row
    for message in messages:
        print(message.SenderEmailAddress)
        if message.Class == 43 :
            if message.SenderEmailType == "EX" :
                if subject in message.Subject and email in message.Sender.GetExchangeUser().PrimarySmtpAddress and not (
                        subjectAd in message.Subject):
                    print(message.Sender.GetExchangeUser().PrimarySmtpAddress)
                    if getData(message, ws, row):
                        row += 1
                    #message.Move(cajas_folder)
            else:
                if subject in message.Subject and email in message.SenderEmailAddress and not (
                        subjectAd in message.Subject):
                    print(message.SenderEmailAddress)
                    if getData(message, ws, row):
                        row += 1
                    #message.Move(cajas_folder)
    if rowAUx != row:
        SaveDoc(book)
    else:
        print("NO BOXES WERE FOUND!")


def InserData(Caja, Sello, TipoCaja, Destino, Salida, ws, fecha, hora, row):
    ##    Fill the excel cells with the info from the email.
    ##    Each cell will be filled invidivually with the especific value
    if CheckRepeat(ws, row, fecha, hora):
        return False
    else:
        fecha = fecha.split('-')
        formdate = fecha[1] + "/" + fecha[2] + "/" + fecha[0]
        today = datetime.date.today()
        ws.cell(
            row=row, column=1).value = datetime.datetime.now().isocalendar()[1]
        ws.cell(row=row, column=2).value = (today.strftime('%b')).upper()
        ws.cell(row=row, column=3).value = today.strftime("%#m/%#d/%Y")

        ##    This wil manage the Remesa number by the week of the year.
        if ((ws.cell(row=row - 1, column=1).value) !=
            (datetime.datetime.now().isocalendar()[1])):
            ws.cell(row=row, column=4).value = 1
        else:
            for xrow in range(row - 1, 0, -1):
                if ws.cell(row=xrow, column=4).value != None:
                    ws.cell(
                        row=row, column=4).value = ws.cell(
                            row=xrow, column=4).value + 1
                    break

        ##    Fill the excell cells with the info from the email.
        ##    Each cell will be filled invidivually with the especific value

        ws.cell(row=row, column=7).value = Sello
        ws.cell(row=row, column=8).value = Caja

        ## Specify when they are Dedicated boxes

        if 'HG' in TipoCaja:
            ws.cell(row=row, column=9).value = "DEDICADO"
        else:
            ws.cell(row=row, column=9).value = TipoCaja

        ws.cell(row=row, column=10).value = Destino.upper()
        t = time.strptime(Salida, "%H:%M")
        ws.cell(row=row, column=11).value = time.strftime("%I:%M %p", t)
        ws.cell(row=row, column=14).value = formdate
        ws.cell(row=row, column=15).value = hora[:8]
        FormatCells(ws, row)

        return True


def getData(message, ws, row):
    #Read the email and parse the key data
    #This variables are keeping the email data
    body = str(message.body).split()
    Caja = body[1]
    Sello = body[4]
    TipoCaja = message.Subject.split()[4]
    fecha = str(message.SentOn).split()[0]
    hora = str(message.SentOn).split()[1]
    # Some exceptions
    if ('Chino' in message.body):
        Destino = body[10] + " " + body[11]
        Salida = body[13]
    else:
        Destino = body[10]
        Salida = body[12]
    ##InserData() return a bool that states if the Data was successfully inserted
    if (InserData(Caja, Sello, TipoCaja, Destino, Salida, ws, fecha, hora,row)):
        print("_______BOX SAVED_______" + "\n")
        print(Caja + " " + Sello + " " + Destino + " " + Salida + " " +
              TipoCaja)
        print("___________________________" + "\n")
        return True
    ##If tha data was repeated the following will be displayed
    else:
        print("_______BOX REPEATED________" + "\n")
        print(Caja + " " + Sello + " " + Destino + " " + Salida + " " + TipoCaja)
        print("____________________________" + "\n")
        return False


def autoUpdateBox():
    cont = 0
    while True:
        book = openpyxl.load_workbook(excelfile)
        cont += 1
        InitEnv()
        if (cont == 30):
            status = True
            while status:
                try:
                    book.save(excelbackup)
                    status = False
                except:
                    print(
                        "ERROR, BACKUP OPEN!, CLOSE TO SAVE, TRYING EVERY 30SECS."
                    )
                    status = True
                    time.sleep(30)
            print("Documento Bakcup actualizado" + "\n")
            cont = 0
        print("Actualizacion cada 5 minutos..." + "\n")
        time.sleep(300)


def InitConf(EmailData):
    book = openpyxl.load_workbook(excelfile)
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    ws = book.worksheets[5]
    inbox = outlook.GetDefaultFolder(6)  
    messages = inbox.Items
    print(book.worksheets[5])
    IterConf(ws, messages, book, EmailData)


def IterConf(ws, messages, book, EmailData):
    row = getRow(ws)
    for message in messages:
        if message.Class == 43 :
            if message.SenderEmailType == "EX" :
                if EmailData.getSubject() in message.Subject and EmailData.getEmail() in message.Sender.GetExchangeUser().PrimarySmtpAddress:
                    print(message.Sender.GetExchangeUser().PrimarySmtpAddress)
                    if EmailData.getBody() in str(message.body).lower():
                        fecha = str(message.SentOn).split()[0].split('-')
                        EmailData.setStart(fecha[1] + "/" + fecha[2] + "/" + fecha[0])
                        EmailData.setCaja(str(message.Subject).split()[1])
                        EmailData.setEnd(str(deltaDate(EmailData.getStart(), -4)))
                        EmailData.setRow(getDateRow(row, ws, EmailData.getStart()))
                        findBoxbyDate(ws, message, EmailData)
            else:
                if EmailData.getSubject() in message.Subject and EmailData.getEmail() in message.SenderEmailAddress:
                    print(message.SenderEmailAddress)
                    if EmailData.getBody() in str(message.body).lower():
                        fecha = str(message.SentOn).split()[0].split('-')
                        EmailData.setStart(fecha[1] + "/" + fecha[2] + "/" + fecha[0])
                        EmailData.setCaja(str(message.Subject).split()[1])
                        EmailData.setEnd(str(deltaDate(EmailData.getStart(), -4)))
                        EmailData.setRow(getDateRow(row, ws, EmailData.getStart()))
                        findBoxbyDate(ws, message, EmailData)
    SaveDoc(excelfile)

def findBoxbyDate(ws, message, EmailData):
    row = EmailData.getRow()
    print("Row:" +str(row))
    if row != None:
        while ws.cell(row=row, column=3).value != EmailData.getEnd() and row >1:
            if EmailData.getCaja() == ws.cell(row=row, column=8).value:
                ws.cell(row=row, column=20).value = EmailData.getStart()
                ws.cell(
                    row=row, column=21).value = str(message.SentOn).split()[1][:8]
                return True
            row -= 1


def getDateRow(row, ws, formdate):
    for x in range(row, 1, -1):
        if('-' in str(ws.cell(row=x, column=3).value)):
            fecha = str(ws.cell(row=x, column=3).value).split()[0].split('-')
            fecha = fecha[1] + "/" + fecha[2] + "/" + fecha[0]

        else:
            fecha = str(ws.cell(row=x, column=3).value)
        if fecha == formdate:
            return x


def deltaDate(strdate, days):
    timeclass = datetime.datetime.strptime(strdate, "%m/%d/%Y")
    return str(timeclass + timedelta(days=days))


def test(box):
    print(box.getEmail())

def main():
    while True:
        print("""
        1. Actualizar Cajas Excel
        2. Actualizado automatico 
        3. Actualizar Confirmacion Daw Vidal
        4. Actualizar Confirmacion Expeditors
        5. Actualizar Nuestro Envio de Documentos 
        6. Exit/Quit
        9. Selecciona si es tu primera vez ejecutando el programa.
        """)
        ans = input("Seleccione una numero: ")
        if ans == '1':
            InitEnv()

        if ans == '2':
            print("Actualizacion cada 5 minutos..." + "\n")
            autoUpdateBox()

        if ans == '3':
            print("Buscando confirmaciones Daw Vidal..." + "\n")
            EmailData = EmailBox("dawvidal.com","SCHNEIDER ELECTRIC P2","listo")
            InitConf(EmailData)

        elif ans == '4':
            print("Buscando confirmaciones Expeditors..." + "\n")            
            EmailData = EmailBox("expeditors.com","P2","entry")
            InitConf(EmailData)

        elif ans == '5':
            print("Buscando confirmaciones Schneider..." + "\n")
            EmailData = EmailBox("schneider-electric.com","P2","adjunto")
            print(EmailData.getEmail())
            InitConf(EmailData)

        elif ans == '9':
            print("Adios")
            exit()
        elif ans == '6':
            EmailData = EmailBox("schneider-electric.com","P2","adjunto")
            test(EmailData)
main()
