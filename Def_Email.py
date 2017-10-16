import win32com.client
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font
import datetime
import time
import copy
from openpyxl.styles import Alignment

def SaveDoc(book):
    status = True
    while status:
        try:
            book.save('\\\\WSMX02402FP\\Shared\\IMP-EXP\INTERNATIONAL TRADE\\RK & Activos\\EXPORT P2\\2017\HORARIOS EXP P2.xlsx')
            status = False
        except:
            print("ERROR, ARCHIVO DE HORARIOS ABIERTO!, CIERRE PARA GUARDAR, VOLVIENDO A INTENTAR EN 30 SEGUNDOS.")
            status = True
            time.sleep(30)
    print("DOCUMENTO GUARDADO")


def InitEnv():
    print("Buscando cajas nuevas..." + "\n")
    book = openpyxl.load_workbook('\\\\WSMX02402FP\\Shared\\IMP-EXP\INTERNATIONAL TRADE\\RK & Activos\\EXPORT P2\\2017\HORARIOS EXP P2.xlsx')
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    ws = book.worksheets[5]
    inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
                                        # the inbox. You can change that number to reference                              # any other folder
    messages = inbox.Items
    cajas_folder = inbox.Folders.Item("Sistema de Cajas")
    print(book.worksheets[5])
    IterInbox(ws,messages,cajas_folder,book)


def getRow(ws):
    for row in range(900,ws.max_row):
        if ws.cell(row=row,column=1).value == None:
            return row


def CheckRepeat(ws,row,fecha,hora,limit = 30):
    for xrow in range(row-1, limit,-1):
        localdate = ws.cell(row=xrow, column=14).value
        localhour = ws.cell(row=xrow, column=15).value
        if(fecha == localdate and localhour == hora[:8]):
            return True
    return False


def FormatCells(ws,row):
    for x in range(1,21):
        cell = ws.cell(row=row, column=x) 
        cell.font = Font(name="Calibri", size = 9)
        cell.alignment = Alignment(horizontal='center', vertical='bottom')
        if(x == 3 or x == 4 or x == 14):
            cell.font = Font(name="Calibri", size = 9, bold = True)


def IterInbox(ws,messages,cajas_folder,book):
    email = "AdminCajasMty"
    subject = 'P2- Reporte de Caja'
    subjectAd = 'NACIONAL'
    row = getRow(ws)
    rowAUx = row
    for message in messages:
        if subject in message.Subject and email in message.SenderEmailAddress and not(subjectAd in message.Subject):
            x = getData(message,ws,row)
            if(x):
                row += 1
            message.Move(cajas_folder)
    if(rowAUx != row):
        SaveDoc(book)
    else:
        print("No hubo cajas por guardar!" )


def InserData(Caja,Sello,TipoCaja,Destino,Salida,ws,fecha,hora,row):
            ##    Fill the excell cells with the info from the email.
            ##    Each cell will be filled invidivually with the especific value
    CheckStatus = CheckRepeat(ws,row,fecha,hora)

    if(CheckStatus):
        return False
    else:
        fecha = fecha.split('-')
        formdate = fecha[1] + "/" + fecha[2] + "/" + fecha[0]
        today = datetime.date.today()
        ws.cell(row=row, column=1).value = datetime.datetime.now().isocalendar()[1]
        ws.cell(row=row, column=2).value = (today.strftime('%b')).upper()
        ws.cell(row=row, column=3).value = today.strftime("%#m/%#d/2017")

        ##    This wil manage the Remesa number by the week of the year.
        if((ws.cell(row=row-1, column = 1).value) != (datetime.datetime.now().isocalendar()[1])):
            ws.cell(row=row, column=4).value = 1
        else:
            for xrow in range(row-1,0,-1):
                if ws.cell(row=xrow, column=4).value != None:
                    ws.cell(row=row, column=4).value = ws.cell(row=xrow, column=4).value + 1
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
        t = time.strptime( Salida , "%H:%M")
        ws.cell(row=row, column=11).value = time.strftime( "%I:%M %p", t)
        ws.cell(row=row, column=14).value = formdate
        ws.cell(row=row, column=15).value = hora[:8]
        FormatCells(ws,row)

        return True


def getData(message,ws,row):
    #Read the Email and get the key data from it.
    #This are the variables that will have the data of the email
    body = str(message.body).split()    
    Caja = body[1]
    Sello = body[4]
    TipoCaja = message.Subject.split()[4]
    fecha = str(message.SentOn).split()[0]
    hora = str(message.SentOn).split()[1]
    if('Chino' in message.body):
        Destino = body[10] + " " + body[11]
        Salida = body[13]
    else:
        Destino = body[10]
        Salida = body[12]

    x = InserData(Caja,Sello,TipoCaja,Destino,Salida,ws,fecha,hora,row)
    if(x):
        print("_______Caja Guardada_______" + "\n")
        print(Caja + " " + Sello + " " + Destino + " " + Salida + " " + TipoCaja)
        print("___________________________" + "\n")
        return True
    else:
        print("_______Caja Repetida________" + "\n")
        print(Caja + " " + Sello + " " + Destino + " " + Salida + " " + TipoCaja)
        print("____________________________" + "\n")
        return False


def autoUpdateBox():
    Cont=0
    while True:
        book = openpyxl.load_workbook('\\\\WSMX02402FP\\Shared\\IMP-EXP\INTERNATIONAL TRADE\\RK & Activos\\EXPORT P2\\2017\HORARIOS EXP P2.xlsx')
        Cont+=1
        InitEnv()
        if(Cont==30):
                status = True
                while status:
                    try:
                        book.save('\\\\WSMX02402FP\\Shared\\IMP-EXP\INTERNATIONAL TRADE\\RK & Activos\\EXPORT P2\\2017\HORARIOS EXP P2 - BACKUP.xlsx')
                        status = False
                    except:
                        print("ERROR, ARCHIVO DE BACKUP ABIERTO!, CIERRE PARA GUARDAR, VOLVIENDO A INTENTAR EN 30 SEGUNDOS.")
                        status = True
                        time.sleep(30)
                print("Documento Bakcup actualizado"+ "\n")
                Cont=0
        print("Actualizacion cada 5 minutos..." + "\n")
        time.sleep(300)


def InitDaw():
    print("Buscando cajas nuevas..." + "\n")
    book = openpyxl.load_workbook('\\\\WSMX02402FP\\Shared\\IMP-EXP\INTERNATIONAL TRADE\\RK & Activos\\EXPORT P2\\2017\HORARIOS EXP P2.xlsx')
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    ws = book.worksheets[5]
    inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
                                        # the inbox. You can change that number to reference                              # any other folder
    messages = inbox.Items
    print(book.worksheets[5])
    IterDaw(ws,messages,book)


def IterDaw(ws,messages,book):
    email = "dawvital.com"
    subject = 'P2'
    row = getRow(ws)
    for message in messages:
        if subject in message.Subject and email in message.SenderEmailAddress :
            if 'listo' in str(message.body).lower():
                fecha = str(message.SentOn).split('-')
                #2017-10-17
                #10/17/2017
                formdate = fecha[1] + "/" + fecha[2] + "/" + fecha[0]
                timeclass = datetime.datetime.strptime(formdate, "%m-%d-%Y")
                print(str(timeclass))
                caja = str(message.Subject).split()[1]


def test():
    fecha = '10/17/2017'
    #2017-10-17
    #10/17/2017
    #formdate = fecha[1] + "/" + fecha[2] + "/" + fecha[0]
    timeclass = datetime.datetime.strptime(fecha, "%m-%d-%Y")
    print(str(timeclass))
    caja = str(message.Subject).split()[1]


def main():
    while True:    
        print ("""
        1. Actualizar Cajas Excel
        2. Actualizado automatico
        3. Actualizar Confirmacion Daw Vital
        4. Actualizar Confirmacion Expeditors
        5. Exit/Quit
        9. Selecciona si es tu primera vez ejecutando el programa.
        """)
        ans = input("Seleccione una numero: ")
        if ans == '1':
            InitEnv()

        if ans == '2':
            print("Actualizacion cada 5 minutos..."+ "\n")
            autoUpdateBox()

        elif ans == '5':
            print("Adios")
            exit() 
        elif ans == '6':
            InitDaw() 

main()