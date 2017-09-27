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
	ws = book.worksheets[6]
	inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
	                                    # the inbox. You can change that number to reference                              # any other folder
	messages = inbox.Items
	cajas_folder = inbox.Folders.Item("Sistema de Cajas")
	IterInbox(ws,messages,cajas_folder,book)


def getRow(ws):
	for row in range(900,ws.max_row):
		if ws.cell(row=row,column=1).value == None:
			return row


def CheckRepeat(ws,row,fecha,hora,limit = 30):
	for x in range(row-1, limit,-1):
		localdate = ws.cell(row=x, column=14).value
		localhour = ws.cell(row=x, column=15).value
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
			##	Fill the excell cells with the info from the email.
			##	Each cell will be filled invidivually with the especific value
	CheckStatus = CheckRepeat(ws,row,fecha,hora)

	if(CheckStatus):
		return False

	else:
		today = datetime.date.today()
		ws.cell(row=row, column=1).value = datetime.datetime.now().isocalendar()[1]
		ws.cell(row=row, column=2).value = (today.strftime('%b')).upper()
		ws.cell(row=row, column=3).value = today.strftime("%#m/%#d/2017")

		##	This wil manage the Remesa number by the week of the year.
		if((ws.cell(row=row-1, column = 1).value) != (datetime.datetime.now().isocalendar()[1])):
			ws.cell(row=row, column=4).value = 1
		else:
			ws.cell(row=row, column=4).value = ws.cell(row=row-1, column=4).value + 1

		##	Fill the excell cells with the info from the email.
		##	Each cell will be filled invidivually with the especific value

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
		ws.cell(row=row, column=14).value = fecha
		ws.cell(row=row, column=15).value = hora[:8]
		FormatCells(ws,row)

		return True


def getData(message,ws,row):
	#Read the Email and get the key data from it.
	#This are the variables that will have the data of the email
	body = message.body	
	Caja = body.split()[1]
	Sello = body.split()[4]
	TipoCaja = message.Subject.split()[4]
	fecha = str(message.SentOn).split()[0]
	hora = str(message.SentOn).split()[1]
	if('Chino' in message.body):
		Destino = body.split()[10] + " " + body.split()[11]
		Salida = body.split()[13]
	else:
		Destino = body.split()[10]
		Salida = body.split()[12]

	x=InserData(Caja,Sello,TipoCaja,Destino,Salida,ws,fecha,hora,row)
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

def autoUpdateBox(delay):
	while True:
		InitEnv()
		time.sleep(delay)


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
			delay = input("Actualizacion cada x Segundos:")
			autoUpdateBox(delay)

		elif ans == '5':
			print("Adios")
			exit() 



main()