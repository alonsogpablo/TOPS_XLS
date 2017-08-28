import win32com.client
from win32com.client import Dispatch, constants

const = win32com.client.constants
olMailItem = 0x0
obj = win32com.client.Dispatch("Outlook.Application")
newMail = obj.CreateItem(olMailItem)
newMail.Subject = "Tops Diarios"
newMail.Body = "Hola, buenas. \n\nEstos son los tops de ayer.\n\nUn saludo,\nPablo."
newMail.To = "pablo.alonso@vodafone.com;jesus.paris@celfinet.com"
attachment1 = r"C:\\Users\\palonso0\\PycharmProjects\\TOPS_XLS\\TOPS.xlsx"

newMail.Attachments.Add(Source=attachment1)
newMail.display()

newMail.send()