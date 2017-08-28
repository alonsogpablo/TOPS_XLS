from win32com.client import Dispatch
from openpyxl import load_workbook

path='C:\\Users\\palonso0\\PycharmProjects\\TOPS_XLS'

excel = Dispatch('Excel.Application')
wb = excel.Workbooks.Open(path+'\\'+'TOPS.xlsx')

#Activate second sheet
count = wb.Sheets.count

for i in range(count):
    i=i+1

    excel.Worksheets(i).Activate()

    #Autofit column in active sheet
    excel.ActiveSheet.Columns.AutoFit()



#Or simply save changes in a current file
wb.Save()
wb.Close()

wb1=load_workbook(path+'\\'+'TOPS.xlsx')
wb1.active=0
wb1.save(path+'\\'+'TOPS.xlsx')