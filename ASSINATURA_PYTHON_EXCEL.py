import PyPDF2
from win32com import client
import os


#Função para encerrar tarefa Excel
def closeFile():

    try:
        os.system('TASKKILL /F /IM excel.exe')

    except Exception:
        print("KU")

#Definindo locais de trabalho
xlApp = client.Dispatch("Excel.Application")
books = xlApp.Workbooks.Open(r'') #Local do arquivo excel
ws = books.Worksheets[0]
ws.Visible = 1
ws.ExportAsFixedFormat(0, r'') #Local para salvar o pdf

closeFile()

#Lendo arquivos pdf para mesclagem
arqui1 = open('','rb')
arqui2 = open('','rb')
dadosImg1 = PyPDF2.PdfFileReader(arqui1)
dadosImg2 = PyPDF2.PdfFileReader(arqui2)

merge = PyPDF2.PdfFileMerger()

merge.append(dadosImg1)
merge.append(dadosImg2)

nome = input('Digite o nome do documento: ')

merge.write(r"{}.pdf".format(nome))