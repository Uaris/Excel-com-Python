# Criando um arquivo no Excel com o Python.
# Obs: "Import xlsxwriter as xls" é uma convenção pessoal que usei para simplificar as coisas.

import xlsxwriter as xls
import os

#definindo o local da planilha
path = r'C:\Users\User\Documents\MAP\Udemy Power BI\Primeiro arquivo.xlsx'

planilha = xls.Workbook(path)
trabalho = planilha.add_worksheet() #cria uma nova planilha em branco com o nome "sheet 1"


#Adicionando dados
trabalho.write("A1","Nome")
trabalho.write("B1","Idade")
trabalho.write("A2","Evandro")
trabalho.write("B2","25")

#fechando a planilha no código
planilha.close()

#abrindo a planilha no Excel
os.startfile(path)
