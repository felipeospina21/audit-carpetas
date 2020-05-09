import openpyxl as xl
import os
import getopt
import glob
import sys
import itertools
from time import time
from datetime import datetime
from pathlib import Path, PureWindowsPath

# python acuerdos.py < input_file.txt > output_file.txt
# Unificar oc_audit y dict_acuerdos
# Renombrar dictionary a create_dict o similar
# Renombrar print_dict a dict_to_excel o similar
# Renombrar variables de bucles for (más facil identificación de que son)

inicio = time()
# Variables
meses = {1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril", 5:"Mayo", 6:"Junio", 7:"Julio", 8:"Agosto", 9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"}
now = datetime.now()
year = str(now.year)
mes = now.month - 1
mes = meses.get(mes)
path_oc = Path('C:/Documentos Empresa/OneDrive - MINEROS/Desktop/INDICADORES/{}/{}/Ts_Comprador({}).xlsx'.format(year, mes, mes))
path_informe = Path('C:/Documentos Empresa/OneDrive - MINEROS/Desktop/INDICADORES/{}/{}/informe.xlsx'.format(year, mes))
wb = xl.load_workbook(path_oc)
wb_inf = xl.load_workbook(path_informe)
ws = wb['oc']
ws_inf = wb_inf['informe']
d = {}
d_acuerdos = {}
lista_oc = []
lista_acuerdos = []
lista_concat = []

OC = '4500056418'
# OC = '4500056613'
path_audit = str(Path('//Nasestrella/oc/2020/Nacionales'))
path_audit_oc = str(Path('//Nasestrella/oc/2020/Nacionales/{}'.format(OC)))

lista_excel = []
lista_pdf = []
lista_msg = []
lista_resto = []
lista_ordenes = os.listdir(path_audit_oc)
print(lista_ordenes)
limite = len(lista_ordenes)
for elements in lista_ordenes:
	try:
		nombre, extension = elements.split('.')
		if extension == 'xlsx':
			lista_excel.append(nombre)
		elif extension == 'pdf':
			lista_pdf.append(nombre)
		elif extension == 'msg':
			lista_msg.append(nombre)
		else:
			lista_resto.append(nombre)
	except ValueError:
		lista_resto.append(elements)


print(lista_excel)
print(lista_pdf)
print(lista_msg)
print(lista_resto)
# folder = os.path.exists(path_audit_oc)
# print ('Resultado Método 1: {}'.format(folder))

# for ordenes in lista_ordenes:
# 	if ordenes.find(OC) != -1:
# 		print('Resultado Método 2: {}'.format(ordenes))
# 		print(ordenes.find(OC))
# 		break

# lista = glob.glob(f'{path_audit}/{OC}/*.msg')
# print(lista)
# print('---')
# for elements in lista:
# 	ruta, archivo = elements.split(f'{OC}\\')
# lista2 = [archivo for elements in lista]
# print (lista2)

# lista = [glob.glob(f'{path_audit}/{OC}/*.msg')]
# 4500040018-40019
# 4500055424 - 10091983 10 20 (Porta Asiento y Aguja)
# '4500056855.msg'
#  '4500058281-4500058282'
#print (time() - inicio)