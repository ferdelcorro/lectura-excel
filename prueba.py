# -*- encoding: utf-8 -*-

from os import getcwd
from os.path import join, normpath
"""
para esto necesitan hacer un 
$sudo pip install xlrd
$sudo pip install xlwt
"""
from xlrd import open_workbook
from xlwt import Workbook
from time import sleep


#primero veamos donde estámos parados
path = getcwd()

#unimos donde estamos con el archivo que queremos abrir, en este caso yo se
#que se encuentra en ese dir

path = normpath(join(path, 'MODELO IMPORTADOR.xls'))
print path


#abrimos el archivo exel
wb = open_workbook(path)


#sheets = pestañas del archivo exel, veamos solo la primera
s = wb.sheets()[0]
print s.name
datos = []

#recorremos las filas
for row in range(s.nrows):
    filas = []
    
    #recorremos las columnas
    for col in range(s.ncols):
        filas.append(s.cell(row, col).value)
    
    #por las dudas, datos es [[datos de la fila 0], [datos de la fila 1], ....]
    datos.append(filas)


#aquí los imprimimos para poder verlos, hacemos un sleep para verlo más lento
#for elem in datos:
    #print elem
    #sleep(2)


#ahora vamos a escribir otro exel
workbook = Workbook() #Creamos un objeto exel
sheet = workbook.add_sheet('test') #Le adjuntamos una hoja


#recorremos todos los datos y escribimos
n_fil = 0
for fil in datos:
    n_col = 0
    for col in fil:
        sheet.write(n_fil, n_col, col)
        n_col += 1
    n_fil += 1

#guardamos
workbook.save('test.xls')