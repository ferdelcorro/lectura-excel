# -*- encoding: utf-8 -*-

from os import getcwd
from os.path import join, normpath
"""
para esto necesitan hacer un 
$sudo pip install xlrd
"""
from xlrd import open_workbook
from time import sleep

"""
primero veamos donde estámos parados
"""
path = getcwd()
"""
unimos donde estamos con el archivo que queremos abrir, en este caso yo se
que se encuentra en ese dir
"""
path = normpath(join(path, 'MODELO IMPORTADOR.xls'))
print path

"""
abrimos el archivo exel
"""
wb = open_workbook(path)

"""
sheets = pestañas del archivo exel, veamos solo la primera
"""
s = wb.sheets()[0]
print s.name
max_col = 0
datos = []
"""
recorremos las filas
"""
for row in range(s.nrows):
    filas = []
    """
    recorremos las columnas
    """
    for col in range(s.ncols):
        """
        max_col contiene la máxima cantidad de columnas con datos, para no
        guardar datos vacíos luego que se acaba la tabla
        """
        if (int(col) >= max_col) and (s.cell(row, col).value is not None):
            max_col += 1
        if (int(col) <= max_col) and (s.cell(row, col).value is not None):
            filas.append(s.cell(row, col).value)
    """
    por las dudas, datos es [[datos de la fila 0], [datos de la fila 1], ....]
    """
    datos.append(filas)
print max_col

"""
aquí los imprimimos para poder verlos, hacemos un sleep para verlo más lento
"""
for elem in datos:
    print elem
    sleep(5)