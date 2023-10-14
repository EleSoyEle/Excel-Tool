import openpyxl
import numpy as np
from openpyxl import styles
import os
import sys

filename = "" #Direccion del archivo, con su nombre y extension
path = os.path.join(filename)
#Verifica si el archivo existe, en caso de que no, se sale
if not os.path.exists(path):
    sys.exit()
wb = openpyxl.load_workbook(path)
sheet = wb.active

#Genera datos aleatorios y los escribe en el sheet
data = np.random.randint(-100,100,10000)

sheet["A1"] = "Datos"
for i,d in enumerate(data):
    sheet["A{}".format(i+2)] = d

#A partir de aca, se comenzará a elaborar el algoritmo
#Dado que ya tenemos la variable data, trabajaremos con ella
#En caso de querer obtener los datos del sheet, se hace esto
'''
data = []
#Numero de datos manejados
s = 10000
#Suponiendo que la tabla este en la columna A y empecemos desde el 1
#Estamos suponiendo que los datos son enteros, en caso de ser flotantes se cambia int por float
#Si usas flotantes cuidado porque no son fracciones, 1/3 no es 0.333333333330
for i in range(s):
    data.append(int(sheet["A{}".format(i+2)]))
#Se convierte a array de numpy
data = np.array(data)
'''
def random_c():
    txt = ""
    cls = np.random.randint(0,101,[3])
    for n in cls:
        if n == 100:
            txt += "FF"
        elif n<=9:
            txt += "0"+str(n)
        else:
            txt += str(n)
    return txt

p = [False for l in range(len(data))]
c = np.zeros_like(data)

for i,d in enumerate(data):
    if not p[i]:
        for k,dd in enumerate(data[i+1:]):
            if d == -dd and not p[k]:
                c[i] = k
                c[k] = i
                p[i]=True
                p[k]=True
                color = random_c()
                sheet["A{}".format(i+k+3)].fill = styles.PatternFill("solid",start_color=color,end_color=color)
                sheet["A{}".format(i+2)].fill = styles.PatternFill("solid",start_color=color,end_color=color)
                break
        continue
    
#Y listo
wb.save("libro2.xlsx")