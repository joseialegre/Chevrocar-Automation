from openpyxl import load_workbook
from colorama import Fore, Style

import csv
from openpyxl.styles import Alignment

archivo1 = input("ingrese nombre de archivo de Exportar:     ")
archivo2 = input("ingrese nombre de archivo de TiendaNube:     ")

archivo1= archivo1+".xlsx"
archivo2= archivo2+".csv"

wb1 = load_workbook(archivo1)

sheet1 = wb1.active

azul = Fore.BLUE
rojo = Fore.RED
verde = Fore.GREEN
amarillo = Fore.YELLOW
reset = Style.RESET_ALL


with open(archivo2, 'r') as csvfile:
    reader = csv.reader(csvfile, delimiter=';')
    lista_tiendanube = list(reader)

    for filaExportar in sheet1.iter_rows(min_row=2, values_only=True):
        for index, filaTienda in enumerate(lista_tiendanube):
            # Comparamos SKU
            if str(filaTienda[16]) == str(filaExportar[0]):
                print(amarillo + "Copiando los valores de... " +azul+ str(filaTienda[16]) + reset)
                print("Antes: " + rojo + str(filaTienda[9]) + reset + " -->" +  " Después: " + azul + verde + str(filaExportar[2]) + reset)
                lista_tiendanube[index][9] = filaExportar[2]

with open(archivo2, 'w', newline='') as csvfile:
    writer = csv.writer(csvfile, delimiter=';')
    writer.writerows(lista_tiendanube)


wb1.close()

# SKU de Exportar columna A = 0
# SKU de tiendanube columna Q = 16

# PRECIO para Tienda columna J (numero 9) es el precio
# PRECIO de Exportar columna C (numero 2)