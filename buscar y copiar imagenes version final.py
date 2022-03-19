import shutil, xlrd, glob, xlwt, tkinter
from tkinter.filedialog import *

#creamos una lista vacia para almacenar los codigos posteriormente
files_to_find = []
lista_codigos_buscar_final = []
lista_verificacion= []
lista_verificacion2= []
lista_codigos_no_encontrados = []
#Abrimos el archivo excel y cargamos todos los codigos en la lista vacia.
data = xlrd.open_workbook("prueba.xls")
sheet1 = data.sheet_by_index(0)

for j in range(sheet1.nrows):
    files_to_find.append(sheet1.cell_value(j, 0))


for elementos in files_to_find:
    lista_codigos_buscar_final.append(str(elementos))
lista_final_final = [ elem[:-2] for elem in lista_codigos_buscar_final ]    

#Buscamos los codigos en esta ruta
input("Presione ENTER para elegir la carpeta donde buscar las imagenes")
input_path_buscar_foto = askdirectory()
input("Presione ENTER para elegir la carpeta de destino")
ruta_guardar_fotos = askdirectory()
ruta_buscar_fotos = input_path_buscar_foto + "/**/*.jpg"
busqueda_de_archivos=glob.glob(ruta_buscar_fotos, recursive=True)

#aqui si
for i in busqueda_de_archivos:
    lista_verificacion.extend(i.replace(".jpg", "").replace("\\", "_").split("_"))
    for codigo_a_buscar in lista_final_final:
        if codigo_a_buscar in i.replace(".jpg", "").replace("\\", "_").split("_"):

            shutil.copy(i, ruta_guardar_fotos)


#creo una funcion que sirve para verificar si un elemento es un numero
def es_numero(n):
    try:
        float(n)
    except ValueError:
        return False
    return True
#convierto todos las rutas de fotos a sus codigos, para que queden los codigos limpios
for u in lista_verificacion:
    if es_numero(u):
        lista_verificacion2.append(u)

#verifico cuales codigos no fueron encontrados y lo imprimo
for codigo_no_encotrado in lista_final_final:
    if codigo_no_encotrado not in lista_verificacion2:
        lista_codigos_no_encontrados.append(codigo_no_encotrado)

lista_codigos_no_encontrados = list(map(int, lista_codigos_no_encontrados))


def extractDigits(lst): 
    return [[el] for el in lst]

lista_codigos_no_encontrados = extractDigits(lista_codigos_no_encontrados)

workbook = xlwt.Workbook()
sheet = workbook.add_sheet("hoja1")
for i in range(0,len(lista_codigos_no_encontrados)):
    for j in range(0,1):
        sheet.write(i, j, lista_codigos_no_encontrados[i][j])
workbook.save("archivos_no_encontrados.xls")         



            

    

