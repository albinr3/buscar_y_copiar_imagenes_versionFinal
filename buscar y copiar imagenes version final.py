import shutil, xlrd, glob, xlwt, tkinter, os
from tkinter.filedialog import *
from tkinter import filedialog

#creamos una lista vacia para almacenar los codigos posteriormente
files_to_find = []
lista_codigos_buscar_final = []
lista_verificacion= []
lista_verificacion2= []
lista_codigos_no_encontrados = []

#Aqui ponemos al usuario a elegir la ruta del archivo excel que tiene los codigos
input("Presione ENTER para elegir el archivo excel que contiene los codigos: ")
input_path_archivo_excel = filedialog.askopenfilename(initialdir = os.getcwd(), title = "Seleccionar archivo Excel", filetypes=(("Hoja de cálculo de Microsoft Excel 97-2003", "*.xls"), 
                                                       ("all files", "*.*")))
#Abrimos el archivo excel y cargamos todos los codigos en la lista vacia.
data = xlrd.open_workbook(input_path_archivo_excel)
sheet1 = data.sheet_by_index(0)

#Aqui recorremos la primera columna del archivo excel para obtener los codigos y almacenarlos en una lista
for j in range(sheet1.nrows):
    files_to_find.append(sheet1.cell_value(j, 0))

#Aqui convertimos los codigos a string para que puedan ser manipulados
for elementos in files_to_find:
    lista_codigos_buscar_final.append(str(elementos))

#Los codigos vienen con .0, ejemplo: 3546.0 y con este codigo le quitamos esa parte y queda con 3546
lista_final_final = [ elem[:-2] for elem in lista_codigos_buscar_final ]    

#Aqui seleccionamos la carpeta de donde eleigir las fotos y donde quieres guardar el resultado
input("Presione ENTER para elegir la carpeta donde buscar las imagenes: ")
input_path_buscar_foto = askdirectory()

input("Presione ENTER para elegir la carpeta de destino: ")
ruta_guardar_fotos = askdirectory()
ruta_buscar_fotos = input_path_buscar_foto + "/**/*.jpg"    #a la ruta le agregamos esto para que busque en todas las carpetas y solo archivos jpg
busqueda_de_archivos=glob.glob(ruta_buscar_fotos, recursive=True) #aqui usamos el modulo glob para que busque todas las rutas de archivos jpg y crea una lista con todas

'''
Aqui recorremos todas las rutas y le quitamos el .jpg a todas las imagenes,
luego sustituimos los slash de las rutas por guiones bajos para poder asi dividir
todas las palabras de la ruta, luego usamos un ultimo split para poder dividir
todas las palabras y asi las imagenes /234_126_456.jpg se convertiran en _234_126_456
y posteriormente se convertiran en '234''126''456', asi se podra realizar una
busqueda correctamente.
'''
for i in busqueda_de_archivos:
    lista_verificacion.extend(i.replace(".jpg", "").replace("\\", "_").split("_")) #este lo utilizamos para cargar esa lista para luego hacer la verificacion de los elementos no encontrados
    for codigo_a_buscar in lista_final_final:
        if codigo_a_buscar in i.replace(".jpg", "").replace("\\", "_").split("_"):

            shutil.copy(i, ruta_guardar_fotos) #aqui se copia todos los elementos coincidentes


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

#verifico cuales codigos no fueron encontrados y lo almaceno en una lista
for codigo_no_encotrado in lista_final_final:
    if codigo_no_encotrado not in lista_verificacion2:
        lista_codigos_no_encontrados.append(codigo_no_encotrado)

lista_codigos_no_encontrados = list(map(int, lista_codigos_no_encontrados)) #aqui convierto esos codigos a enteros

#Creo una funcion que sirve para convertir elementos de una lista en pequeñas listas, porque asi es que se puede entrar a u archivo excel
def extractDigits(lst): 
    return [[el] for el in lst]

lista_codigos_no_encontrados = extractDigits(lista_codigos_no_encontrados)

#Aqui creo un archivo con todos los codigos no encontrados
workbook = xlwt.Workbook()
sheet = workbook.add_sheet("hoja1")
for i in range(0,len(lista_codigos_no_encontrados)):
    for j in range(0,1):
        sheet.write(i, j, lista_codigos_no_encontrados[i][j])
workbook.save("archivos_no_encontrados.xls")         



            

    

