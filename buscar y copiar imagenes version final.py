import shutil, xlrd, glob

#creamos una lista vacia para almacenar los codigos posteriormente
files_to_find = []
lista_codigos_buscar_final = []
lista_verificacion= []
lista_verificacion2= []
#Abrimos el archivo excel y cargamos todos los codigos en la lista vacia.
data = xlrd.open_workbook("prueba.xls")
sheet1 = data.sheet_by_index(0)

for j in range(sheet1.nrows):
    files_to_find.append(sheet1.cell_value(j, 0))




#
for elementos in files_to_find:
    lista_codigos_buscar_final.append(str(elementos))
lista_final_final = [ elem[:-2] for elem in lista_codigos_buscar_final ]    

#Buscamos los codigos en esta ruta

busqueda_de_archivos=glob.glob("C:/Users/Albin Rodriguez/Desktop/FOTOS PRODUCTOS/**/*.jpg", recursive=True)

#aqui si
for i in busqueda_de_archivos:
    lista_verificacion.extend(i.replace(".jpg", "").replace("\\", "_").split("_"))
    for codigo_a_buscar in lista_final_final:
        if codigo_a_buscar in i.replace(".jpg", "").replace("\\", "_").split("_"):

            shutil.copy(i, 'C:/Users/Albin Rodriguez/Pictures/carpeta4')


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

print("##################################################################################################################################################################################")
#verifico cuales codigos no fueron encontrados y lo imprimo
for codigos_no_encotrado in lista_final_final:
    if codigos_no_encotrado not in lista_verificacion2:
        print(f"El codigo: {codigos_no_encotrado} no fue encontrado!")
    
            

    

