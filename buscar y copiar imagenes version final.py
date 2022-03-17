import shutil, xlrd, glob

#creamos una lista vacia para almacenar los codigos posteriormente
files_to_find = []
lista_codigos_buscar_final = []


#Abrimos el archivo excel y cargamos todos los codigos en la lista vacia.
data = xlrd.open_workbook("libro4.xls")
sheet1 = data.sheet_by_index(0)

for i in range(sheet1.nrows):
    files_to_find.append(sheet1.cell_value(i, 0))

for elementos in files_to_find:
    lista_codigos_buscar_final.append(str(elementos))
listaActualizada = [ elem[:-2] for elem in lista_codigos_buscar_final ]    
#Buscamos los codigos en esta ruta

yu=glob.glob("C:/Users/Albin Rodriguez/Desktop/FOTOS PRODUCTOS/**/*.jpg", recursive=True)


for i in yu:
    for kaka in listaActualizada:
        if kaka in i.replace(".jpg", "").replace("\\", "_").split("_"):
            shutil.copy(i, 'C:/Users/Albin Rodriguez/Pictures/carpeta4')

