"""
2025.05.19.
El objetivo de este módulo es almacenar los valores de un excel.
Se busca iterar el xml 'sheet1' e ir agregando los valores a una lista_var. Y al cambiar de row, añadirla a matriz y luego lista_var = []
"""
import zipfile
import time
import tkinter as tk
from tkinter import filedialog
root = tk.Tk()
root.withdraw()
#ruta_completa = r"C:\Users\criis\OneDrive\Documentos\Coding\Ejemplo Datos_5000_rows.xlsm"
#ruta_completa = r"C:\Users\criis\OneDrive\Documentos\Coding\otros\Ejemplo Datos.xlsm"
#ruta_completa = r"C:\Users\criis\OneDrive\Documentos\Coding\Ejemplo Transacciones por report.xlsm"

##Esto de abajo busca mostrar la información al final
def show_info():
    from tkinter import simpledialog
    respuesta = messagebox.askyesnocancel("Impresión de datos.", "¿Desea imprimir la matriz?")
    if respuesta is True:
        import sys
        ubicacion_inicial = simpledialog.askinteger("Fila inicial", "¿Desde qué fila quieres imprimir?") or 1
        ubicacion_objetivo = simpledialog.askinteger("Fila final", "¿Hasta qué fila quieres imprimir?") or 20
        #ubicacion_objetivo = int(input("Rows final: "))+1 ##Esto te pregunta hasta qué row quieres extraer
        #print(matriz[0])
        print("\t".join(str(valor) for valor in matriz[0]))
        for i in range(ubicacion_inicial, ubicacion_objetivo+ubicacion_inicial):
            #print(matriz[i])
            fila = [str(valor) for valor in matriz[i]]
            print("\t".join(fila))
    else:
        pass


ruta_completa = filedialog.askopenfilename(
    initialdir=r"C:\Users\criis\Documents\Coding",
    title="Selecciona el archivo de Excel",
    filetypes=[("Archivos de Excel", "*.xlsm *.xlsx"), ("Todos los archivos", "*.*")]
)
if not ruta_completa:
    ruta_completa = r"C:\Users\criis\Documents\Coding\Ejemplo Transacciones por report.xlsm"
#ruta_completa = r"C:\Users\criis\OneDrive\Documentos\Coding\Ejemplo Transacciones por report.xlsm"
Timer0 = time.time()
print(f"Usando el archivo {ruta_completa}")


##Objetos base antes de iterar:
lectura = zipfile.ZipFile(ruta_completa, "r")
sharedstring_xml = lectura.read("xl/sharedStrings.xml")
sharedstr_list = []
##Creamos una lista con todos los valores de sharedstrings
import re
pattern = re.compile(b'<t>(.*?)</t>')
valores_en_bytes = pattern.findall(sharedstring_xml)
for v in valores_en_bytes:
    sharedstr_list.append(v.decode('utf-8'))
#print(sharedstr_list[:50])

##Creamos una lista de listas con todos los valores de sheet_xml
sheet_xml = lectura.read("xl/worksheets/sheet1.xml") ##Esto contiene toda la información de la hoja

##Ahora que tenemos los <sharedstr>, quiero iterar a través del sheet_str y almacenar los valores buscados: zReportID & Amount, columnas: a & e
##Encontramos lastrow & lastcol  en bytes
pos1 = sheet_xml.find(b'<dimension ref="')
pos2 = sheet_xml.find(b'"', pos1 + 18)
ref = sheet_xml[pos1 + 17 : pos2]
StrVar1, StrVar2 = ref.split(b':')
ColBytes = StrVar2[:1] #Obtiene la letra de la columna as bytes
NumBytes = StrVar2[1:] ##Obtiene LastRow as bytes
lastcol = ColBytes.decode('utf-8') #Transforma a string
lastcol = ord(lastcol) - 64
lastrow = int(NumBytes) #Transforma a integer

#lettercolumns = ["A","B","C","D","E","F","G","H", "I", "J", "K", "L", "M", "N", "O","P","Q","R","S","T","U","V","W","X","Y","Z"]
#matriz = [[0] * lastcol for i in range(lastrow)]  ##Crea una matriz, redim(1 to Lastrow, 1 to lastcol)
matriz = []
row_end = 0
count = 0
row_start = sheet_xml.find(b'<row r=', row_end)
# col_start = sheet_xml.find(b'<c r=',row_start)
#col_end = sheet_xml.find(b'<c', col_start, row_end)
##Deseaba referenciar una sacción en el xml para sobre él buscar y evitar varios '.find'; o simplemente usar varios buscar.

for i in range(0, lastrow):
    lista_var = []
    row_end = sheet_xml.find(b'</row', row_start)
    col_start = sheet_xml.find(b'<c r=',row_start)
    # print(f"Row = {sheet_xml[row_start:row_end+3].decode('utf-8')}")
    while col_start != -1:
        col_end = sheet_xml.find(b'/', col_start, row_end)+2 ##Esto es lo que haría normalmente, definir inicio y fin.
        # print(f"Col data = {sheet_xml[col_start:col_end].decode("utf-8")}")
        text_type = sheet_xml.find(b't="s', col_start, col_end)
        pos1 = sheet_xml.find(b'<v>', col_start, col_end)
        pos2 = sheet_xml.find(b'</v', pos1, col_end)
        strvar = sheet_xml[pos1+3:pos2].decode('utf-8')
        if pos1 == -1:
            strvar = 0
        if text_type == -1: ##si es integer...
            strvar = int(strvar)
        else:
            strvar = sharedstr_list[int(strvar)]
        lista_var.append(strvar)
        count = count + 1
        if count > 10000000:
            exit()
        col_start = sheet_xml.find(b'<c r=', col_end, row_end)
    row_start = row_end + 3 ##</row> + 6
    #row_end = sheet_str.find(b'<row r =', row_start) -1
    # row_end = sheet_str.find(b'</row', row_start)
    # for j in range(0, lastcol): ##Buscaremos cada valor.
        # columna_actual = ord(j)
        # col_start = sheet_str.find(b"<c r=")
        # col_start = sheet_str.find(f'<c r="{columna_actual}',row_start, row_end)

print("\n")
print(f"Lectura de archivo completa, la información se guardó como [Matriz] & [df_matriz], como matriz & dataframes respectivamente")
#print(df_matriz.iloc[0:10])
#for fila in matriz[:10]:
 #   print(fila)
Timer1 = time.time()
ExecTime = Timer1 - Timer0
print(ExecTime)
#exit()
print(f"Iteraciones totales = {count:,.0f}")
print("-- End -- ")
print("\n")

def main():
    pass
if __name__ == '__main__':
    main()

