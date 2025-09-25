"""
2025.08.12. Tiempo de ejecución en un archivo de 2'044,616 datos == 2.13s. El archivo ahora revisa el xml en diversas formas y almacena en una matriz_
correctamente.

2025.05.19.
El objetivo de este módulo es almacenar los valores de un excel.
Se busca iterar el xml 'sheet1' e ir agregando los valores a una lista_var. Y al cambiar de row, añadirla a matriz y luego lista_var = []
"""
import zipfile
import time
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
import matplotlib.pyplot as plt

    

def show_info(): ##Lo dió chatgpt 100%. Es para mostrar la matriz.
    for fila in matriz[:20]:
        print(fila)
    # MAX_ROWS = 20
    # root = tk.Tk(); root.withdraw()

    # resp = messagebox.askyesnocancel("Impresión de datos", "¿Desea imprimir la matriz?", parent=root)
    # if resp is not True:
        # root.destroy(); return
    # if not matriz or not isinstance(matriz[0], (list, tuple)):
        # messagebox.showwarning("Aviso", "La matriz está vacía o no es válida.", parent=root)
        # root.destroy(); return

    # ini = simpledialog.askinteger("Fila inicial", "¿Desde qué fila quieres imprimir? (mín=1)", minvalue=1, initialvalue=1, parent=root)
    # if ini is None:
        # root.destroy(); return
    # fin = simpledialog.askinteger("Fila final", "¿Hasta qué fila quieres imprimir?",
                                  # minvalue=ini,
                                  # initialvalue=min(ini + MAX_ROWS - 1, len(matriz) - 1),
                                  # parent=root)
    # if fin is None:
        # root.destroy(); return

    # fin = min(fin, len(matriz) - 1, ini + MAX_ROWS - 1)
    # headers = [str(v) for v in matriz[0]]
    # data = [[str(v) for v in row] for row in matriz[ini:fin + 1]]
    # root.destroy()

    # fig, ax = plt.subplots(figsize=(min(12, 1.2 * len(headers)), 0.6 * (len(data) + 2)))
    # ax.axis('off')
    # tbl = ax.table(cellText=data, colLabels=headers, loc='center')
    # tbl.auto_set_font_size(False)
    # tbl.set_fontsize(9)
    # tbl.scale(1, 1.2)
    # ax.set_title(f"Matriz filas {ini}-{fin} (máx {MAX_ROWS})", pad=10)
    # plt.tight_layout()
    # plt.show()

def leer_file(ruta_completa):
    print(f"Usando el archivo {ruta_completa}")
    ##Objetos base antes de iterar:
    lectura = zipfile.ZipFile(ruta_completa, "r")
    sharedstring_xml = lectura.read("xl/sharedStrings.xml")
    sharedstr_list = []
    ##Creamos una lista con todos los valores de sharedstrings
    import re
    pattern = re.compile(b'<t[^>]*>(.*?)</t>')
    valores_en_bytes = pattern.findall(sharedstring_xml)
    for v in valores_en_bytes:
        sharedstr_list.append(v.decode('utf-8'))

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
    Timer3 = time.time()
    # print(f"{Timer3-Timervar} Tiempo antes del bucle")

    #lettercolumns = ["A","B","C","D","E","F","G","H", "I", "J", "K", "L", "M", "N", "O","P","Q","R","S","T","U","V","W","X","Y","Z"]
    #matriz = [[0] * lastcol for i in range(lastrow)]  ##Crea una matriz, redim(1 to Lastrow, 1 to lastcol)
    matriz = []
    timer_list = []
    row_end = 0
    count = 0
    row_start = sheet_xml.find(b'<row r=', row_end)
    countvar = 0
    # col_start = sheet_xml.find(b'<c r=',row_start)
    #col_end = sheet_xml.find(b'<c', col_start, row_end)
    avg_time = 0.505301/10000
    ##Deseaba referenciar una sacción en el xml para sobre él buscar y evitar varios '.find'; o simplemente usar varios buscar.
    for i in range(0, lastrow):
        lista_var = []
        row_end = sheet_xml.find(b'</row', row_start)
        col_start = sheet_xml.find(b'<c r=',row_start)
        # print(f"Row = {sheet_xml[row_start:row_end+3].decode('utf-8')}")
        while col_start != -1:
            col_end = sheet_xml.find(b'/', col_start, row_end)+2 ##Esto es lo que haría normalmente, definir inicio y fin.
            
            ##Esto de abajo es en caso de que se salte un <c>, ejemplo row=1, <c r=B2 instead of < r=A2 
            # try:
            cell_var = ord(sheet_xml[col_start +6: col_start +7].decode("utf-8"))-64 -1
            count3 = 0
            while len(lista_var) < cell_var:
                # print(cell_var)
                lista_var.append("")
                count3 +=1 
                if count3 >27:
                    print("Fallaste we")
                    exit()
            
            ##Identificamos el valor entre <v>
            #text_time1 = time.time()
            posvar = sheet_xml.find(b'<f', col_start,col_end)
            if posvar != -1: ##if >t="str"< then:
                col_end = sheet_xml.find(b'/c', col_start, row_end) + 2
            pos1 = sheet_xml.find(b'<v>', col_start, col_end)
            pos2 = sheet_xml.find(b'</v', pos1, col_end)
            if pos1 == -1 or pos2==-1: ##Si no hay valor alguno...
                strvar = ""
            else:
                strvar = sheet_xml[pos1+3:pos2].decode('utf-8') ##Avg = 0.505301/10000
            fila = sheet_xml[col_start:col_end]
            #text_time2 = time.time()
            #text_time = text_time2-text_time1
            #timer_list.append(text_time)
            
            ##Identificamos si tiene un text_type... ##Tiempo de ejecución promedio de este bucle:  0.003146e-6
            text_type = sheet_xml.find(b't=', col_start, col_end) ##
            if text_type != -1:
                pos_text = sheet_xml.find(b'"',text_type+4, col_end)
                valorvar = sheet_xml[text_type+3:pos_text].decode("utf-8")
                if valorvar == "s":
                    strvar = sharedstr_list[int(strvar)] ##Si es sharedstring, accedemos a él.
            lista_var.append(strvar)
            count = count + 1
            if count > 3000000:
                exit()
            col_start = sheet_xml.find(b'<c r=', col_end, row_end)
        row_start = sheet_xml.find(b'<row', row_end)
        matriz.append(lista_var)
    print(f"Iteraciones totales = {count:,.0f}")
    return matriz

if __name__ == '__main__':
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    ruta_completa = filedialog.askopenfilename(
        initialdir=r"C:\Users\criis\Documents\Coding",
        title="Selecciona el archivo de Excel",
        filetypes=[("Archivos de Excel", "*.xlsm *.xlsx"), ("Todos los archivos", "*.*")]
    )
    if not ruta_completa:
        ruta_completa = null
    
    Timer0 = time.time()
    matriz = leer_file(ruta_completa)
    Timer1 = time.time()
    print("\n")
    show_info()
    
    ExecTime = Timer1 - Timer0
    print(f"Lectura de archivo completa, la información se guardó como [matriz]")
    print(f"Tiempo de lectura: {ExecTime}")


    

    print("-- End -- ")
    print("\n")
    # import csv
    # with open("xlsx_reader_test.csv", mode="w", newline="", encoding="utf-8") as archivo_csv: ##Esto crea un csv con los datos, por si quisieramos corroborar
        # writer = csv.writer(archivo_csv)
        # writer.writerows(matriz)