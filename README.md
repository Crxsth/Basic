# Matrix_creation_bytes /readme format by chatgpt as it's my first project :)

**Autor:** Cristh
**Fecha de última modificación:** 18 de julio de 2025

## Description

This project extracts information from an xlsx or xlsm and creates a list of lists.
It takes 1.458s on average to read 1.050 Million cells.

I started this project because with pandas and other libraries (that at the time I didn't know very well) take up to 20s just to read the excel.


## How does this work?
- We use 'tkinter' to select a file (we can't change this to save 0.2s).
- It creates a list of lists (I call it as 'matrix') of all the strings of the excel.
- Iterates over the information and adds the information to the matrix, whether integer or string.


## Requisitos

- Python 3.x
- No requiere paquetes externos, solo módulos estándar (`zipfile`, `re`, `tkinter`, `time`).

## Uso
1. The script can be imported to select a file to read (but I haven't tested that yet).
2. The script will process the file and create a list of lists. 

## Personalización
- It's expected to define a file path to use instead of use the tkinter.

- Puedes modificar el archivo para elegir otra hoja (`sheet2.xml`, etc.).
- Si necesitas solo algunas columnas, ajusta la lógica de búsqueda de celdas.
- Si quieres exportar los resultados, puedes convertir la matriz a CSV o DataFrame.

## Ejemplo de uso
```python
# Ejecuta desde consola:
python Matrix_creation_bytes.py


