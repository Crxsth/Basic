# Matrix_creation_bytes /readme format by chatgpt as it's my first project :)

**Autor:** Cristh
**Fecha de última modificación:** 18 de Agosto de 2025

## Description

This project extracts information from an xlsx or xlsm and creates a list of lists.
It takes on average 1.9s to read 1.050 Million cells.

I started this project because with pandas and other libraries (that at the time I didn't know very well) take up to 20s just to read the excel.



## How does this work?
- If used natively, open a folder and select the folder to read. If not, you can call it as a  function to read a file.
- It creates a list of lists (I call it as 'matrix') of all the strings of the excel.
- Iterates over the information and adds the information to the matrix, it converts formulas to strings.


## Requisitos

- Python 3.x
- No external, only repositories: (`zipfile`, `re`, `tkinter`, `time`).

## Uso
1. Just import the function: "leer_file(ruta_completa)" and set the full route as 'ruta_completa'.
2. The script will process the file and will return a list of lists named: "matriz"

## Notes
- It's not expected to read a file with multiple sheets (I haven't thought about it until now, July 22, 2025).
- It doesn't read 'inline strings' yet.

#Hardware:
- Laptop: Asus Vivovook Go 15
- CPU: AMD Ryzen 5 7520u with Radeon Graphics (8 CPUs), ~2.8GHz
- 16GB RAM

## Ejemplo de uso
```python
# Ejecuta desde consola:
python Matrix_creation_bytes.py


