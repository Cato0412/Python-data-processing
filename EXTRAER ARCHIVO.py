
import pandas as pd
import os

ruta = "C:\\Users\\lapmxdf558\\Documents\\JUAN\\BONOS PY\\"

# Buscar archivos que contengan 'Rutero'
archivos = [f for f in os.listdir(ruta) if 'Rutero' in f]

if archivos:
    archivo_rutero = os.path.join(ruta, archivos[0])
    dfrut = pd.read_excel(archivo_rutero, sheet_name='RUTERO',skiprows=4)
    print("Archivo cargado:", archivo_rutero)
else:
    print("No se encontró ningún archivo con 'Rutero' en el nombre.")

archivos = [f for f in os.listdir(ruta) if 'Personal' in f]

if archivos:
    archivo_personal = os.path.join(ruta, archivos[0])
    dfper = pd.read_excel(archivo_personal, sheet_name='PERSONAL',skiprows=4)
    palabra=archivo_personal.replace(ruta,"")
    print("Archivo cargado:", palabra)
else:
    print("No se encontró ningún archivo con 'Rutero' en el nombre.")


