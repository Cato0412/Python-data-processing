import pandas as pd
import numpy as np
import os as os
import random 

totales={}

RUTA=r"C:\\Users\\lapmxdf558\\Documents\\JUAN\\BONOS PY\\JULIOESC\\"

#CARGAR RUTERO
archivos = [f for f in os.listdir(RUTA) if 'Rutero' in f]
if archivos:
    archivo_rutero = os.path.join(RUTA, archivos[0])
    dfrut = pd.read_excel(archivo_rutero, sheet_name='RUTERO',skiprows=4)
    palabra=archivo_rutero.replace(RUTA,"")
    print("Archivo cargado:", palabra)
else:
    print("No se encontró ningún archivo con 'Rutero' en el nombre.")

dfrut["CLIENTE"]="BIC"

CFIJAS=["Usuario Virtual","Usuario APP Promotor","Zona - Region","Codigo RO","Nombre de Tienda"]

for i in range(1,6):
    
    i=str(i)

    COLVARIABLESS=dfrut.loc[:,f"S{i}-LUNES":f"S{i}-DOMINGO"]

    COLTOT=CFIJAS+list(COLVARIABLESS.columns)

    dfRUTOT=dfrut[COLTOT]

    dfRUTOT=pd.melt(dfRUTOT,CFIJAS,value_name=f"S{i}-VISITAS",var_name="DIA")
    
    dfRUTOT["DIA"]=dfRUTOT["DIA"].str.split("-").str[1]

    dfRUTOT[f"S{i}-VISITAS"] = pd.to_numeric(dfRUTOT[f"S{i}-VISITAS"], errors='coerce').fillna(0)

    dfRUTOT=dfRUTOT[dfRUTOT[f"S{i}-VISITAS"]>0]

    dfRUTOT=dfRUTOT[["Usuario Virtual","Nombre de Tienda",f"S{i}-VISITAS","DIA"]]
    dfRUTOT=dfRUTOT.rename(columns={
        "Usuario Virtual":"Promotor",
        "Nombre de Tienda":"Tienda",
        f"S{i}-VISITAS":"un",
        "DIA":"Dia"
    })

    dfRUTOT["Llave"]=dfRUTOT["Promotor"].astype(str)+dfRUTOT["Dia"].astype(str).str.capitalize()

    dfRUTOT["Orden"] = dfRUTOT.groupby("Llave").cumcount() + 1

    dfRUTOT=dfRUTOT.drop(["un","Llave"],axis=1)

    dfRUTOT["Dia"]=dfRUTOT["Dia"].astype(str).str.upper()

    dfRUTOT=dfRUTOT[["Promotor","Tienda","Orden","Dia"]]

    # AGREGAR "Apoyo en Tienda" y "Expos" COMO PRIMERAS VISITAS
    promotores = dfRUTOT["Promotor"].unique()
    dias = dfRUTOT["Dia"].unique()
    
    nuevas_filas = []
    for promotor in promotores:
        for dia in dias:
            nuevas_filas.append({"Promotor": promotor, "Tienda": "Apoyo en Tienda", "Orden": 1, "Dia": dia})
            nuevas_filas.append({"Promotor": promotor, "Tienda": "Expos", "Orden": 2, "Dia": dia})
    
    df_nuevas = pd.DataFrame(nuevas_filas)
    
    # Ajustar el orden de las visitas existentes
    dfRUTOT["Orden"] = dfRUTOT["Orden"] + 2
    
    # Combinar los dataframes
    dfRUTOT = pd.concat([df_nuevas, dfRUTOT], ignore_index=True)
    
    # Ordenar por Promotor, Dia y Orden
    dfRUTOT = dfRUTOT.sort_values(by=["Promotor", "Dia", "Orden"]).reset_index(drop=True)

    #Carpeta donde se reciben los nuevos archivos
    NOMBREFINAL=f"SEMANAS Nov\\S{i}.xlsx"
    with pd.ExcelWriter(RUTA + NOMBREFINAL, engine='openpyxl') as writer:
        dfRUTOT.to_excel(writer, sheet_name=f'S{i}', index=False)

    print(f"TIENDAS S{i} CREADO")
    print(f"S{i} cuenta con {len(dfRUTOT)} visitas")