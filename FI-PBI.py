import pandas as pd
import glob
import os 
import numpy as np
import warnings
warnings.simplefilter("ignore", UserWarning)


RUTA="C:\\Users\\lapmxdf558\\Documents\\JUAN\\REPORTES BIC\\BI\\ORIGENES\\FI\\INDIVIDUALES\\"
ARCHIVOS = glob.glob(os.path.join(RUTA, "TS*.xlsx"))
dfTODOS = []  # lista para guardar cada bloque

#CARGANDO RUTERO
archivos = [f for f in os.listdir(RUTA) if 'Rutero' in f]
if archivos:
    archivo_RUT = os.path.join(RUTA, archivos[0])
    dfRUT = pd.read_excel(archivo_RUT, sheet_name='RUTERO',skiprows=4)
    palabra=archivo_RUT.replace(RUTA,"")
    print("Archivo cargado:", palabra)
else:
    print("No se encontró ningún archivo con 'Rutero' en el nombre.")

dfRUTERO1=dfRUT[["Usuario APP Promotor","Usuario Virtual"]].drop_duplicates(subset="Usuario APP Promotor")
dfRUTERO2=dfRUT[["Usuario APP Promotor","Nombre Supervisor"]].drop_duplicates(subset="Usuario APP Promotor")



#CARGANDO ARCHIVOS
for ARCHIVO in ARCHIVOS:
    # Leer el archivo
    df = pd.read_excel(ARCHIVO)
    nombre = os.path.basename(ARCHIVO)  # "TS35.xlsx"
    semana = ''.join(filter(str.isdigit, nombre))  # "35"
    
    # Agregar la columna de semana
    df["SEMANA"] = semana  
    df["OBJETIVO"] = 1
    df["Solucionado"]=df["Solucionado"].fillna("NR")

    CONDICIONES=[
        df["Solucionado"]=="Si",
        df["Solucionado"]=="No",
        df["Solucionado"]=="NR"
        ]

    RESULTADOS=[
        1,
        0,
        0
    ]

    df["REALIZADO"]=np.select(CONDICIONES,RESULTADOS,default=0)
    dfTODOS.append(df)

df_final = pd.concat(dfTODOS, ignore_index=True)
df_final=pd.merge(df_final,dfRUTERO1,left_on="Usuario Promotor",right_on="Usuario APP Promotor",how="left")
df_final["LLAVE"]=df_final["Usuario Virtual"].astype(str)+"-"+df_final["Código Tienda"].astype(str)
df_final=pd.merge(df_final,dfRUTERO2,left_on="Usuario Promotor",right_on="Usuario APP Promotor",how="left")
df_final=df_final.drop(columns={"Usuario APP Promotor_y","Usuario APP Promotor_x"})
df_final=df_final.rename(columns={"Usuario Virtual":"RUTA","Nombre Supervisor":"SUP"})

NOMBREFINAL="RESULTADO\\JUNTADO.xlsx"
with pd.ExcelWriter(RUTA + NOMBREFINAL, engine='openpyxl') as writer:
    df_final.to_excel(writer, sheet_name='EJECUCION_TAREAS', index=False)

