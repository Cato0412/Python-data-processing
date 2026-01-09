import pandas as pd
import numpy as np
import os as oss

RUTA=r"C:\\Users\\lapmxdf558\\Documents\\JUAN\\BONOS PY\\POS\\"
archivo="ventas_plantilla.xlsx"

dfORIGEN=pd.read_excel(RUTA+archivo)

QUITARACT=["%","Last"]
DELIMACT="|".join(QUITARACT)
QUITARLST=["%","Act"]
DELIMLST="|".join(QUITARLST)

dfACT=dfORIGEN.loc[:,~dfORIGEN.columns.str.contains(DELIMACT,case=False,na=False)]
dfLST=dfORIGEN.loc[:,~dfORIGEN.columns.str.contains(DELIMLST,case=False,na=False)]

dfACT.columns=dfACT.columns.str.replace("Suma de Act ","",regex=False)
dfLST.columns=dfLST.columns.str.replace("Suma de Last ","",regex=False)

dfACT=pd.melt(dfACT,id_vars=["ID TIENDA","TIENDA"],var_name="MES",value_name="VENTAS 2025")
dfLST=pd.melt(dfLST,id_vars=["ID TIENDA","TIENDA"],var_name="MES",value_name="VENTAS 2024")

dfACT["LLAVE"]=dfACT["MES"]+"-"+dfACT["ID TIENDA"]
dfLST["LLAVE"]=dfLST["MES"]+"-"+dfLST["ID TIENDA"]
dfLST_AC=dfLST[["LLAVE","VENTAS 2024"]]

dfCONSOLIDADO=pd.merge(dfACT,dfLST_AC,on="LLAVE",how="outer")
dfCONSOLIDADO=dfCONSOLIDADO.drop("LLAVE",axis=1)

#EN CASO DE QUERER FILTRAR DENTRO DEL ARCHIVO XLSX , COMENTAR ESTAS DOS LINEAS
#-------------------------------------------------------------------------------------------'
#MES_FILTRO="May" #MES DE FILTRADO <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
#dfCONSOLIDADO=dfCONSOLIDADO[dfCONSOLIDADO["MES"]==MES_FILTRO]
#-------------------------------------------------------------------------------------------'

VENTAS2025=dfCONSOLIDADO["VENTAS 2025"].sum()
VENTAS2024=dfCONSOLIDADO["VENTAS 2024"].sum()
IX=VENTAS2025/VENTAS2024

#CAMBIO DE NOMBRE DE COLUMNA
dfCONSOLIDADO=dfCONSOLIDADO.rename(columns={"ID TIENDA":"ID_TIENDA"})


print("BASE DE DATOS DE VENTAS CARGADA CORRECTAMENTE")
print(f"PREVIO AL RUTERO")
print(f"SE TUVIERON VENTAS DE ${VENTAS2025:,.2F} EN 2025")
print(f"SE TUVIERON VENTAS DE ${VENTAS2024:,.2F} EN 2024")
print(f"CON UN IX DE {IX:,.2%}")

if IX>1:
    print(f"UN {-(1-IX):,.2%} POR ENCIMA DEL AÑO PASADO")
else:
    print(f"UN {1-IX:,.2%} POR DEBAJO DEL AÑO PASADO")


#CARGAR RUTERO
archivos = [f for f in oss.listdir(RUTA) if 'Rutero' in f]
if archivos:
    archivo_rutero = oss.path.join(RUTA, archivos[0])
    dfrut = pd.read_excel(archivo_rutero, sheet_name='RUTERO',skiprows=4)
    palabra=archivo_rutero.replace(RUTA,"")
    print("Archivo cargado:", palabra)
else:
    print("No se encontró ningún archivo con 'Rutero' en el nombre.")

#CRUZAR RUTERO POR TIENDA 
dfrut=dfrut[["ID_TIENDA","Usuario Virtual","Usuario APP Promotor","Area Nielsen","Estado","Canal de Distribución","Cadena","Formato"]]
dfVENTASxRUTERO=pd.merge(dfCONSOLIDADO,dfrut,on="ID_TIENDA",how="inner")

TIENDAS_s_COINCIDENCIA=(len(dfCONSOLIDADO)-len(dfVENTASxRUTERO))/12

print(f"NO HUBO COINCIDENCIA CON {TIENDAS_s_COINCIDENCIA} TIENDAS ")

#CARGAR HC

archivos = [f for f in oss.listdir(RUTA) if 'HC' in f]
if archivos:
    archivo_HC = oss.path.join(RUTA, archivos[0])
    dfHC = pd.read_excel(archivo_HC, sheet_name='HC ',skiprows=2)
    palabra=archivo_HC.replace(RUTA,"")
    print("Archivo cargado:", palabra)
else:
    print("No se encontró ningún archivo con 'Rutero' en el nombre.")

dfHC["NCOMPLETO"]=dfHC["APELLIDO PATERNO"]+" "+dfHC["APELLIDO MATERNO"]+" "+dfHC["NOMBRE (S)"]
dfHC["NCOMPLETO"]=dfHC["NCOMPLETO"].fillna("")
dfHC["NOMBRE COMPLETO"]=np.where(dfHC["NCOMPLETO"]=="","VACANTE",dfHC["NCOMPLETO"])
dfHC_CORTO=dfHC[["RUTA","COORDINADOR","SUPERVISOR","NOMBRE COMPLETO"]]
dfHC_CORTO=dfHC_CORTO.drop_duplicates(subset="RUTA")

dfVENTASCOMPLETO=pd.merge(dfVENTASxRUTERO,dfHC_CORTO,left_on="Usuario Virtual",right_on="RUTA",how="left")
dfVENTASCOMPLETO=dfVENTASCOMPLETO.drop(columns=["RUTA"],axis=1)
dfVENTASCOMPLETO["STATUS"]=np.where(dfVENTASCOMPLETO["NOMBRE COMPLETO"]=="VACANTE","INACTIVO","ACTIVO")
dfVENTASxRUTA=pd.pivot_table(dfVENTASCOMPLETO,index="Usuario Virtual",values=["VENTAS 2025","VENTAS 2024"]).reset_index()


NOMBREFINAL="FINAL VENTAS.xlsx"
with pd.ExcelWriter(RUTA + NOMBREFINAL, engine='openpyxl') as writer:
    dfCONSOLIDADO.to_excel(writer, sheet_name='CONSOLIDADO', index=False)
    dfVENTASxRUTERO.to_excel(writer, sheet_name='CONSOLIDADOxRUTERO', index=False)
    dfVENTASCOMPLETO.to_excel(writer, sheet_name='COMPLETO', index=False)
    dfVENTASxRUTA.to_excel(writer, sheet_name='VENTASxRUTA', index=False)


print(f"ARCHIVO CREADO: {NOMBREFINAL}")



print("EJECUTANDO MACRO") 


import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True  

try:
    libro = excel.Workbooks.Open(r"C:\Users\lapmxdf558\Documents\JUAN\BONOS PY\POS\POS_BIC.xlsm")
    excel.Application.Run("'POS_BIC.xlsm'!ABRIRyCOPIAR")
except Exception as e:
    print("La macro cerró el libro, conexión finalizada:", e)

finally:
    try:
        excel.Quit()
    except:
        print("Excel ya estaba cerrado.")

print("MACRO EJECUTADA")