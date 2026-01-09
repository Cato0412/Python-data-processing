import pandas as pd
import glob
import os 
import numpy as np
import locale
from datetime import date
import warnings
warnings.simplefilter("ignore", UserWarning)



RUTAG="C:\\Users\\lapmxdf558\\Documents\\JUAN\\BONOS PY\\HBT\\"
EXRUT="RUTERO\\"
EXPER="PERSONAL\\"
EXVENTAS="VENTAS\\"
EXFI="FI\\"
EXEF="EFECTIVIDAD\\"

HOY = pd.Timestamp.today().normalize()

archivos = [f for f in os.listdir(RUTAG) if 'Personal' in f]
if archivos:
    archivo_PER = os.path.join(RUTAG, archivos[0])
    dfPER = pd.read_excel(archivo_PER, sheet_name='PERSONAL')
    palabra=archivo_PER.replace(RUTAG,"")
    print("Archivo cargado:", palabra)
else:
    print("No se encontró ningún archivo con 'Rutero' en el nombre.")

dfPER=dfPER[["Usuario APP","RUTA"]]


#APILAR RUTEROS POR MES
#-------------------------------------------------------------------------------------------------------------

RUTAR=RUTAG+EXRUT

archivosR=glob.glob(os.path.join(RUTAR,"Rutero*.xlsx"))
dfTODOSR=[]

for archivoR in archivosR:
    dfRUT=pd.read_excel(archivoR,sheet_name='RUTERO',skiprows=4)
    nombre=os.path.basename(archivoR)
    nombre=nombre.split(" ")
    mes=nombre[1]
    mes = mes.rsplit(".", 1)[0]
    dfRUT=dfRUT[['ID_TIENDA','TIENDA ID_CUBO','Nombre Promotor','Usuario Virtual','Usuario APP Promotor','Nombre Supervisor','Usuario Virtual Supervisor','Usuario App Supervisor','Usuario App Coordinador','Zona - Region','Nombre de Tienda','Latitud','Longitud','Area Nielsen','Canal de Distribución','Cadena','Formato','Numero de Visitas (Clasificacion)']]
    dfRUT=dfRUT.rename(columns={'Numero de Visitas (Clasificacion)':"FR"})
    dfRUT["HORAS TRBAJADAS"]=dfRUT["FR"]*4*8
    dfRUT["MINUTOS TRABAJADOS"]=dfRUT["HORAS TRBAJADAS"]*60
    dfRUT["MES"]=mes
    dfTODOSR.append(dfRUT)

df_finalR=pd.concat(dfTODOSR,ignore_index=True)

NOMBREFINALR="FFR.xlsx"

print("SE APILARON LOS RUTEROS POR MES CORRECTAMENTE")

#APLIAR PERSONALES POR MES
#-----------------------------------------------------------------------------------------------------------

RUTAP=RUTAG+EXPER

archivosP=glob.glob(os.path.join(RUTAP,"*Personal*.xlsx"))
dfTODOSP=[]

for archivoP in archivosP:
    dfPER=pd.read_excel(archivoP,sheet_name='PERSONAL')
    nombre=os.path.basename(archivoP)
    nombre=nombre.split(" ")[1]
    mes=nombre.rsplit(".",1)[0]
    dfPER=dfPER[['Tipo de Usuario','Usuario Agencia','Nombre Completo','Usuario Virtual','Usuario APP','Contraseña','RUTA','Nombre Completo','Latitud','Longitud','Fecha de ingreso','Supervisor Asignado OK','Coordinador Asignado']]
    dfPER=dfPER.rename(columns={"Fecha de ingreso":"FECHA NAC"})
    dfPER["FECHA NAC"] = pd.to_datetime(dfPER["FECHA NAC"], format="%d-%m-%Y", errors="coerce")
    dfPER["AÑOS"] = ((HOY - dfPER["FECHA NAC"]).dt.days / 365).round(2)
    dfPER["FECHA NAC"] = dfPER["FECHA NAC"].dt.strftime("%d/%m/%Y")
    dfPER["AÑOS"]=dfPER["AÑOS"].fillna(0)
    dfPER["MES"]=mes
    dfTODOSP.append(dfPER)

df_finalP=pd.concat(dfTODOSP,ignore_index=True)

NOMBREFINALP="FFP.xlsx"


print("SE APILARON LOS PERSONALES POR MES CORRECTAMENTE")

#APILAR EFECTIVIDAD POR MES
#-------------------------------------------------------------------------------------------------------------

RUTAEF=RUTAG+EXEF

archivosEF=glob.glob(os.path.join(RUTAEF,"Efectividad*.xlsx"))
dfTODOSEF=[]

for archivoEF in archivosEF:
    dfEF=pd.read_excel(archivoEF,sheet_name="Efectividad")
    nombre=os.path.basename(archivoEF)
    nombre=nombre.split(" ")[1]
    mes=nombre.split(".")[0]
    dfEF=pd.merge(dfEF,dfPER,left_on="Usuario Promotor",right_on="Usuario APP",how="left")
    dfEF=pd.merge(dfEF,dfPER,left_on="Usuario Supervisor",right_on="Usuario APP",how="left")
    dfEF=dfEF.rename(columns={"RUTA_x":"RUTA","RUTA_y":"ID SUP"})
    dfEF=dfEF[['Fecha','Primer Nivel Geográfico','Zona','Usuario Coordinador','Coordinador','Usuario Supervisor','ID SUP','Supervisor','Usuario Promotor','RUTA','Personal Promotor','Tienda','Cadena','Formato','Canal de Distribución','Tipo de Tienda','Check IN','Check OUT','Tiempo en PDV','Foto']]
    dfEF["Tiempo en PDV"]=dfEF["Tiempo en PDV"].fillna("")
    dfEF["Columna1"]=np.where(dfEF['Tiempo en PDV']=="",0,1)
    dfEF=dfEF.rename(columns={"Foto":"Columna2"})
    dfEF["MES"]=mes
    dfEF=dfEF[['Fecha','Primer Nivel Geográfico','Zona','Usuario Coordinador','Coordinador','Usuario Supervisor','ID SUP','Supervisor','Usuario Promotor','RUTA','Personal Promotor','Tienda','Cadena','Formato','Canal de Distribución','Tipo de Tienda','Check IN','Check OUT','Tiempo en PDV','Columna1','Columna2','MES']]
    dfTODOSEF.append(dfEF)

df_finalEF=pd.concat(dfTODOSEF,ignore_index=True)

NOMBREFINALEF="FFEF.xlsx"


print("SE APILARON LAS EFECTIVIDADES POR MES CORRECTAMENTE")


#APILAR FI POR MES
#------------------------------------------------------------------------------------------------------------------

RUTAFI=RUTAG+EXFI
archivosFI=glob.glob(os.path.join(RUTAFI,"FI*.xlsx"))
dfTODOSFI=[]

for archivoFI in archivosFI:
    dfFI=pd.read_excel(archivoFI, sheet_name="EJECUCION_TAREAS")
    nombre=os.path.basename(archivoFI)
    nombre=nombre.split(" ")[1]
    mes=nombre.split(".")[0]
    dfPER1=dfPER[["Usuario APP","RUTA"]]
    dfFI=pd.merge(dfFI,dfPER1,left_on="Usuario Promotor",right_on="Usuario APP",how="left")
    dfFI["MES"]=mes
    dfFI=dfFI.drop("Usuario APP",axis=1)
    dfTODOSFI.append(dfFI)

df_finalFI=pd.concat(dfTODOSFI)

NOMBREFINALFI="FFFI.xlsx"



print("SE APILARON LOS FOCOS DE IMPLEMENTACION POR MES CORRECTAMENTE")




#ACONDICIONAR VENTAS
#----------------------------------------------------------------------------------------------

RUTAV=RUTAG+EXVENTAS
archivo="ventas_plantilla.xlsx"

dfORIGEN=pd.read_excel(RUTAV+archivo)

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


#CAMBIO DE NOMBRE DE COLUMNA
dfCONSOLIDADO=dfCONSOLIDADO.rename(columns={"ID TIENDA":"ID_TIENDA"})


print("BASE DE DATOS DE VENTAS CARGADA CORRECTAMENTE")
print(f"PREVIO AL RUTERO")

#CARGAR RUTERO
archivos = [f for f in os.listdir(RUTAG) if 'Rutero' in f]
if archivos:
    archivo_rutero = os.path.join(RUTAG, archivos[0])
    dfrut = pd.read_excel(archivo_rutero, sheet_name='RUTERO',skiprows=4)
    palabra=archivo_rutero.replace(RUTAG,"")
    print("Archivo cargado:", palabra)
else:
    print("No se encontró ningún archivo con 'Rutero' en el nombre.")

#CRUZAR RUTERO POR TIENDA 
dfrut=dfrut[["ID_TIENDA","Usuario Virtual","Usuario APP Promotor","Area Nielsen","Estado","Canal de Distribución","Cadena","Formato","Nombre de Tienda"]]
dfVENTASxRUTERO=pd.merge(dfCONSOLIDADO,dfrut,on="ID_TIENDA",how="inner")

CONDICIONES=(
    dfVENTASxRUTERO["MES"]=="Ene",
    dfVENTASxRUTERO["MES"]=="Feb",
    dfVENTASxRUTERO["MES"]=="Mar",
    dfVENTASxRUTERO["MES"]=="Abr",
    dfVENTASxRUTERO["MES"]=="May",
    dfVENTASxRUTERO["MES"]=="Jun",
    dfVENTASxRUTERO["MES"]=="Jul",
    dfVENTASxRUTERO["MES"]=="Ago",
    dfVENTASxRUTERO["MES"]=="Sep",
    dfVENTASxRUTERO["MES"]=="Oct",
    dfVENTASxRUTERO["MES"]=="Nov",
    dfVENTASxRUTERO["MES"]=="Dic"
)

RESPUESTAS=(
    "Enero",
    "Febrero",
    "Marzo",
    "Abril",
    "Mayo",
    "Junio",
    "Julio",
    "Agosto",
    "Septiembre",
    "Octubre",
    "Noviembre",
    "Diciembre"
)


dfVENTASxRUTERO["MESB"]=np.select(CONDICIONES,RESPUESTAS,default="")

print("VENTAS ACOMODADAS")





with pd.ExcelWriter(os.path.join(RUTAG ,"ACUMULADO MESES.xlsx"), engine='openpyxl') as writer:
    df_finalFI.to_excel(writer, sheet_name='EJEC TAREAS', index=False)
    df_finalEF.to_excel(writer, sheet_name='EFECTIVIDAD', index=False)
    df_finalR.to_excel(writer, sheet_name='RUTERO', index=False)
    df_finalP.to_excel(writer, sheet_name='PER', index=False)
    dfVENTASxRUTERO.to_excel(writer, sheet_name='V ACT', index=False)

print("ARCHIVO CREADO 'ACUMULADO MESES'")