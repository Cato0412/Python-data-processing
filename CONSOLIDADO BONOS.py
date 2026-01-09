


#CREACION DE CONSOLIDADO GENERAL DE TODOS LOS KPI'S


#AOCMODO BASE DE VENTAS
#--------------------------------------------------------------------------------------------
import pandas as pd
import numpy as np

RUTA="C:\\Users\\lapmxdf558\\Documents\\JUAN\\BONOS PY\\"
NOMBREARCHIVO="ventas_plantilla.xlsx"
NOMBREFINAL="ventas_limpio.xlsx"

QUITARACT=["%","Last"]
DELIMACT="|".join(QUITARACT)
QUITARLST=["%","Act"]
DELIMLST="|".join(QUITARLST)


#ARCHIVO DE ORIGEN DE VENTAS 
dforigen=pd.read_excel(RUTA+NOMBREARCHIVO)
# Se localizan las columnas que no tengan ("Suma de %") el signo ~ hace la negacion del enunciado
df1=dforigen.loc[:,~dforigen.columns.str.contains(DELIMACT,case=False,na=False)]
df2=dforigen.loc[:,~dforigen.columns.str.contains(DELIMLST,case=False,na=False)]

df1.columns = df1.columns.str.replace('Suma de Act ', '', regex=False)
df2.columns = df1.columns.str.replace('Suma de Last ', '', regex=False)

#ANULAR LA DINAMIZCION DE COLUMNAS
df1=pd.melt(df1,id_vars=["ID TIENDA","TIENDA"],var_name="MES",value_name="VENTAS 2025")
df2=pd.melt(df2,id_vars=["ID TIENDA","TIENDA"],var_name="MES",value_name="VENTAS 2024")


df1["LLAVE"]=df1["MES"]+"-"+df1["ID TIENDA"]
df2["LLAVE"]=df2["MES"]+"-"+df2["ID TIENDA"]

#df1 VENTAS ACTUALES
df1=df1[["ID TIENDA","TIENDA","LLAVE","MES","VENTAS 2025"]]
#df2 VENTAS PASADAS
df2=df2[["ID TIENDA","TIENDA","LLAVE","MES","VENTAS 2024"]]

#DATAFRAME DEL CONSOLIDADO DE VENTAS
dfunion=pd.merge(df1,df2,on='LLAVE',how="inner")
dfunion=dfunion.drop(["ID TIENDA_y","TIENDA_y","MES_y"],axis=1)
dfunion=dfunion.rename(columns={"ID TIENDA_x":"ID_TIENDA","TIENDA_x":"TIENDA","MES_x":"MES"})


#print(list(dfunion.columns))
print("-----------------------------------------------------------------------")
print("SE CONTRUYO BIEN LA BASE DE VENTAS SIN CRUZAR")

#DFUNION ES LA BASE AOCMODADA SIN CRUZARSE AUN CON NINGUNA OTRA BASE

#INVOCAR RUTERO
import pandas as pd
import os
#----------------------------------------------------------------------------------------------------------------------------------------------------------------'
ruta = "C:\\Users\\lapmxdf558\\Documents\\JUAN\\BONOS PY\\"  #<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
#----------------------------------------------------------------------------------------------------------------------------------------------------------------'

# BUSCAR ARCHIVOS QUE CONTENGAN 'Rutero'
archivos = [f for f in os.listdir(ruta) if 'Rutero' in f]
if archivos:
    archivo_rutero = os.path.join(ruta, archivos[0])
    dfrut = pd.read_excel(archivo_rutero, sheet_name='RUTERO',skiprows=4)
    palabra=archivo_rutero.replace(ruta,"")
#SE CARGO EL RUTERO EN LA CARPETA
    print("Archivo cargado:", palabra)
else:
    print("No se encontró ningún archivo con 'Rutero' en el nombre.")


#CRUCE DE BASE DE VENTAS CON BASE DE RUTERO 
dfrutC=dfrut[["ID_TIENDA","Nombre Promotor","Usuario Virtual","Nombre Supervisor","Usuario Virtual Supervisor"]]
dfunion=pd.merge(dfunion,dfrutC,on='ID_TIENDA',how="inner")
dfunion=dfunion[["ID_TIENDA","TIENDA","LLAVE","Nombre Promotor","Usuario Virtual","Nombre Supervisor","Usuario Virtual Supervisor","MES","VENTAS 2025","VENTAS 2024"]]


#CREACION DE LIBRO DE EXCEL CON LA BASE DE VENTAS CRUZADA CON RUTERO 
with pd.ExcelWriter(RUTA + NOMBREFINAL, engine='openpyxl') as writer:
    dforigen.to_excel(writer, sheet_name='Original', index=False)
    df1.to_excel(writer, sheet_name='Limpio_ACT', index=False)
    df2.to_excel(writer, sheet_name='Limpio_LST', index=False)
    dfunion.to_excel(writer, sheet_name='UNIDO', index=False)


DIF=(len(df1)-len(dfunion))/12
#ARCHIVO FINAL DE VENTAS RENOMBRADO
VENTASFINAL=dfunion #<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
print(f"NO UBO COINCIDENCIAS DE {DIF} TIENDAS")
print(f"SE CRUZO CORRECTAMENTE LA BASE DE VENTA CON BASE DE RUTERO")
#print(list(dfunion.columns))



#BUSCAR BASE DE PERSONAL 
ruta = "C:\\Users\\lapmxdf558\\Documents\\JUAN\\BONOS PY\\"
#BUSCAR ARCHIVOS QUE CONTENGAN 'Personal'
archivos = [f for f in os.listdir(ruta) if 'Personal' in f]
if archivos:
    archivo_personal = os.path.join(ruta, archivos[0])
    dfper = pd.read_excel(archivo_personal, sheet_name='PERSONAL')
    palabra=archivo_personal.replace(ruta,"")
#SE CARGO EL PERSONAL EN LA CARPETA
    print("Archivo cargado:", palabra)
else:
    print("No se encontró ningún archivo con 'Personal' en el nombre.")



#RUTAS UNICAS 
rutasU=dfrut.drop_duplicates(subset=["Usuario Virtual"])
rutasU=rutasU["Usuario Virtual"]
nrutas=len(rutasU)

#FILTRO PARA ELEGIR MES DE TRABAJO
#-------------------------------------------------
MES="Jun" # <<<<<<<<<<<<<<<<<<<<<<<<MES DE TRABAJO
#-------------------------------------------------

MESFILTRADO=VENTASFINAL[VENTASFINAL["MES"]==MES]
#VENTAS POR RUTA
tabla=pd.pivot_table(MESFILTRADO,index="Usuario Virtual",values=["VENTAS 2025","VENTAS 2024"],aggfunc='sum')
tabla["IXx"]=tabla["VENTAS 2025"]/tabla["VENTAS 2024"]

#-------------------------------------------------------------------------------------------------------------------------------------------------------------
#OPCION DE TABLA DINAMICA POR MES DESPLEGADO
#tabla=pd.pivot_table(VENTASFINAL,index="Usuario Virtual",columns="MES",values=["VENTAS 2025","VENTAS 2024"],aggfunc='sum')
#with pd.ExcelWriter(RUTA + "tab.xlsx", engine='openpyxl') as writer:
#    tabla.reset_index().to_excel(writer, sheet_name='Original', index=False)
#-------------------------------------------------------------------------------------------------------------------------------------------------------------

#CALCULOS GENERALES VENTAS
VENTAS_2025=tabla["VENTAS 2025"].sum()
VENTAS_2024=tabla["VENTAS 2024"].sum()
IX=VENTAS_2025/VENTAS_2024
print(f"En 2025 se tuvieron ventas de ${VENTAS_2025:,.2f}")
print(f"En 2024 se tuvieron ventas de ${VENTAS_2024:,.2f}")
print(f"Con un IX  de {IX:,.2%}")

#VENTAS POR RUTA 
df_RxV=pd.merge(rutasU,tabla,on="Usuario Virtual",how="outer")
VENTASRUT=df_RxV #<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

#TABLA DE FRECUENCIA MENSUAL POR EMPLEADO
dffrec=pd.pivot_table(data=dfrut,index="Usuario Virtual",values="Numero de Visitas (Clasificacion)",aggfunc='sum')
dffrec["Frecuencia Mensual"]=dffrec["Numero de Visitas (Clasificacion)"]*4
dffrec.drop("Numero de Visitas (Clasificacion)",axis=1,inplace=True)
FRECMEN=dffrec #<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

#TABLA DE RUTAS CON CRUCE EN VENTAS Y FRECUENCIA MENSUAL 
dfU_VxFr=pd.merge(VENTASRUT,FRECMEN,on="Usuario Virtual",how="outer")

#BUSCAR ASISTENCIA MENSUAL 
archivos = [f for f in os.listdir(ruta) if 'ASISTENCIA' in f]
if archivos: 
    archivo_ASIS = os.path.join(ruta, archivos[0])
    dfasis = pd.read_excel(archivo_ASIS, sheet_name='Efectividad')
    palabra=archivo_ASIS.replace(ruta,"")
#SE CARGO EL PERSONAL EN LA CARPETA
    print("Archivo cargado:", palabra)
else:
    print("No se encontró ningún archivo con 'Personal' en el nombre.")

#RUTAS CON USUARIOS APP
dfUApp=dfrut.drop_duplicates(subset="Usuario Virtual")[["Usuario Virtual","Usuario APP Promotor"]]
#ASISTENCIA POR RUTA Y USUARIO APP 
df_ASISxR=pd.merge(dfasis,dfUApp,left_on="Usuario Promotor",right_on="Usuario APP Promotor",how="outer")
df_ASISxR["Tiempo en PDV"] = df_ASISxR["Tiempo en PDV"].fillna("00:00:00").replace("", "00:00:00")
df_ASISxR["Tiempo en PDV"] = pd.to_timedelta(df_ASISxR["Tiempo en PDV"])
df_ASISxR["Minutos Trabajados"] = df_ASISxR["Tiempo en PDV"].dt.total_seconds() / 60

#MINUTOS MENSUALES
ASISTENCIA = pd.pivot_table(df_ASISxR, index="Usuario Virtual", values="Minutos Trabajados", aggfunc="sum").reset_index()
#VISITAS REALIZADAS POR RUTA
dfASIS_R=pd.pivot_table(df_ASISxR,index="Usuario Virtual",values="Visitas Realizadas",aggfunc="sum").reset_index()
#VENTAS POR FRECUENCIA MENSUAL Y USUARIO APP POR RUTA
dfU_VxFrxUapp=pd.merge(dfU_VxFr,dfUApp,on="Usuario Virtual",how="outer")
dfU_VxFrxUapp["Minutos Objetivo"]=dfU_VxFrxUapp["Frecuencia Mensual"]*8*60

df_C_in_out=pd.merge(dfU_VxFrxUapp,ASISTENCIA,on="Usuario Virtual",how="outer")
#CONSOLIDADO DE RUTA POR VENTAS, MINUTOS OBJ, MINUTOS REALIZADOS, VISITAS REALIZADAS , VISITAS OBJETIVO(FRECUENCIA MENSUAL)
df_C_in_out=pd.merge(df_C_in_out,dfASIS_R,on="Usuario Virtual",how="outer")

#------------------------------------------------------------------------------------------------------------------------
DHab=25 # DIAS HABILES MENSUALES<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
#------------------------------------------------------------------------------------------------------------------------

df_C_in_out["Minutos mensuales"]=DHab*8*60
df_C_in_out["Dias habiles"]=DHab


#CARGAR FI (FOCOS DE IMPLEMENTACION)
archivos = [f for f in os.listdir(ruta) if 'TAREAS' in f]
if archivos:
    archivo_FI = os.path.join(ruta, archivos[0])
    dfFI = pd.read_excel(archivo_FI, sheet_name='EJECUCION_TAREAS')
    palabra=archivo_FI.replace(ruta,"")
#SE CARGO EL PERSONAL EN LA CARPETA
    print("Archivo cargado:", palabra)
else:
    print("No se encontró ningún archivo con 'Personal' en el nombre.")

#FILLNA PARA QUE LOS ESPCIOS VACIOS(NA) TOMEN EL VALOR INDICADO PARA POSTERIOR TRATAMIENTO
dfFI=pd.merge(dfFI,dfUApp,left_on="Usuario Promotor", right_on="Usuario APP Promotor",how="outer")
dfFI["TAREAS OBJETIVO"]=1
dfFI["Solucionado"]=dfFI["Solucionado"].fillna("NR")

#DOS FORMAS DE COLUMNAS CONDICIONALES
#PRIMERA MANERA
#DEFINIR FUNCION DONDE LA VARIABLE SERA LA COLUMNA DESEADA Y SE IMPLEMENTARA EN CADA FILA, TAMBIEN SE PUEDE USAR FUNCION LAMBDA
def implementado(valor):
    if valor=="Si":
        return 1
    elif valor=="No":
        return 0
    else:
        return 0

dfFI["TAREAS REALIZADAS"]=dfFI["Solucionado"].apply(implementado)


#SEGUNDA MANERA
#EL EUIVALENTE A UN IF EN EXCEL (SE PUEDE ANIDAR)
dfFI["TAREAS REALIZADASS"]=np.where(dfFI["Solucionado"]=="Si",1,np.where(dfFI["Solucionado"]=="No",0,np.where(dfFI["Solucionado"]=="NR",0,0)))
#FOCOS DE IMPLEMENTACION POR RUTA
dfFIxR=pd.pivot_table(dfFI,index="Usuario Virtual",values=["TAREAS REALIZADAS","TAREAS OBJETIVO"],aggfunc="sum").reset_index()

#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
#CONSOLIDADO GENERAL SIN NOMBRES DE HC, DEBIDO A QUE A VECES SE PIDEN REPORTES DE TIEMPOS ANTEIRORES (ES MAS VARIABLE), DEPENDERA MAS DEL HC QUE SE CARGUE EN LA CARPETA
TODO_S_HC=pd.merge(df_C_in_out,dfFIxR,on="Usuario Virtual",how="outer") #CONSOLIDADO S/NOMBRES <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
#------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#CARGAR HC
archivos = [f for f in os.listdir(ruta) if 'HC' in f]
if archivos:
    archivo_HC = os.path.join(ruta, archivos[0])
    dfHC = pd.read_excel(archivo_HC, sheet_name='HC ',skiprows=2)
    palabra=archivo_HC.replace(ruta,"")
#SE CARGO EL PERSONAL EN LA CARPETA
    print("Archivo cargado:", palabra)
else:
    print("No se encontró ningún archivo con 'Personal' en el nombre.")

#INSERTAR COLUMNA DE NOMBRE COMPLETO
dfHC["COMPLETO"]=dfHC["APELLIDO PATERNO"]+" "+dfHC["APELLIDO MATERNO"]+" "+dfHC["NOMBRE (S)"]
#INSERTER ANTES DE LA COLUMNA APELLIDO PATERNO
HCcol=list(dfHC.columns)


#SE USA EL FILLNA PORQUE LOS VACIOS LOS DETECTA PANDAS COMO NA
dfHC["NOMBRE COMPLETO"]=np.where(dfHC["COMPLETO"].fillna("")=="","VACANTE",dfHC["COMPLETO"])
#TABLA DE RUTA POR NOMBRE GENERAL (PROMOTORES Y SUPERVISORES)
df_RxNom=dfHC[["RUTA","NOMBRE COMPLETO","ESTADO","COORDINADOR","SUPERVISOR"]]

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
TODO_C_HC=pd.merge(TODO_S_HC,df_RxNom,left_on="Usuario Virtual",right_on="RUTA",how="left")#CONSOLIDADO C/NOMBRES <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
#------------------------------------------------------------------------------------------------------------------------------------------------------------------

archivo1="CONSOLIDADO SIN HC.xlsx"
archivo2="COPNSOLIDADO CON HC.xlsx"
with pd.ExcelWriter(RUTA + archivo1, engine='openpyxl') as writer:
 #   VENTASRUT.reset_index().to_excel(writer, sheet_name='O1', index=False)
    #  FRECMEN.reset_index().to_excel(writer, sheet_name='O2', index=False)
    #dfU_VxFr.to_excel(writer,sheet_name='O3',index=False)
    #dfUApp.to_excel(writer,sheet_name='O4',index=False)
    #df_ASISxR.to_excel(writer,sheet_name='O5',index=False)
    #dfASIS_R.to_excel(writer,sheet_name='O6',index=False)
    #ASISTENCIA.to_excel(writer,sheet_name='O7',index=False)
    #dfU_VxFrxUapp.to_excel(writer,sheet_name='O8',index=False)
    #df_C_in_out.to_excel(writer,sheet_name='O9',index=False)
    #dfFI.to_excel(writer,sheet_name='O10',index=False)
    #dfFIxR.to_excel(writer,sheet_name='O11',index=False)
    TODO_S_HC.to_excel(writer,sheet_name='O12',index=False)

print(f"Archivo creado: {archivo1}")

with pd.ExcelWriter(RUTA + archivo2, engine='openpyxl') as writer:
    #dfHC.to_excel(writer, sheet_name='O1', index=False)
    #df_RxNom.to_excel(writer, sheet_name='O2', index=False)
    TODO_C_HC.to_excel(writer, sheet_name='O3', index=False)



print(f"Archivo creado: {archivo2}")
print("-----------------------------------------------------------------------------------------")
