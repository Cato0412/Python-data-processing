import pandas as pd
import glob
import os 
import numpy as np
from datetime import datetime
import warnings
from pathlib import Path
warnings.simplefilter("ignore", UserWarning)


class ProcesadorDatos:
    """Clase para procesar y consolidar datos de ruteros, personal, efectividad y ventas"""
    
    def __init__(self, ruta_base=None):
        """
        Inicializa el procesador con la ruta base
        
        Args:
            ruta_base: Ruta base del directorio. Si es None, usa el directorio actual
        """
        if ruta_base is None:
            self.ruta_base = Path.cwd()
        else:
            self.ruta_base = Path(ruta_base)
        
        # Subdirectorios
        self.rutas = {
            'rutero': self.ruta_base / "RUTERO",
            'personal': self.ruta_base / "PERSONAL",
            'ventas': self.ruta_base / "VENTAS",
            'fi': self.ruta_base / "FI",
            'efectividad': self.ruta_base / "EFECTIVIDAD"
        }
        
        self.hoy = pd.Timestamp.today().normalize()
        self.df_personal_original = None
        
    def extraer_mes_archivo(self, ruta_archivo):
        """
        Extrae el mes del nombre del archivo
        
        Args:
            ruta_archivo: Ruta completa del archivo
            
        Returns:
            Mes extraído del nombre del archivo
        """
        try:
            nombre = os.path.basename(ruta_archivo)
            partes = nombre.split(" ")
            if len(partes) > 1:
                mes = partes[1].rsplit(".", 1)[0]
                return mes
            return None
        except Exception as e:
            print(f"⚠️  Error extrayendo mes de {ruta_archivo}: {e}")
            return None
    
    def buscar_archivo(self, patron, directorio=None):
        """
        Busca el primer archivo que coincida con el patrón
        
        Args:
            patron: Patrón a buscar en el nombre del archivo
            directorio: Directorio donde buscar. Si es None, usa ruta_base
            
        Returns:
            Ruta del archivo encontrado o None
        """
        if directorio is None:
            directorio = self.ruta_base
        
        archivos = [f for f in os.listdir(directorio) if patron in f]
        
        if archivos:
            ruta_archivo = os.path.join(directorio, archivos[0])
            print(f"✓ Archivo cargado: {archivos[0]}")
            return ruta_archivo
        else:
            print(f"⚠️  No se encontró archivo con '{patron}' en {directorio}")
            return None
    
    def cargar_personal_original(self):
        """Carga el archivo de personal más reciente para usar en merges"""
        try:
            archivo_personal = self.buscar_archivo('Personal', self.ruta_base)
            if archivo_personal:
                df = pd.read_excel(archivo_personal, sheet_name='PERSONAL')
                self.df_personal_original = df[["Usuario APP", "RUTA"]].copy()
                print(f"✓ Personal original cargado: {len(self.df_personal_original)} registros")
                return True
            return False
        except Exception as e:
            print(f"❌ Error cargando personal original: {e}")
            return False
    
    def apilar_ruteros(self):
        """Apila todos los archivos de rutero por mes"""
        print("\n📁 Procesando RUTEROS...")
        
        try:
            archivos = glob.glob(str(self.rutas['rutero'] / "Rutero*.xlsx"))
            
            if not archivos:
                print("⚠️  No se encontraron archivos de rutero")
                return None
            
            lista_dfs = []
            
            for archivo in archivos:
                try:
                    df = pd.read_excel(archivo, sheet_name='RUTERO', skiprows=4)
                    mes = self.extraer_mes_archivo(archivo)
                    
                    if mes is None:
                        continue
                    
                    # Seleccionar columnas
                    columnas_necesarias = [
                        'ID_TIENDA', 'TIENDA ID_CUBO', 'Nombre Promotor', 
                        'Usuario Virtual', 'Usuario APP Promotor', 'Nombre Supervisor',
                        'Usuario Virtual Supervisor', 'Usuario App Supervisor',
                        'Usuario App Coordinador', 'Zona - Region', 'Nombre de Tienda',
                        'Latitud', 'Longitud', 'Area Nielsen', 'Canal de Distribución',
                        'Cadena', 'Formato', 'Numero de Visitas (Clasificacion)'
                    ]
                    
                    df = df[columnas_necesarias].copy()
                    df.rename(columns={'Numero de Visitas (Clasificacion)': "FR"}, inplace=True)
                    
                    # Cálculos
                    df["HORAS TRABAJADAS"] = df["FR"] * 4 * 8
                    df["MINUTOS TRABAJADOS"] = df["HORAS TRABAJADAS"] * 60
                    df["MES"] = mes
                    
                    lista_dfs.append(df)
                    
                except Exception as e:
                    print(f"⚠️  Error procesando {os.path.basename(archivo)}: {e}")
                    continue
            
            if lista_dfs:
                df_final = pd.concat(lista_dfs, ignore_index=True)
                print(f"✓ Ruteros apilados: {len(df_final)} registros de {len(lista_dfs)} archivos")
                return df_final
            else:
                print("❌ No se pudieron procesar ruteros")
                return None
                
        except Exception as e:
            print(f"❌ Error en apilar_ruteros: {e}")
            return None
    
    def apilar_personal(self):
        """Apila todos los archivos de personal por mes"""
        print("\n👥 Procesando PERSONAL...")
        
        try:
            archivos = glob.glob(str(self.rutas['personal'] / "*Personal*.xlsx"))
            
            if not archivos:
                print("⚠️  No se encontraron archivos de personal")
                return None
            
            lista_dfs = []
            
            for archivo in archivos:
                try:
                    df = pd.read_excel(archivo, sheet_name='PERSONAL')
                    mes = self.extraer_mes_archivo(archivo)
                    
                    if mes is None:
                        continue
                    
                    # Seleccionar columnas
                    columnas_necesarias = [
                        'Tipo de Usuario', 'Usuario Agencia', 'Nombre Completo',
                        'Usuario Virtual', 'Usuario APP', 'Contraseña', 'RUTA',
                        'Latitud', 'Longitud', 'Fecha de ingreso',
                        'Supervisor Asignado OK', 'Coordinador Asignado'
                    ]
                    
                    df = df[columnas_necesarias].copy()
                    df.rename(columns={"Fecha de ingreso": "FECHA NAC"}, inplace=True)
                    
                    # Procesar fechas
                    df["FECHA NAC"] = pd.to_datetime(df["FECHA NAC"], format="%d-%m-%Y", errors="coerce")
                    df["AÑOS"] = ((self.hoy - df["FECHA NAC"]).dt.days / 365).round(2)
                    df["FECHA NAC"] = df["FECHA NAC"].dt.strftime("%d/%m/%Y")
                    df["AÑOS"] = df["AÑOS"].fillna(0)
                    df["MES"] = mes
                    
                    lista_dfs.append(df)
                    
                except Exception as e:
                    print(f"⚠️  Error procesando {os.path.basename(archivo)}: {e}")
                    continue
            
            if lista_dfs:
                df_final = pd.concat(lista_dfs, ignore_index=True)
                print(f"✓ Personal apilado: {len(df_final)} registros de {len(lista_dfs)} archivos")
                return df_final
            else:
                print("❌ No se pudieron procesar archivos de personal")
                return None
                
        except Exception as e:
            print(f"❌ Error en apilar_personal: {e}")
            return None
    
    def apilar_efectividad(self):
        """Apila todos los archivos de efectividad por mes"""
        print("\n📊 Procesando EFECTIVIDAD...")
        
        if self.df_personal_original is None:
            print("❌ Se necesita cargar personal original primero")
            return None
        
        try:
            archivos = glob.glob(str(self.rutas['efectividad'] / "Efectividad*.xlsx"))
            
            if not archivos:
                print("⚠️  No se encontraron archivos de efectividad")
                return None
            
            lista_dfs = []
            
            for archivo in archivos:
                try:
                    df = pd.read_excel(archivo, sheet_name="Efectividad")
                    mes = self.extraer_mes_archivo(archivo)
                    
                    if mes is None:
                        continue
                    
                    # Merge con personal (promotor)
                    df = pd.merge(
                        df, 
                        self.df_personal_original.rename(columns={"RUTA": "RUTA_PROMOTOR"}),
                        left_on="Usuario Promotor", 
                        right_on="Usuario APP", 
                        how="left"
                    )
                    df.drop("Usuario APP", axis=1, inplace=True)
                    
                    # Merge con personal (supervisor)
                    df = pd.merge(
                        df,
                        self.df_personal_original.rename(columns={"RUTA": "ID_SUP", "Usuario APP": "Usuario APP_SUP"}),
                        left_on="Usuario Supervisor",
                        right_on="Usuario APP_SUP",
                        how="left"
                    )
                    df.drop("Usuario APP_SUP", axis=1, inplace=True)
                    
                    # Renombrar y seleccionar columnas
                    df.rename(columns={"RUTA_PROMOTOR": "RUTA"}, inplace=True)
                    
                    columnas_finales = [
                        'Fecha', 'Primer Nivel Geográfico', 'Zona', 'Usuario Coordinador',
                        'Coordinador', 'Usuario Supervisor', 'ID_SUP', 'Supervisor',
                        'Usuario Promotor', 'RUTA', 'Personal Promotor', 'Tienda',
                        'Cadena', 'Formato', 'Canal de Distribución', 'Tipo de Tienda',
                        'Check IN', 'Check OUT', 'Tiempo en PDV', 'Foto'
                    ]
                    
                    # Verificar que las columnas existan
                    columnas_existentes = [col for col in columnas_finales if col in df.columns]
                    df = df[columnas_existentes].copy()
                    
                    # Procesar datos
                    df["Tiempo en PDV"] = df["Tiempo en PDV"].fillna("")
                    df["Columna1"] = np.where(df['Tiempo en PDV'] == "", 0, 1)
                    df.rename(columns={"Foto": "Columna2"}, inplace=True)
                    df["MES"] = mes
                    
                    lista_dfs.append(df)
                    
                except Exception as e:
                    print(f"⚠️  Error procesando {os.path.basename(archivo)}: {e}")
                    continue
            
            if lista_dfs:
                df_final = pd.concat(lista_dfs, ignore_index=True)
                print(f"✓ Efectividad apilada: {len(df_final)} registros de {len(lista_dfs)} archivos")
                return df_final
            else:
                print("❌ No se pudieron procesar archivos de efectividad")
                return None
                
        except Exception as e:
            print(f"❌ Error en apilar_efectividad: {e}")
            return None
    
    def apilar_fi(self):
        """Apila todos los archivos de Focos de Implementación por mes"""
        print("\n🎯 Procesando FOCOS DE IMPLEMENTACIÓN...")
        
        if self.df_personal_original is None:
            print("❌ Se necesita cargar personal original primero")
            return None
        
        try:
            archivos = glob.glob(str(self.rutas['fi'] / "FI*.xlsx"))
            
            if not archivos:
                print("⚠️  No se encontraron archivos FI")
                return None
            
            lista_dfs = []
            
            for archivo in archivos:
                try:
                    df = pd.read_excel(archivo, sheet_name="EJECUCION_TAREAS")
                    mes = self.extraer_mes_archivo(archivo)
                    
                    if mes is None:
                        continue
                    
                    # Merge con personal
                    df = pd.merge(
                        df,
                        self.df_personal_original,
                        left_on="Usuario Promotor",
                        right_on="Usuario APP",
                        how="left"
                    )
                    
                    df.drop("Usuario APP", axis=1, inplace=True)
                    df["MES"] = mes
                    
                    lista_dfs.append(df)
                    
                except Exception as e:
                    print(f"⚠️  Error procesando {os.path.basename(archivo)}: {e}")
                    continue
            
            if lista_dfs:
                df_final = pd.concat(lista_dfs, ignore_index=True)
                print(f"✓ FI apilados: {len(df_final)} registros de {len(lista_dfs)} archivos")
                return df_final
            else:
                print("❌ No se pudieron procesar archivos FI")
                return None
                
        except Exception as e:
            print(f"❌ Error en apilar_fi: {e}")
            return None
    
    def procesar_ventas(self):
        """Procesa y acondiciona los datos de ventas"""
        print("\n💰 Procesando VENTAS...")
        
        try:
            archivo_ventas = self.rutas['ventas'] / "ventas_plantilla.xlsx"
            
            if not archivo_ventas.exists():
                print(f"❌ No se encontró {archivo_ventas}")
                return None
            
            df_origen = pd.read_excel(archivo_ventas)
            
            # Filtrar columnas Act y Last
            columnas_act = [col for col in df_origen.columns 
                           if not any(x in col for x in ['%', 'Last'])]
            columnas_lst = [col for col in df_origen.columns 
                           if not any(x in col for x in ['%', 'Act'])]
            
            df_act = df_origen[columnas_act].copy()
            df_lst = df_origen[columnas_lst].copy()
            
            # Limpiar nombres de columnas
            df_act.columns = df_act.columns.str.replace("Suma de Act ", "", regex=False)
            df_lst.columns = df_lst.columns.str.replace("Suma de Last ", "", regex=False)
            
            # Melt (unpivot)
            df_act = pd.melt(
                df_act,
                id_vars=["ID TIENDA", "TIENDA"],
                var_name="MES",
                value_name="VENTAS 2025"
            )
            
            df_lst = pd.melt(
                df_lst,
                id_vars=["ID TIENDA", "TIENDA"],
                var_name="MES",
                value_name="VENTAS 2024"
            )
            
            # Crear llaves y merge
            df_act["LLAVE"] = df_act["MES"] + "-" + df_act["ID TIENDA"]
            df_lst["LLAVE"] = df_lst["MES"] + "-" + df_lst["ID TIENDA"]
            
            df_consolidado = pd.merge(
                df_act,
                df_lst[["LLAVE", "VENTAS 2024"]],
                on="LLAVE",
                how="outer"
            )
            
            df_consolidado.drop("LLAVE", axis=1, inplace=True)
            df_consolidado.rename(columns={"ID TIENDA": "ID_TIENDA"}, inplace=True)
            
            print(f"✓ Ventas procesadas: {len(df_consolidado)} registros")
            
            # Cargar rutero actual y cruzar
            archivo_rutero = self.buscar_archivo('Rutero', self.ruta_base)
            
            if archivo_rutero:
                df_rut = pd.read_excel(archivo_rutero, sheet_name='RUTERO', skiprows=4)
                
                columnas_rutero = [
                    "ID_TIENDA", "Usuario Virtual", "Usuario APP Promotor",
                    "Area Nielsen", "Estado", "Canal de Distribución",
                    "Cadena", "Formato", "Nombre de Tienda"
                ]
                
                df_rut = df_rut[columnas_rutero].copy()
                
                df_ventas_rutero = pd.merge(
                    df_consolidado,
                    df_rut,
                    on="ID_TIENDA",
                    how="inner"
                )
                
                # Mapear nombres de meses
                meses_map = {
                    "Ene": "Enero", "Feb": "Febrero", "Mar": "Marzo",
                    "Abr": "Abril", "May": "Mayo", "Jun": "Junio",
                    "Jul": "Julio", "Ago": "Agosto", "Sep": "Septiembre",
                    "Oct": "Octubre", "Nov": "Noviembre", "Dic": "Diciembre"
                }
                
                df_ventas_rutero["MESB"] = df_ventas_rutero["MES"].map(meses_map).fillna("")
                
                print(f"✓ Ventas cruzadas con rutero: {len(df_ventas_rutero)} registros")
                return df_ventas_rutero
            else:
                print("⚠️  No se encontró rutero para cruzar ventas")
                return df_consolidado
                
        except Exception as e:
            print(f"❌ Error en procesar_ventas: {e}")
            return None
    
    def generar_archivo_consolidado(self, nombre_salida="ACUMULADO MESES.xlsx"):
        """
        Ejecuta todo el proceso y genera el archivo consolidado
        
        Args:
            nombre_salida: Nombre del archivo de salida
        """
        print("=" * 60)
        print("🚀 INICIANDO PROCESAMIENTO DE DATOS")
        print("=" * 60)
        
        # Cargar personal original primero
        if not self.cargar_personal_original():
            print("❌ No se pudo cargar personal original. Abortando.")
            return False
        
        # Procesar cada tipo de archivo
        df_fi = self.apilar_fi()
        df_efectividad = self.apilar_efectividad()
        df_rutero = self.apilar_ruteros()
        df_personal = self.apilar_personal()
        df_ventas = self.procesar_ventas()
        
        # Verificar que al menos tengamos algunos datos
        dfs_disponibles = {
            'EJEC TAREAS': df_fi,
            'EFECTIVIDAD': df_efectividad,
            'RUTERO': df_rutero,
            'PER': df_personal,
            'V ACT': df_ventas
        }
        
        dfs_validos = {k: v for k, v in dfs_disponibles.items() if v is not None}
        
        if not dfs_validos:
            print("\n❌ No se generaron datos válidos. Abortando.")
            return False
        
        # Generar archivo Excel
        print("\n💾 Generando archivo consolidado...")
        
        try:
            ruta_salida = self.ruta_base / nombre_salida
            
            with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
                for nombre_hoja, df in dfs_validos.items():
                    df.to_excel(writer, sheet_name=nombre_hoja, index=False)
                    print(f"  ✓ Hoja '{nombre_hoja}' creada con {len(df)} registros")
            
            print("\n" + "=" * 60)
            print(f"✅ ARCHIVO CREADO EXITOSAMENTE: {nombre_salida}")
            print(f"📍 Ubicación: {ruta_salida}")
            print("=" * 60)
            
            return True
            
        except Exception as e:
            print(f"\n❌ Error al generar archivo: {e}")
            return False


# ============================================================================
# EJECUCIÓN PRINCIPAL
# ============================================================================

if __name__ == "__main__":
    # Opción 1: Usar ruta específica (como en tu código original)
    # RUTA_BASE = r"C:\Users\lapmxdf558\Documents\JUAN\BONOS PY\HBT"
    
    # Opción 2: Usar directorio actual (más portable)
    RUTA_BASE = None  # Cambia a None para usar el directorio actual
    
    # Crear procesador y ejecutar
    procesador = ProcesadorDatos(RUTA_BASE)
    procesador.generar_archivo_consolidado()