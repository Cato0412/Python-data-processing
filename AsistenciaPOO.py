import pandas as pd
import numpy as np
import os
from datetime import datetime
import win32com.client
from typing import Tuple, Dict, Optional


class ConfiguracionAsistencia:
    """Configuración centralizada del sistema"""
    RUTA_BASE = r"C:\Users\lapmxdf558\Documents\JUAN\BONOS PY\ASISTENCIA"
    RUTA_SUP = r"C:\Users\lapmxdf558\Documents\JUAN\BONOS PY\ASISTENCIA\SUP"
    RUTA_PLANTILLA = r"C:\Users\lapmxdf558\Documents\JUAN\BONOS PY\ASISTENCIA\PLANTILLA ASISTENCIA.xlsm"
    MACRO_NOMBRE = "ACONDICIONAR"
    HORAS_LABORALES = 8
    TIEMPO_TRASLADO = 1.5
    
    DIAS_VALIDOS = ["LUN", "MAR", "MIER", "JUE", "VIE", "SAB", "DOM"]
    SEMANAS_VALIDAS = range(1, 6)
    
    MAPEO_COLUMNAS_RUTERO = {
        "S1-LUNES": "LUN-S1", "S1-MARTES": "MAR-S1", "S1-MIERCOLES": "MIER-S1",
        "S1-JUEVES": "JUE-S1", "S1-VIERNES": "VIE-S1", "S1-SABADO": "SAB-S1", "S1-DOMINGO": "DOM-S1",
        "S2-LUNES": "LUN-S2", "S2-MARTES": "MAR-S2", "S2-MIERCOLES": "MIER-S2",
        "S2-JUEVES": "JUE-S2", "S2-VIERNES": "VIE-S2", "S2-SABADO": "SAB-S2", "S2-DOMINGO": "DOM-S2",
        "S3-LUNES": "LUN-S3", "S3-MARTES": "MAR-S3", "S3-MIERCOLES": "MIER-S3",
        "S3-JUEVES": "JUE-S3", "S3-VIERNES": "VIE-S3", "S3-SABADO": "SAB-S3", "S3-DOMINGO": "DOM-S3",
        "S4-LUNES": "LUN-S4", "S4-MARTES": "MAR-S4", "S4-MIERCOLES": "MIER-S4",
        "S4-JUEVES": "JUE-S4", "S4-VIERNES": "VIE-S4", "S4-SABADO": "SAB-S4", "S4-DOMINGO": "DOM-S4",
        "S5-LUNES": "LUN-S5", "S5-MARTES": "MAR-S5", "S5-MIERCOLES": "MIER-S5",
        "S5-JUEVES": "JUE-S5", "S5-VIERNES": "VIE-S5", "S5-SABADO": "SAB-S5", "S5-DOMINGO": "DOM-S5"
    }


class CargadorArchivos:
    """Maneja la carga de archivos Excel"""
    
    def __init__(self, ruta_base: str):
        self.ruta_base = ruta_base
    
    def cargar_efectividad(self) -> Optional[pd.DataFrame]:
        """Carga el archivo de Efectividad"""
        return self._cargar_archivo('Efectividad', 'Efectividad')
    
    def cargar_personal(self) -> Optional[pd.DataFrame]:
        """Carga el archivo de Personal"""
        df = self._cargar_archivo('Personal', 'PERSONAL')
        if df is not None:
            self._diagnosticar_columnas(df, "PERSONAL")
            df.columns = df.columns.str.strip()
        return df
    
    def cargar_rutero(self) -> Optional[pd.DataFrame]:
        """Carga el archivo de Rutero"""
        df = self._cargar_archivo('Rutero', 'RUTERO', skiprows=4)
        if df is not None:
            df.columns = df.columns.str.strip()
        return df
    
    def _cargar_archivo(self, keyword: str, sheet_name: str, skiprows: int = 0) -> Optional[pd.DataFrame]:
        """Método genérico para cargar archivos Excel"""
        try:
            archivos = [f for f in os.listdir(self.ruta_base) if keyword in f]
            if not archivos:
                print(f"⚠️  No se encontró ningún archivo con '{keyword}' en el nombre.")
                return None
            
            ruta_archivo = os.path.join(self.ruta_base, archivos[0])
            df = pd.read_excel(ruta_archivo, sheet_name=sheet_name, skiprows=skiprows)
            print(f"✓ Archivo cargado: {archivos[0]}")
            return df
        except Exception as e:
            print(f"❌ Error al cargar {keyword}: {e}")
            return None
    
    def _diagnosticar_columnas(self, df: pd.DataFrame, nombre: str):
        """Muestra diagnóstico de columnas"""
        print(f"\n{'='*60}")
        print(f"COLUMNAS DISPONIBLES EN ARCHIVO {nombre}:")
        print(f"{'='*60}")
        for i, col in enumerate(df.columns, 1):
            print(f"{i:2}. '{col}'")
        print(f"{'='*60}\n")


class ProcesadorEfectividad:
    """Procesa datos de efectividad y asistencia"""
    
    def __init__(self, config: ConfiguracionAsistencia):
        self.config = config
    
    def procesar_efectividad(self, df: pd.DataFrame) -> pd.DataFrame:
        """Procesa el dataframe de efectividad"""
        # Seleccionar columnas necesarias
        columnas = [
            "Fecha", "Usuario Coordinador", "Coordinador", "Usuario Supervisor", 
            "Supervisor", "Usuario Promotor", "Personal Promotor", "Código Tienda", 
            "Tienda", "Cadena", "Formato", "Canal de Distribución", "Check IN", 
            "Check OUT", "Visitas Programadas", "Visitas Realizadas", "Tiempo en PDV"
        ]
        df = df[columnas].copy()
        
        # Procesar horas
        df = self._procesar_tiempo_pdv(df)
        df = self._procesar_check_in_out(df)
        
        return df
    
    def _procesar_tiempo_pdv(self, df: pd.DataFrame) -> pd.DataFrame:
        """Convierte tiempo en PDV a formato decimal"""
        df = df.fillna("").replace("", "00:00:00")
        df["PDV THr"] = pd.to_timedelta(df["Tiempo en PDV"])
        df["PDV HNum"] = df["PDV THr"].dt.total_seconds() / 3600
        return df
    
    def _procesar_check_in_out(self, df: pd.DataFrame) -> pd.DataFrame:
        """Procesa horas de check in y check out"""
        check_in_dt = pd.to_datetime(df["Check IN"], errors="coerce", format="%d-%m-%Y - %H:%M:%S")
        check_out_dt = pd.to_datetime(df["Check OUT"], errors="coerce", format="%d-%m-%Y - %H:%M:%S")
        
        df["Check IN HORAS"] = (
            check_in_dt.dt.hour + 
            check_in_dt.dt.minute / 60 + 
            check_in_dt.dt.second / 3600
        ).fillna(0)
        
        df["Check OUT HORAS"] = (
            check_out_dt.dt.hour + 
            check_out_dt.dt.minute / 60 + 
            check_out_dt.dt.second / 3600
        ).fillna(0)
        
        return df
    
    def calcular_horas_extremas(self, df: pd.DataFrame) -> pd.DataFrame:
        """Calcula horas máximas y mínimas por promotor"""
        df_max = pd.pivot_table(
            df, 
            index="Usuario Promotor", 
            values="Check OUT HORAS", 
            aggfunc=self._horas_maximas
        ).reset_index()
        
        df_min = pd.pivot_table(
            df, 
            index="Usuario Promotor", 
            values="Check IN HORAS", 
            aggfunc=self._horas_minimas
        ).reset_index()
        
        df_checks = pd.merge(df_max, df_min, on="Usuario Promotor", how="inner")
        df_checks["DIF"] = df_checks["Check OUT HORAS"] - df_checks["Check IN HORAS"]
        
        return df_checks
    
    @staticmethod
    def _horas_maximas(valores):
        """Calcula hora máxima válida"""
        horas_validas = [h for h in valores if h > 0]
        return max(horas_validas) if horas_validas else 0
    
    @staticmethod
    def _horas_minimas(valores):
        """Calcula hora mínima válida"""
        horas_validas = [h for h in valores if h > 0]
        return min(horas_validas) if horas_validas else 0
    
    def consolidar_asistencia(self, df_efectividad: pd.DataFrame) -> pd.DataFrame:
        """Consolida información de asistencia"""
        df_checks = self.calcular_horas_extremas(df_efectividad)
        
        # Horas en PDV por usuario
        df_tpdv = pd.pivot_table(
            df_efectividad, 
            index="Usuario Promotor", 
            values="PDV HNum", 
            aggfunc="sum"
        ).reset_index()
        
        df_hrs = pd.merge(df_checks, df_tpdv, on="Usuario Promotor", how="inner")
        
        # Clasificar asistencia
        df_hrs = self._clasificar_asistencia(df_hrs)
        
        # Visitas realizadas
        df_realizadas = pd.pivot_table(
            df_efectividad, 
            index="Usuario Promotor", 
            values="Visitas Realizadas", 
            aggfunc="sum"
        ).reset_index()
        
        df_consolidado = pd.merge(df_hrs, df_realizadas, on="Usuario Promotor", how="inner")
        df_consolidado = self._clasificar_cumplimiento(df_consolidado)
        
        return df_consolidado
    
    def _clasificar_asistencia(self, df: pd.DataFrame) -> pd.DataFrame:
        """Clasifica la asistencia según check in/out"""
        condiciones = [
            (df["Check OUT HORAS"] == 0) & (df["Check IN HORAS"] != 0),
            (df["Check OUT HORAS"] != 0) & (df["Check IN HORAS"] == 0),
            (df["Check OUT HORAS"] == 0) & (df["Check IN HORAS"] == 0),
            (df["Check OUT HORAS"] != 0) & (df["Check IN HORAS"] != 0)
        ]
        
        resultados = [
            "NO HIZO CHECK OUT",
            "NO HIZO CHECK IN",
            "NO ASISTIO",
            "ASISTIO"
        ]
        
        df["ASISTENCIA"] = np.select(condiciones, resultados, default="")
        return df
    
    def _clasificar_cumplimiento(self, df: pd.DataFrame) -> pd.DataFrame:
        """Clasifica cumplimiento de horas"""
        condiciones = [
            df["DIF"] < self.config.HORAS_LABORALES,
            df["DIF"] >= self.config.HORAS_LABORALES
        ]
        
        resultados = [
            f"NO CUMPLIO CON SUS {self.config.HORAS_LABORALES} HORAS",
            f"CUMPLIO CON SUS {self.config.HORAS_LABORALES} HORAS"
        ]
        
        df["CUMPLIMIENTO"] = np.select(condiciones, resultados, default="")
        return df


class ProcesadorRutero:
    """Procesa datos del rutero"""
    
    def __init__(self, config: ConfiguracionAsistencia):
        self.config = config
    
    def procesar_rutero(self, df: pd.DataFrame) -> pd.DataFrame:
        """Procesa el dataframe de rutero"""
        df = df.rename(columns=self.config.MAPEO_COLUMNAS_RUTERO)
        
        columnas_fijas = ["ID_TIENDA", "Nombre de Tienda", "Usuario Virtual", "Usuario APP Promotor"]
        columnas_rango = df.loc[:, "LUN-S1":"DOM-S5"].columns.tolist()
        columnas_todas = columnas_fijas + columnas_rango
        
        df_filtrado = df[columnas_todas]
        df_filtrado = pd.melt(
            df_filtrado, 
            columnas_fijas, 
            value_name="Visitas Programadas", 
            var_name="DIA-SEMANA"
        )
        
        return df_filtrado
    
    def obtener_visitas_programadas(self, df_rutero: pd.DataFrame, dia: str, semana: int) -> pd.DataFrame:
        """Obtiene visitas programadas para un día y semana específicos"""
        codigo = f"{dia}-S{semana}"
        df_programadas = df_rutero[df_rutero["DIA-SEMANA"] == codigo]
        df_programadas = pd.pivot_table(
            df_programadas, 
            index="Usuario APP Promotor", 
            values="Visitas Programadas", 
            aggfunc="sum"
        ).reset_index()
        df_programadas = df_programadas.rename(columns={"Usuario APP Promotor": "Usuario Promotor"})
        
        return df_programadas


class ConstructorReporte:
    """Construye el reporte final de asistencia"""
    
    def __init__(self, config: ConfiguracionAsistencia):
        self.config = config
    
    def construir_reporte_principal(
        self, 
        df_asistencia: pd.DataFrame, 
        df_programadas: pd.DataFrame, 
        df_personal: pd.DataFrame
    ) -> pd.DataFrame:
        """Construye el reporte principal de asistencia"""
        # Unir asistencia con visitas programadas
        df_reporte = pd.merge(df_asistencia, df_programadas, on="Usuario Promotor", how="left")
        
        # Unir con información de personal
        df_personal_filtrado = df_personal[[
            "Usuario APP", "Usuario Virtual", "Nombre Completo", 
            "Supervisor Asignado OK", "Coordinador Asignado"
        ]]
        
        df_reporte = pd.merge(
            df_reporte, 
            df_personal_filtrado, 
            left_on="Usuario Promotor", 
            right_on="Usuario APP", 
            how="left"
        )
        
        # Filtrar solo asistencias
        df_reporte = df_reporte[df_reporte["ASISTENCIA"] != "NO ASISTIO"]
        
        # Calcular métricas
        df_reporte = self._calcular_metricas(df_reporte)
        
        # Seleccionar y ordenar columnas
        columnas_finales = [
            "Coordinador Asignado", "Supervisor Asignado OK", "Nombre Completo", 
            "Usuario Virtual", "Visitas Programadas", "Visitas Realizadas", 
            "DIFERENCIA VISITAS", "EFECTIVIDAD VISITAS", "DIF", "EFECTIVIDAD HORAS", 
            "Check IN HORAS", "Check OUT HORAS", "CUMPLIMIENTO", "ALCANCE HORAS", "ASISTENCIA"
        ]
        
        df_reporte = df_reporte[columnas_finales].round(2)
        
        return df_reporte
    
    def _calcular_metricas(self, df: pd.DataFrame) -> pd.DataFrame:
        """Calcula métricas de efectividad y alcance"""
        df["DIFERENCIA VISITAS"] = df["Visitas Programadas"] - df["Visitas Realizadas"]
        df["EFECTIVIDAD VISITAS"] = df["Visitas Realizadas"] / df["Visitas Programadas"]
        
        # Considerar tiempo de traslados
        df["DIF"] = df["DIF"] + self.config.TIEMPO_TRASLADO
        df["EFECTIVIDAD HORAS"] = df["DIF"] / self.config.HORAS_LABORALES
        
        # Calcular alcance de horas
        condiciones = [
            (df["DIF"] - self.config.HORAS_LABORALES) >= 0,
            (df["DIF"] - self.config.HORAS_LABORALES) < 0
        ]
        
        resultados = [
            "TIEMPO RESTANTE " + (df["DIF"] - self.config.HORAS_LABORALES).round(2).astype(str) + " HORAS",
            "TIEMPO FALTANTE " + (-1 * (df["DIF"] - self.config.HORAS_LABORALES)).round(2).astype(str) + " HORAS"
        ]
        
        df["ALCANCE HORAS"] = np.select(condiciones, resultados, default="")
        
        return df
    
    def construir_reporte_supervisores(self, df_asistencia: pd.DataFrame, df_personal: pd.DataFrame) -> pd.DataFrame:
        """Construye el reporte por supervisores"""
        df_sup = pd.pivot_table(
            df_asistencia,
            index="Supervisor Asignado OK",
            values=["Visitas Programadas", "Visitas Realizadas", "DIF"],
            aggfunc="sum"
        ).reset_index()
        
        df_sup["DIFERENCIA VISITAS"] = df_sup["Visitas Programadas"] - df_sup["Visitas Realizadas"]
        df_sup["EFECTIVIDAD VISITAS"] = df_sup["Visitas Realizadas"] / df_sup["Visitas Programadas"]
        
        # Contar promotores por supervisor
        df_sup_count = pd.pivot_table(
            df_asistencia,
            index="Supervisor Asignado OK",
            values=["Coordinador Asignado"],
            aggfunc="count"
        ).reset_index()
        
        df_sup = pd.merge(df_sup, df_sup_count, on="Supervisor Asignado OK", how="left")
        df_sup["EFECTIVIDAD HORAS"] = df_sup["DIF"] / (
            self.config.HORAS_LABORALES * df_sup["Coordinador Asignado"]
        )
        
        # Agregar información de ruta y coordinador
        df_sc = df_personal[["RUTA", "Coordinador Asignado", "Supervisor Asignado OK"]].drop_duplicates(
            subset="Supervisor Asignado OK"
        )
        df_sup = pd.merge(df_sup, df_sc, on="Supervisor Asignado OK", how="left")
        
        # Procesar RUTA
        df_sup["RUTA"] = df_sup["RUTA"].astype(str)
        condiciones = [
            df_sup["RUTA"].str.len() == 3,
            df_sup["RUTA"].str.len() > 3
        ]
        resultados = [
            df_sup["RUTA"],
            df_sup["RUTA"].str[:3]
        ]
        df_sup["RUTA"] = np.select(condiciones, resultados, default="OK")
        
        # Seleccionar columnas finales
        df_sup = df_sup[[
            "Coordinador Asignado_y", "Supervisor Asignado OK", "RUTA", 
            "Coordinador Asignado_x", "Visitas Programadas", "Visitas Realizadas", 
            "DIFERENCIA VISITAS", "EFECTIVIDAD VISITAS", "DIF", "EFECTIVIDAD HORAS"
        ]]
        
        return df_sup
    
    def construir_reporte_tiendas(self, df_efectividad: pd.DataFrame, df_rutero: pd.DataFrame) -> pd.DataFrame:
        """Construye el reporte por tiendas"""
        df_tiendas_r = df_rutero[["Codigo RO", "Nombre de Tienda", "Clasificacion"]]
        df_tiendas = pd.merge(
            df_efectividad, 
            df_tiendas_r, 
            left_on="Código Tienda", 
            right_on="Codigo RO", 
            how="left"
        )
        
        df_tiendas = pd.pivot_table(
            df_tiendas,
            index=["Supervisor", "Personal Promotor", "Nombre de Tienda", "Clasificacion"],
            values="Visitas Realizadas",
            aggfunc="sum"
        ).reset_index()
        
        df_tiendas = df_tiendas[df_tiendas["Visitas Realizadas"] > 0]
        
        return df_tiendas


class GestorArchivos:
    """Gestiona la creación y manipulación de archivos"""
    
    @staticmethod
    def guardar_excel(ruta: str, nombre: str, dataframes: Dict[str, pd.DataFrame]):
        """Guarda múltiples dataframes en un archivo Excel"""
        ruta_completa = os.path.join(ruta, nombre)
        try:
            with pd.ExcelWriter(ruta_completa, engine='openpyxl') as writer:
                for sheet_name, df in dataframes.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"✓ Archivo creado: {nombre}")
        except Exception as e:
            print(f"❌ Error al guardar Excel: {e}")
    
    @staticmethod
    def limpiar_carpeta(ruta: str):
        """Elimina todos los archivos de una carpeta"""
        try:
            for archivo in os.listdir(ruta):
                ruta_archivo = os.path.join(ruta, archivo)
                if os.path.isfile(ruta_archivo):
                    os.remove(ruta_archivo)
            print(f"✓ Carpeta limpia: {ruta}")
        except Exception as e:
            print(f"❌ Error al limpiar carpeta: {e}")


class EjecutorMacro:
    """Ejecuta macros de Excel"""
    
    @staticmethod
    def ejecutar_macro(ruta_libro: str, nombre_macro: str):
        """Ejecuta una macro específica en un libro de Excel"""
        print("▶ Ejecutando macro...")
        excel = None
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = True
            
            libro = excel.Workbooks.Open(ruta_libro)
            nombre_libro = os.path.basename(ruta_libro)
            excel.Application.Run(f"'{nombre_libro}'!{nombre_macro}")
            
            print("✓ Macro ejecutada correctamente")
        except Exception as e:
            print(f"⚠️  La macro cerró el libro: {e}")
        finally:
            if excel:
                try:
                    excel.Quit()
                except:
                    print("⚠️  Excel ya estaba cerrado")


class InterfazUsuario:
    """Maneja la interacción con el usuario"""
    
    @staticmethod
    def solicitar_dia() -> str:
        """Solicita al usuario el día de la semana"""
        print("\n" + "="*60)
        print("DÍAS VÁLIDOS:")
        print("LUNES: LUN | MARTES: MAR | MIÉRCOLES: MIER")
        print("JUEVES: JUE | VIERNES: VIE | SÁBADO: SAB | DOMINGO: DOM")
        print("="*60)
        
        while True:
            try:
                dia = input("\nIntroduzca el día de la semana (abreviado): ").upper().strip()
                if dia in ConfiguracionAsistencia.DIAS_VALIDOS:
                    return dia
                else:
                    print("❌ Día inválido. Intente de nuevo.")
            except Exception as e:
                print(f"❌ Error: {e}")
    
    @staticmethod
    def solicitar_semana() -> int:
        """Solicita al usuario el número de semana"""
        print("\n" + "="*60)
        print("SEMANAS VÁLIDAS: 1, 2, 3, 4, 5")
        print("="*60)
        
        while True:
            try:
                semana = int(input("\nIntroduzca el número de semana (1-5): "))
                if semana in ConfiguracionAsistencia.SEMANAS_VALIDAS:
                    return semana
                else:
                    print("❌ Número fuera de rango. Intente de nuevo.")
            except ValueError:
                print("❌ Entrada inválida. Introduzca un número.")


class SistemaAsistencia:
    """Clase principal que orquesta todo el sistema"""
    
    def __init__(self):
        self.config = ConfiguracionAsistencia()
        self.cargador = CargadorArchivos(self.config.RUTA_BASE)
        self.proc_efectividad = ProcesadorEfectividad(self.config)
        self.proc_rutero = ProcesadorRutero(self.config)
        self.constructor = ConstructorReporte(self.config)
        self.gestor = GestorArchivos()
        self.interfaz = InterfazUsuario()
    
    def ejecutar(self):
        """Ejecuta el flujo completo del sistema"""
        print("\n" + "="*60)
        print("SISTEMA DE PROCESAMIENTO DE ASISTENCIA")
        print("="*60 + "\n")
        
        # 1. Cargar archivos
        print("📁 Cargando archivos...")
        df_efectividad = self.cargador.cargar_efectividad()
        df_personal = self.cargador.cargar_personal()
        df_rutero_raw = self.cargador.cargar_rutero()
        
        if df_efectividad is None or df_personal is None or df_rutero_raw is None:
            print("❌ Error: No se pudieron cargar todos los archivos necesarios")
            return
        
        # 2. Procesar efectividad
        print("\n⚙️  Procesando datos de efectividad...")
        df_efectividad = self.proc_efectividad.procesar_efectividad(df_efectividad)
        df_asistencia = self.proc_efectividad.consolidar_asistencia(df_efectividad)
        
        # 3. Procesar rutero
        print("⚙️  Procesando rutero...")
        df_rutero = self.proc_rutero.procesar_rutero(df_rutero_raw)
        
        # 4. Solicitar día y semana
        dia = self.interfaz.solicitar_dia()
        semana = self.interfaz.solicitar_semana()
        
        df_programadas = self.proc_rutero.obtener_visitas_programadas(df_rutero, dia, semana)
        
        # 5. Construir reportes
        print("\n📊 Construyendo reportes...")
        df_reporte_principal = self.constructor.construir_reporte_principal(
            df_asistencia, df_programadas, df_personal
        )
        df_reporte_supervisores = self.constructor.construir_reporte_supervisores(
            df_reporte_principal, df_personal
        )
        df_reporte_tiendas = self.constructor.construir_reporte_tiendas(
            df_efectividad, df_rutero_raw
        )
        
        # 6. Guardar archivo Excel
        print("\n💾 Guardando archivo Excel...")
        fecha_hoy = datetime.today().strftime("%d%m%Y")
        nombre_archivo = "ASISTENCIA.xlsx"
        
        dataframes = {
            'ASISTENCIA': df_reporte_principal,
            'SUPERVISORES': df_reporte_supervisores,
            'TIENDAS': df_reporte_tiendas
        }
        
        self.gestor.guardar_excel(self.config.RUTA_BASE, nombre_archivo, dataframes)
        
        # 7. Limpiar carpeta supervisores
        print("\n🧹 Limpiando carpeta SUP...")
        self.gestor.limpiar_carpeta(self.config.RUTA_SUP)
        
        # 8. Ejecutar macro
        print("\n📝 Ejecutando macro de Excel...")
        EjecutorMacro.ejecutar_macro(self.config.RUTA_PLANTILLA, self.config.MACRO_NOMBRE)
        
        print("\n" + "="*60)
        print("✅ PROCESO COMPLETADO EXITOSAMENTE")
        print("="*60 + "\n")


def main():
    """Función principal"""
    try:
        sistema = SistemaAsistencia()
        sistema.ejecutar()
    except Exception as e:
        print(f"\n❌ Error fatal en el sistema: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()