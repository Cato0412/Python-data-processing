import pandas as pd
import openpyxl
from tkinter import Tk, filedialog
import os

def seleccionar_archivo(titulo):
    """Abre un diálogo para seleccionar un archivo"""
    root = Tk()
    root.withdraw()
    archivo = filedialog.askopenfilename(
        title=titulo,
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    root.destroy()
    return archivo

def main():
    print("=" * 60)
    print("CRUCE DE INFORMACIÓN - RUTERO Y COBERTURASERVICIO")
    print("=" * 60)
    
    # Seleccionar archivo RUTERO
    print("\n1. Selecciona el archivo que contiene 'RUTERO'")
    archivo_rutero = seleccionar_archivo("Selecciona el archivo RUTERO")
    
    if not archivo_rutero or 'RUTERO' not in archivo_rutero.upper():
        print("⚠️  Advertencia: El archivo seleccionado no contiene 'RUTERO' en su nombre")
        continuar = input("¿Deseas continuar de todos modos? (s/n): ")
        if continuar.lower() != 's':
            print("Operación cancelada")
            return
    
    # Seleccionar archivo COBERTURASERVICIO
    print("\n2. Selecciona el archivo que contiene 'COBERTURASERVICIO'")
    archivo_cobertura = seleccionar_archivo("Selecciona el archivo COBERTURASERVICIO")
    
    if not archivo_cobertura or 'COBERTURASERVICIO' not in archivo_cobertura.upper():
        print("⚠️  Advertencia: El archivo seleccionado no contiene 'COBERTURASERVICIO' en su nombre")
        continuar = input("¿Deseas continuar de todos modos? (s/n): ")
        if continuar.lower() != 's':
            print("Operación cancelada")
            return
    
    print("\n" + "=" * 60)
    print("PROCESANDO ARCHIVOS...")
    print("=" * 60)
    
    try:
        # Cargar todos los archivos Excel con todas sus hojas
        excel_rutero = pd.ExcelFile(archivo_rutero)
        excel_cobertura = pd.ExcelFile(archivo_cobertura)
        
        print(f"\n📄 Archivo RUTERO contiene las hojas: {excel_rutero.sheet_names}")
        print(f"📄 Archivo COBERTURASERVICIO contiene las hojas: {excel_cobertura.sheet_names}")
        
        # Verificar que existan las hojas necesarias
        if 'RUTERO' not in excel_rutero.sheet_names:
            print("\n❌ Error: No se encontró la hoja 'RUTERO' en el primer archivo")
            return
        
        if 'TERRITORIO' not in excel_cobertura.sheet_names:
            print("\n❌ Error: No se encontró la hoja 'TERRITORIO' en el archivo COBERTURASERVICIO")
            return
        
        # Leer las hojas específicas para el cruce
        df_rutero = pd.read_excel(archivo_rutero, sheet_name='RUTERO')
        df_territorio = pd.read_excel(archivo_cobertura, sheet_name='TERRITORIO ')
        
        print("\n✅ Hojas cargadas correctamente")
        print(f"   - RUTERO: {len(df_rutero)} filas")
        print(f"   - TERRITORIO: {len(df_territorio)} filas")
        
        # Verificar columnas necesarias
        columnas_rutero_necesarias = ['TIENDA ID_CUBO', 'ID_TIENDA']
        columnas_faltantes_rutero = [col for col in columnas_rutero_necesarias if col not in df_rutero.columns]
        
        if columnas_faltantes_rutero:
            print(f"\n❌ Error: Faltan columnas en RUTERO: {columnas_faltantes_rutero}")
            print(f"   Columnas disponibles: {list(df_rutero.columns)}")
            return
        
        if 'NOMBRE COMPLETO' not in df_territorio.columns:
            print(f"\n❌ Error: Falta la columna 'NOMBRE COMPLETO' en TERRITORIO")
            print(f"   Columnas disponibles: {list(df_territorio.columns)}")
            return
        
        print("\n✅ Todas las columnas necesarias están presentes")
        
        # Crear diccionario para el mapeo desde TERRITORIO
        # Usamos ID_TIENDA como clave para hacer el cruce
        territorio_dict = df_territorio.set_index('NOMBRE COMPLETO')['NOMBRE COMPLETO'].to_dict()
        
        # Realizar el cruce y actualización
        print("\n🔄 Realizando cruce de información...")
        actualizaciones = 0
        
        # Comparar y actualizar
        for idx, row in df_rutero.iterrows():
            tienda_id = row.get('ID_TIENDA')
            nombre_actual = row.get('TIENDA ID_CUBO')
            
            # Buscar coincidencia en TERRITORIO por ID_TIENDA o nombre
            if pd.notna(tienda_id):
                # Buscar en territorio si existe un nombre completo correspondiente
                matches = df_territorio[df_territorio['NOMBRE COMPLETO'].str.contains(str(tienda_id), case=False, na=False)]
                
                if not matches.empty:
                    nombre_territorio = matches.iloc[0]['NOMBRE COMPLETO']
                    
                    # Si no coinciden, actualizar
                    if str(nombre_actual) != str(nombre_territorio):
                        df_rutero.at[idx, 'TIENDA ID_CUBO'] = nombre_territorio
                        actualizaciones += 1
        
        print(f"✅ Se realizaron {actualizaciones} actualizaciones")
        
        # Crear archivo de salida
        nombre_salida = os.path.join(
            os.path.dirname(archivo_rutero),
            "RUTERO_ACTUALIZADO.xlsx"
        )
        
        print(f"\n💾 Guardando archivo actualizado...")
        
        # Usar openpyxl para mantener todas las hojas
        with pd.ExcelWriter(nombre_salida, engine='openpyxl') as writer:
            # Guardar la hoja RUTERO actualizada
            df_rutero.to_excel(writer, sheet_name='RUTERO', index=False)
            
            # Copiar las demás hojas del archivo original
            for sheet_name in excel_rutero.sheet_names:
                if sheet_name != 'RUTERO':
                    df_temp = pd.read_excel(archivo_rutero, sheet_name=sheet_name)
                    df_temp.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"\n✅ ¡Proceso completado exitosamente!")
        print(f"📁 Archivo guardado en: {nombre_salida}")
        print(f"📊 Total de actualizaciones: {actualizaciones}")
        
    except Exception as e:
        print(f"\n❌ Error durante el procesamiento: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
    input("\nPresiona Enter para salir...")