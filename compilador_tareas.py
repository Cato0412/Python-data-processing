import pandas as pd
from pathlib import Path

# ===================== CONFIGURACIÓN =====================
CARPETA = r"C:\Users\lapmxdf558\Documents\Archivos Alejandro\Genomma Mayoreo\Rutero Mayoreo\Rutero Enero\Tareas"
ARCHIVO_SALIDA = "Encuestas_Apiladas.xlsx"

# ===================== FUNCIONES =====================

def leer_archivo_excel(ruta):
    """
    Lee un archivo Excel y retorna el DataFrame
    """
    try:
        df = pd.read_excel(ruta)
        # Eliminar columnas duplicadas
        df = df.loc[:, ~df.columns.duplicated()]
        return df
    except Exception as e:
        print(f"   ❌ Error: {str(e)}")
        return None

def diagnosticar_archivo(df, nombre):
    """
    Muestra información del archivo
    """
    print(f"\n   📊 Dimensiones: {df.shape[0]:,} filas × {df.shape[1]} columnas")
    print(f"   📋 Columnas: {', '.join(df.columns[:8])}")
    if len(df.columns) > 8:
        print(f"              ... y {len(df.columns) - 8} más")

def encontrar_columnas_comunes(lista_dfs):
    """
    Encuentra las columnas que están en TODOS los DataFrames
    Incluye automáticamente columnas de fotos aunque no estén en todos
    """
    if not lista_dfs:
        return set()
    
    # Empezar con las columnas del primer DataFrame
    columnas_comunes = set(lista_dfs[0].columns)
    
    # Intersectar con las columnas de cada DataFrame
    for df in lista_dfs[1:]:
        columnas_comunes = columnas_comunes & set(df.columns)
    
    # NUEVO: Agregar TODAS las columnas de fotos aunque no estén en todos
    columnas_fotos = set()
    for df in lista_dfs:
        for col in df.columns:
            if 'foto' in col.lower() and 'categoría' in col.lower():
                columnas_fotos.add(col)
    
    # Combinar columnas comunes + todas las fotos
    columnas_finales = columnas_comunes | columnas_fotos
    
    return columnas_finales

def apilar_con_columnas_comunes(lista_dfs, nombres_archivos):
    """
    Apila los DataFrames usando columnas comunes + todas las fotos
    """
    print("\n" + "="*70)
    print("🔍 ANÁLISIS DE COLUMNAS")
    print("="*70)
    
    # Encontrar columnas comunes (incluye fotos)
    columnas_finales = encontrar_columnas_comunes(lista_dfs)
    
    # Separar columnas comunes reales de columnas de fotos
    columnas_fotos = {col for col in columnas_finales if 'foto' in col.lower() and 'categoría' in col.lower()}
    columnas_comunes_reales = columnas_finales - columnas_fotos
    
    print(f"\n✅ Columnas comunes en TODOS los archivos ({len(columnas_comunes_reales)}):")
    for col in sorted(columnas_comunes_reales):
        print(f"   • {col}")
    
    if columnas_fotos:
        print(f"\n📸 Columnas de FOTOS incluidas ({len(columnas_fotos)}):")
        for col in sorted(columnas_fotos):
            print(f"   • {col}")
    
    # Mostrar columnas únicas por archivo (excluyendo fotos)
    print(f"\n📋 Otras columnas únicas por archivo:")
    for i, df in enumerate(lista_dfs):
        cols_unicas = set(df.columns) - columnas_finales
        if cols_unicas:
            print(f"\n   {nombres_archivos[i]} ({len(cols_unicas)} únicas):")
            for col in sorted(list(cols_unicas)[:10]):
                print(f"      • {col}")
            if len(cols_unicas) > 10:
                print(f"      ... y {len(cols_unicas) - 10} más")
    
    # Apilar usando columnas finales
    print("\n" + "="*70)
    print("📦 APILANDO DATOS")
    print("="*70)
    
    lista_filtrada = []
    for i, df in enumerate(lista_dfs):
        # Agregar columnas faltantes (especialmente fotos) con None
        df_expandido = df.copy()
        for col in columnas_finales:
            if col not in df_expandido.columns:
                df_expandido[col] = None
        
        # Filtrar solo columnas finales
        df_filtrado = df_expandido[sorted(columnas_finales)].copy()
        
        # Agregar columna de origen
        df_filtrado['archivo_origen'] = nombres_archivos[i]
        
        lista_filtrada.append(df_filtrado)
        
        print(f"   ✅ {nombres_archivos[i]}: {len(df_filtrado):,} registros")
    
    # Concatenar
    df_apilado = pd.concat(lista_filtrada, ignore_index=True)
    
    print(f"\n📊 RESULTADO:")
    print(f"   • Total de registros: {len(df_apilado):,}")
    print(f"   • Columnas finales: {len(df_apilado.columns)}")
    print(f"   • Columnas de fotos: {len(columnas_fotos)}")
    
    return df_apilado

def reordenar_columnas_apiladas(df):
    """
    Reordena las columnas del archivo apilado
    """
    print("\n📐 Reordenando columnas...")
    
    # Orden preferido (columnas comunes típicas)
    orden_preferido = [
        'archivo_origen',
        '# Instancia',
        'Proyecto',
        'Encuesta',
        'Id de Tienda',
        'Encuestador/Tienda',
        'Comunidad',
        'Estado',
        'Municipio',
        'Zona',
        'Región',
        'Fecha Subida',
        'Fecha Respuesta',
        'Geolocalización (Obligatoria)'
    ]
    
    # Columnas que existen en el orden preferido
    cols_ordenadas = [col for col in orden_preferido if col in df.columns]
    
    # Columnas de fotos (ordenadas)
    cols_fotos = sorted([col for col in df.columns if 'foto' in col.lower() and 'categoría' in col.lower()])
    
    # Columnas restantes (excluyendo las ya ordenadas y las fotos)
    cols_restantes = sorted([col for col in df.columns 
                            if col not in cols_ordenadas and col not in cols_fotos])
    
    # Combinar: orden preferido + restantes + fotos al final
    orden_final = cols_ordenadas + cols_restantes + cols_fotos
    
    return df[orden_final]

def generar_reporte_apilado(df):
    """
    Genera un reporte del archivo apilado
    """
    print("\n" + "="*70)
    print("📊 REPORTE DE DATOS APILADOS")
    print("="*70)
    
    print(f"\n📈 RESUMEN:")
    print(f"   • Total de registros: {len(df):,}")
    print(f"   • Total de columnas: {len(df.columns)}")
    
    if 'archivo_origen' in df.columns:
        print(f"\n📂 DISTRIBUCIÓN POR ARCHIVO:")
        dist = df['archivo_origen'].value_counts()
        for archivo, count in dist.items():
            pct = (count / len(df) * 100)
            print(f"   • {archivo:35} → {count:>7,} registros ({pct:>5.1f}%)")
    
    if 'Proyecto' in df.columns:
        print(f"\n🎯 PROYECTOS:")
        proyectos = df['Proyecto'].value_counts()
        for proyecto, count in proyectos.items():
            pct = (count / len(df) * 100)
            print(f"   • {proyecto:35} → {count:>7,} registros ({pct:>5.1f}%)")
    
    if 'Estado' in df.columns:
        print(f"\n📍 COBERTURA GEOGRÁFICA:")
        print(f"   • Estados únicos: {df['Estado'].nunique()}")
        estados_top = df['Estado'].value_counts().head(5)
        for estado, count in estados_top.items():
            print(f"     - {estado:25} → {count:>6,} registros")
    
    if 'Id de Tienda' in df.columns:
        print(f"   • Tiendas únicas: {df['Id de Tienda'].nunique()}")
    
    # Fotos
    cols_fotos = [col for col in df.columns if 'foto' in col.lower() and 'categoría' in col.lower()]
    if cols_fotos:
        print(f"\n📸 COLUMNAS DE FOTOS ({len(cols_fotos)}):")
        for col in cols_fotos:
            con_foto = df[col].notna().sum()
            pct = (con_foto / len(df) * 100) if len(df) > 0 else 0
            print(f"   • {col:35} → {con_foto:>6,} fotos ({pct:>5.1f}%)")
    
    # Completitud
    print(f"\n📊 COMPLETITUD DE COLUMNAS:")
    for col in df.columns:
        completos = df[col].notna().sum()
        pct = (completos / len(df) * 100) if len(df) > 0 else 0
        
        # Solo mostrar las primeras 15 columnas más importantes
        if col in ['archivo_origen', 'Proyecto', 'Encuesta', 'Id de Tienda', 
                   'Estado', 'Municipio', 'Zona', 'Región', 'Fecha Subida',
                   'Encuestador/Tienda', 'Fecha Respuesta', '# Instancia']:
            simbolo = "✅" if pct >= 90 else "⚠️" if pct >= 50 else "❌"
            print(f"   {simbolo} {col:30} → {completos:>7,}/{len(df):,} ({pct:>5.1f}%)")
    
    # Muestra de datos
    print(f"\n👀 MUESTRA DE DATOS (primeras 5 filas):")
    cols_muestra = ['archivo_origen', 'Proyecto', 'Estado', 'Municipio', 'Id de Tienda']
    cols_muestra = [c for c in cols_muestra if c in df.columns]
    
    muestra = df[cols_muestra].head(5)
    
    # Acortar nombres largos
    muestra_mostrar = muestra.copy()
    for col in muestra_mostrar.columns:
        if muestra_mostrar[col].dtype == 'object':
            muestra_mostrar[col] = muestra_mostrar[col].astype(str).str[:30]
    
    print(muestra_mostrar.to_string(index=False))

def guardar_excel(df, ruta):
    """
    Guarda el DataFrame en Excel
    """
    print(f"\n💾 Guardando archivo...")
    print(f"   📁 {ruta}")
    
    try:
        # Limpiar datos
        df_limpio = df.copy()
        
        # Convertir fechas a string
        for col in df_limpio.columns:
            if df_limpio[col].dtype == 'datetime64[ns]':
                df_limpio[col] = df_limpio[col].astype(str)
        
        # Reemplazar infinitos
        df_limpio = df_limpio.replace([float('inf'), float('-inf')], None)
        
        # Guardar
        df_limpio.to_excel(ruta, index=False, engine='openpyxl')
        
        tamaño = Path(ruta).stat().st_size / (1024 * 1024)  # MB
        print(f"\n✅ Archivo guardado exitosamente!")
        print(f"   📊 Tamaño: {tamaño:.2f} MB")
        
        return True
    except Exception as e:
        print(f"\n❌ ERROR al guardar: {str(e)}")
        return False

# ===================== PROCESO PRINCIPAL =====================

def main():
    print("="*70)
    print("📚 APILADOR DE ENCUESTAS")
    print("="*70)
    print(f"\n📁 Carpeta: {CARPETA}")
    
    # Buscar archivos Excel
    archivos = list(Path(CARPETA).glob("*.xlsx")) + list(Path(CARPETA).glob("*.xls"))
    archivos = [f for f in archivos if not f.name.startswith("~$") and 
                f.name not in [ARCHIVO_SALIDA, "Base_Precios_Normalizada.xlsx", 
                               "Competencia_Normalizada.xlsx", "Base_Unificada_Completa.xlsx"]]
    
    print(f"\n📂 Archivos encontrados: {len(archivos)}")
    
    if len(archivos) == 0:
        print("\n❌ No se encontraron archivos Excel para procesar")
        print("   Asegúrate de que hay archivos .xlsx o .xls en la carpeta")
        return
    
    # Leer todos los archivos
    lista_dfs = []
    nombres_archivos = []
    
    for archivo in archivos:
        print(f"\n📄 Leyendo: {archivo.name}")
        df = leer_archivo_excel(archivo)
        
        if df is not None:
            diagnosticar_archivo(df, archivo.name)
            lista_dfs.append(df)
            nombres_archivos.append(archivo.name)
    
    if not lista_dfs:
        print("\n❌ No se pudo leer ningún archivo")
        return
    
    print(f"\n✅ Archivos leídos exitosamente: {len(lista_dfs)}")
    
    # Apilar con columnas comunes
    df_apilado = apilar_con_columnas_comunes(lista_dfs, nombres_archivos)
    
    # Reordenar columnas
    df_apilado = reordenar_columnas_apiladas(df_apilado)
    
    # Generar reporte
    generar_reporte_apilado(df_apilado)
    
    # Guardar
    ruta_salida = Path(CARPETA) / ARCHIVO_SALIDA
    
    if guardar_excel(df_apilado, ruta_salida):
        print("\n" + "="*70)
        print("✨ PROCESO COMPLETADO EXITOSAMENTE")
        print("="*70)
        print(f"\n📦 Archivo final: {ARCHIVO_SALIDA}")
        print(f"   • {len(df_apilado):,} registros totales")
        print(f"   • {len(df_apilado.columns)} columnas comunes")
        print(f"   • {len(nombres_archivos)} archivos combinados")

# ===================== EJECUTAR =====================

if __name__ == "__main__":
    main()