import pandas as pd
from pathlib import Path
import re

# ===================== CONFIGURACIÓN =====================
CARPETA = r"C:\Users\lapmxdf558\Documents\Archivos Alejandro\Genomma Mayoreo\BI\precios\Competencia"
ARCHIVO_SALIDA = "Base_PreciosCompetencia_Normalizada.xlsx"

# ===================== FUNCIONES =====================

def extraer_nombre_producto(columna):
    """
    Extrae el nombre del producto limpiando prefijos de forma agresiva
    """
    texto = str(columna).strip()
    
    # Remover TODOS los prefijos posibles de forma más agresiva
    # Importante: ordenar de más específico a menos específico
    patrones_remover = [
        # Precios con variaciones
        r'^precio\s+promoción\s+o\s+descuento\s+precio\s+regular\s+',
        r'^precio\s+regular\s+\([^)]+\)\s+de\s+',
        r'^precio\s+regular\s+\([^)]+\)\s+',
        r'^precio\s+promoción\s+o\s+descuento\s*',
        r'^precio\s+promocion\s+o\s+descuento\s*',
        r'^precio\s+promoción\s*',
        r'^precio\s+promocion\s*',
        r'^precio\s+regular\s+',
        r'^precio\s+',
        # Inventarios
        r'^inventario\s+',
        # Fotos
        r'^foto\s+de\s+la\s+categoría\s*',
    ]
    
    # Aplicar cada patrón
    for patron in patrones_remover:
        texto = re.sub(patron, '', texto, flags=re.IGNORECASE)
    
    # Limpieza final: remover espacios múltiples y strips
    texto = re.sub(r'\s+', ' ', texto).strip()
    
    return texto

def limpiar_valor_numerico(valor):
    """
    Convierte un valor a numérico PURO, eliminando TODO excepto números y puntos
    """
    if pd.isna(valor):
        return None
    
    valor_str = str(valor).strip()
    
    # Si está vacío
    if valor_str == '' or valor_str.lower() in ['nan', 'none', 'n/a']:
        return None
    
    # LIMPIAR: Eliminar TODO excepto números, puntos y signos negativos
    # Remover: letras, símbolos de moneda, comas, espacios, etc.
    valor_limpio = re.sub(r'[^\d.\-]', '', valor_str)
    
    # Si después de limpiar queda vacío
    if valor_limpio == '' or valor_limpio == '.':
        return None
    
    # Manejar múltiples puntos (ej: "1.234.56" -> "1234.56")
    partes = valor_limpio.split('.')
    if len(partes) > 2:
        # Si hay varios puntos, asumir que los primeros son separadores de miles
        valor_limpio = ''.join(partes[:-1]) + '.' + partes[-1]
    
    # Intentar convertir a número
    try:
        numero = float(valor_limpio)
        
        # Si es un entero, devolverlo como entero
        if numero.is_integer():
            return int(numero)
        
        return numero
    except:
        # Si aún así falla, devolver None
        return None

def procesar_archivo(ruta_archivo):
    """
    Procesa un archivo Excel despivoteando productos
    """
    nombre = Path(ruta_archivo).name
    print(f"\n{'='*70}")
    print(f"📄 {nombre}")
    print('='*70)
    
    try:
        # Leer archivo
        df = pd.read_excel(ruta_archivo)
        df = df.loc[:, ~df.columns.duplicated()]
        
        print(f"📊 Dimensiones: {df.shape[0]:,} filas × {df.shape[1]} columnas")
        
        # Identificar columnas metadata (NO productos)
        cols_metadata = []
        for col in df.columns:
            col_lower = str(col).lower()
            es_metadata = any(x in col_lower for x in [
                'instancia', 'proyecto', 'encuesta', 'tienda', 'encuestador',
                'comunidad', 'estado', 'municipio', 'zona', 'región',
                'fecha subida', 'geolocalización', 'sku', 
                'descripción', 'cantidad', 'total', 'presentación'
            ])
            # Excluir Fecha Respuesta duplicadas pero NO las fotos
            if es_metadata and not re.match(r'fecha respuesta\.\d+', col_lower):
                cols_metadata.append(col)
        
        # Identificar columnas de producto
        cols_inventario = {}
        cols_precio_reg = {}
        cols_precio_promo = {}
        cols_foto = {}  # Para fotos asociadas a productos
        
        for col in df.columns:
            if col in cols_metadata:
                continue
                
            col_lower = str(col).lower()
            
            # INVENTARIO (solo columnas explícitas de "Inventario")
            if 'inventario' in col_lower:
                producto = extraer_nombre_producto(col)
                cols_inventario[producto] = col
            
            # FOTO DE CATEGORÍA (asociar a productos si es posible)
            elif 'foto' in col_lower and 'categoría' in col_lower:
                cols_foto[col] = col
            
            # PRECIO REGULAR
            elif 'precio regular' in col_lower or \
                 ('sin descuento' in col_lower and 'precio' in col_lower) or \
                 ('sin promoción' in col_lower and 'precio' in col_lower):
                producto = extraer_nombre_producto(col)
                if producto not in cols_precio_reg:
                    cols_precio_reg[producto] = col
            
            # PRECIO PROMOCIÓN
            elif 'promoción' in col_lower or 'promocion' in col_lower:
                if 'precio' in col_lower or 'descuento' in col_lower:
                    producto = extraer_nombre_producto(col)
                    if producto not in cols_precio_promo:
                        cols_precio_promo[producto] = col
        
        print(f"\n📦 Columnas detectadas:")
        print(f"   • Metadata: {len(cols_metadata)}")
        print(f"   • Inventario: {len(cols_inventario)}")
        print(f"   • Precio Regular: {len(cols_precio_reg)}")
        print(f"   • Precio Promoción: {len(cols_precio_promo)}")
        print(f"   • Fotos: {len(cols_foto)}")
        
        # Obtener todos los productos únicos
        todos_productos = set()
        todos_productos.update(cols_inventario.keys())
        todos_productos.update(cols_precio_reg.keys())
        todos_productos.update(cols_precio_promo.keys())
        
        if not todos_productos:
            print(f"   ⚠️  No se detectaron productos")
            return None
        
        print(f"   • Productos únicos: {len(todos_productos)}")
        
        # Crear registros normalizados
        registros = []
        
        for idx, row in df.iterrows():
            # Datos base (metadata)
            datos_base = {}
            for col in cols_metadata:
                if col in df.columns:
                    datos_base[col] = row[col]
            
            # Agregar las fotos generales (no asociadas a productos específicos)
            for col_foto in cols_foto:
                datos_base[col_foto] = row[col_foto]
            
            datos_base['archivo_fuente'] = nombre
            
            # Por cada producto único, crear un registro
            for producto in todos_productos:
                registro = datos_base.copy()
                registro['producto'] = producto
                
                # Inventario (limpiar y convertir a numérico)
                val_inv = row[cols_inventario[producto]] if producto in cols_inventario else None
                registro['inventario'] = limpiar_valor_numerico(val_inv)
                
                # Precio Regular (limpiar y convertir a numérico)
                val_reg = row[cols_precio_reg[producto]] if producto in cols_precio_reg else None
                registro['precio_regular'] = limpiar_valor_numerico(val_reg)
                
                # Precio Promoción (limpiar y convertir a numérico)
                val_promo = row[cols_precio_promo[producto]] if producto in cols_precio_promo else None
                registro['precio_promocion'] = limpiar_valor_numerico(val_promo)
                
                registros.append(registro)
        
        # Crear DataFrame
        df_resultado = pd.DataFrame(registros)
        
        # Eliminar filas completamente vacías
        df_resultado = df_resultado[
            df_resultado[['inventario', 'precio_regular', 'precio_promocion']].notna().any(axis=1)
        ]
        
        print(f"✅ Registros normalizados: {len(df_resultado):,}")
        
        return df_resultado
        
    except Exception as e:
        print(f"❌ ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

def procesar_todos():
    """
    Procesa todos los archivos
    """
    print("\n" + "="*70)
    print("🚀 NORMALIZADOR DE ENCUESTAS DE PRECIOS")
    print("="*70)
    
    # Buscar archivos
    archivos = list(Path(CARPETA).glob("*.xlsx")) + list(Path(CARPETA).glob("*.xls"))
    archivos = [f for f in archivos if not f.name.startswith("~$") and f.name != ARCHIVO_SALIDA]
    
    print(f"\n📁 Archivos encontrados: {len(archivos)}")
    for i, archivo in enumerate(archivos, 1):
        print(f"   {i}. {archivo.name}")
    
    if not archivos:
        print("❌ No se encontraron archivos")
        return None
    
    # Procesar cada archivo
    todos_df = []
    for archivo in archivos:
        df = procesar_archivo(archivo)
        if df is not None and not df.empty:
            todos_df.append(df)
    
    if not todos_df:
        print("\n❌ No se pudo procesar ningún archivo")
        return None
    
    # Combinar todos
    print("\n" + "="*70)
    print("🔗 COMBINANDO TODOS LOS ARCHIVOS")
    print("="*70)
    
    # Unificar columnas
    todas_columnas = set()
    for df in todos_df:
        todas_columnas.update(df.columns)
    
    # Asegurar que todos tengan las mismas columnas
    for i, df in enumerate(todos_df):
        for col in todas_columnas:
            if col not in df.columns:
                df[col] = None
        todos_df[i] = df
    
    resultado = pd.concat(todos_df, ignore_index=True)
    
    print(f"\n📊 RESULTADO FINAL:")
    print(f"   • Total de registros: {len(resultado):,}")
    print(f"   • Productos únicos: {resultado['producto'].nunique():,}")
    print(f"   • Archivos procesados: {resultado['archivo_fuente'].nunique()}")
    
    return resultado

def guardar_excel(df):
    """
    Guarda el archivo Excel ordenado y LIMPIO
    """
    print("\n🔧 Limpiando datos antes de guardar...")
    
    # Limpiar datos problemáticos
    df_limpio = df.copy()
    
    # Convertir fechas a string para evitar problemas
    for col in df_limpio.columns:
        if df_limpio[col].dtype == 'datetime64[ns]':
            df_limpio[col] = df_limpio[col].astype(str)
    
    # Reemplazar valores problemáticos
    df_limpio = df_limpio.replace([float('inf'), float('-inf')], None)
    df_limpio = df_limpio.fillna('')
    
    # Limpiar nombres de columnas
    df_limpio.columns = [str(col).strip() for col in df_limpio.columns]
    
    # IMPORTANTE: Definir orden óptimo de columnas
    columnas_orden = [
        # Identificadores primero
        '# Instancia',
        'archivo_fuente',
        'Proyecto',
        'Encuesta',
        
        # Ubicación y tienda
        'Id de Tienda',
        'Encuestador/Tienda',
        'Estado',
        'Municipio',
        'Zona',
        'Región',
        
        # Producto y sus datos (lo más importante)
        'producto',
        'inventario',
        'precio_regular', 
        'precio_promocion',
        
        # Fechas
        'Fecha Subida',
        'Fecha Respuesta'
    ]
    
    # Agregar todas las columnas de fotos al final
    columnas_fotos = sorted([col for col in df_limpio.columns if 'foto' in col.lower() and 'categoría' in col.lower()])
    
    # Combinar: orden definido + fotos al final
    columnas_finales = columnas_orden + columnas_fotos
    
    # Filtrar solo las columnas que existen y queremos
    columnas_existentes = [col for col in columnas_finales if col in df_limpio.columns]
    
    # Crear DataFrame final SOLO con las columnas necesarias
    df_final = df_limpio[columnas_existentes].copy()
    
    print(f"🗑️  Columnas no relacionadas eliminadas")
    print(f"✅ Columnas conservadas: {len(columnas_existentes)}")
    print(f"📸 Columnas de fotos: {len([c for c in columnas_existentes if 'foto' in c.lower()])}")
    
    # Guardar
    print("💾 Guardando archivo...")
    ruta_salida = Path(CARPETA) / ARCHIVO_SALIDA
    
    try:
        df_final.to_excel(ruta_salida, index=False, engine='openpyxl')
        print("✅ Guardado exitoso!")
    except Exception as e:
        print(f"❌ Error al guardar: {e}")
        # Intentar guardar como CSV alternativo
        ruta_csv = Path(CARPETA) / ARCHIVO_SALIDA.replace('.xlsx', '.csv')
        df_final.to_csv(ruta_csv, index=False, encoding='utf-8-sig')
        print(f"✅ Guardado como CSV alternativo: {ruta_csv.name}")
        return
    
    print(f"\n✅ ARCHIVO GUARDADO: {ruta_salida.name}")
    print(f"   Ubicación: {ruta_salida}")
    print(f"   Dimensiones: {df_final.shape[0]:,} filas × {df_final.shape[1]} columnas")
    
    # Muestra de datos (sin fotos para no saturar)
    print(f"\n📋 MUESTRA DE DATOS (primeras 10 filas, sin URLs):")
    columnas_muestra = ['archivo_fuente', 'producto', 'inventario', 'precio_regular', 'precio_promocion', 'Estado']
    columnas_muestra = [c for c in columnas_muestra if c in df_final.columns]
    muestra = df_final[columnas_muestra].head(10).copy()
    
    # Acortar nombres largos para mejor visualización
    if 'producto' in muestra.columns:
        muestra['producto'] = muestra['producto'].astype(str).str[:30]
    if 'archivo_fuente' in muestra.columns:
        muestra['archivo_fuente'] = muestra['archivo_fuente'].astype(str).str[:20]
    
    print(muestra.to_string(index=False))
    
    # Estadísticas
    print(f"\n📈 ESTADÍSTICAS:")
    
    # Contar valores reales (no None, no vacíos, no ceros)
    if 'inventario' in df_final.columns:
        con_inv = df_final['inventario'].notna() & (df_final['inventario'] != '') & (df_final['inventario'] != 0)
        total_inv = con_inv.sum()
    else:
        total_inv = 0
    
    if 'precio_regular' in df_final.columns:
        con_reg = df_final['precio_regular'].notna() & (df_final['precio_regular'] != '') & (df_final['precio_regular'] != 0)
        total_reg = con_reg.sum()
    else:
        total_reg = 0
    
    if 'precio_promocion' in df_final.columns:
        con_promo = df_final['precio_promocion'].notna() & (df_final['precio_promocion'] != '') & (df_final['precio_promocion'] != 0)
        total_promo = con_promo.sum()
    else:
        total_promo = 0
    
    print(f"   • Registros con inventario: {total_inv:,}")
    print(f"   • Registros con precio regular: {total_reg:,}")
    print(f"   • Registros con precio promoción: {total_promo:,}")
    
    # Mostrar promedios si hay datos numéricos
    if 'precio_regular' in df_final.columns:
        precios_numericos = pd.to_numeric(df_final['precio_regular'], errors='coerce')
        if precios_numericos.notna().any():
            promedio = precios_numericos.mean()
            print(f"   • Precio regular promedio: ${promedio:.2f}")

# ===================== EJECUTAR =====================

if __name__ == "__main__":
    df_final = procesar_todos()
    
    if df_final is not None:
        guardar_excel(df_final)
        print("\n" + "="*70)
        print("✨ PROCESO COMPLETADO EXITOSAMENTE")
        print("="*70)
    else:
        print("\n❌ No se pudo completar el proceso")