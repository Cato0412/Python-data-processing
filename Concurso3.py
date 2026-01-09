import pandas as pd
import re

def normalizar_determinante(texto):
    """
    Extrae y normaliza el determinante/tipo de tienda del texto
    """
    texto_lower = texto.lower()
    
    # Mapeo de términos a códigos estándar
    mapeo = {
        'sc': 'SC',
        'walmart': 'SC',
        'wal mart': 'SC',
        'superama': 'SC',
        ' ba ': 'BA',
        '#ba': 'BA',
        'bodega aurrera': 'BA',
        'bodega': 'BA',
        'aurrera': 'BA',
        'sm': 'SM',
        "sam's": 'SM',
        'sams': 'SM',
        'sams club': 'SM',
        'mi bodega': 'MB',
        'mb': 'MB',
    }
    
    for patron, codigo in sorted(mapeo.items(), key=lambda x: -len(x[0])):
        if patron in texto_lower:
            return codigo
    
    return None

def extraer_tipo_tienda_excel(id_completo):
    """
    Extrae el tipo de tienda del ID completo del Excel
    Ejemplo: '2650_2070 - BA SOCONUSCO' -> 'BA'
    """
    # Buscar patrones como "- BA ", "- SC ", "- SM "
    match = re.search(r'-\s*([A-Z]{2,3})\s+', str(id_completo))
    if match:
        tipo = match.group(1)
        # Normalizar algunos casos especiales
        if tipo in ['BA', 'BD']:  # BD también es Bodega
            return 'BA'
        elif tipo in ['SC', 'WM']:  # WM también es Walmart
            return 'SC'
        elif tipo in ['SM', 'SA']:  # SA también es Sam's
            return 'SM'
        else:
            return tipo
    return None

def extraer_numeros(texto):
    """Extrae todos los números de 3-5 dígitos de un texto"""
    numeros = re.findall(r'\b\d{3,5}\b', str(texto))
    return numeros

def es_linea_tienda(linea):
    """Determina si una línea del chat menciona una tienda"""
    linea_lower = linea.lower()
    
    keywords = [
        'bodega', 'aurrera', 'ba ', '#ba', 
        'walmart', 'wal mart', 'sc ', '#sc',
        'sam', 'sams', 'sm ',
        'superama', 'tienda', 'sucursal', 'soriana'
    ]
    
    tiene_keyword = any(kw in linea_lower for kw in keywords)
    tiene_numero = bool(re.search(r'\b\d{3,5}\b', linea))
    
    es_sistema = any(x in linea for x in [
        'creó el grupo', 'añadió', 'te añadió', 
        'cambió', 'salió', 'eliminó', 'Los mensajes'
    ])
    
    return tiene_keyword and tiene_numero and not es_sistema

def buscar_por_id(archivo_txt, archivo_excel, columna_id='TIENDA ID_CUBO',
                  archivo_salida='coincidencias.xlsx'):
    """
    Busca coincidencias en un chat de WhatsApp
    """
    
    print("=" * 80)
    print("BUSCADOR DE TIENDAS EN CHAT DE WHATSAPP (VERSIÓN MEJORADA)")
    print("=" * 80)
    
    # Leer TXT
    try:
        with open(archivo_txt, 'r', encoding='utf-8') as f:
            lineas_todas = [linea.strip() for linea in f if linea.strip()]
        print(f"\n✓ Archivo TXT leído: {len(lineas_todas)} líneas totales")
    except Exception as e:
        print(f"\n✗ Error al leer TXT: {e}")
        return
    
    # Filtrar solo líneas de tiendas
    lineas_txt = [linea for linea in lineas_todas if es_linea_tienda(linea)]
    print(f"✓ Líneas filtradas que mencionan tiendas: {len(lineas_txt)}")
    
    print(f"\n📝 Ejemplos de líneas a analizar:")
    for i, linea in enumerate(lineas_txt[:5], 1):
        print(f"  {i}. {linea[:80]}...")
    
    # Leer Excel
    try:
        df = pd.read_excel(archivo_excel)
        print(f"\n✓ Archivo Excel leído: {len(df)} filas")
    except Exception as e:
        print(f"\n✗ Error al leer Excel: {e}")
        return
    
    if columna_id not in df.columns:
        print(f"\n✗ Error: La columna '{columna_id}' no existe")
        return
    
    # Procesar Excel
    df[columna_id] = df[columna_id].astype(str).str.strip()
    
    # Extraer ID numérico del formato: "2650_2070 - BA SOCONUSCO"
    df['ID_Numerico'] = df[columna_id].str.extract(r'_(\d+)', expand=False)
    
    # Extraer tipo de tienda del ID completo
    df['Tipo_Tienda'] = df[columna_id].apply(extraer_tipo_tienda_excel)
    
    # Crear clave de búsqueda: "BA_2070"
    df['Clave_Busqueda'] = df['Tipo_Tienda'] + '_' + df['ID_Numerico']
    
    # Crear también set de IDs solos
    ids_disponibles = set(df['ID_Numerico'].dropna())
    
    print(f"\n📊 Estadísticas del Excel:")
    print(f"   Total de tiendas: {len(df)}")
    print(f"   IDs únicos: {df['ID_Numerico'].nunique()}")
    
    # Mostrar tipos de tienda encontrados
    tipos = df['Tipo_Tienda'].value_counts()
    print(f"\n   Tipos de tienda detectados:")
    for tipo, count in tipos.head(10).items():
        print(f"     {tipo}: {count} tiendas")
    
    print(f"\n   Ejemplos de claves creadas:")
    for ejemplo in df['Clave_Busqueda'].dropna().head(5):
        print(f"     - {ejemplo}")
    
    # Analizar TXT
    print(f"\n🔍 Analizando chat de WhatsApp...")
    matches_encontrados = {}
    stats = {
        'con_tipo': 0,
        'sin_tipo': 0,
        'match_tipo_id': 0,
        'match_solo_id': 0,
        'sin_match': 0
    }
    
    for linea in lineas_txt:
        tipo_txt = normalizar_determinante(linea)
        numeros = extraer_numeros(linea)
        
        encontro_match = False
        
        # Estrategia 1: Buscar con tipo + ID
        if tipo_txt and numeros:
            stats['con_tipo'] += 1
            for num in numeros:
                clave = f"{tipo_txt}_{num}"
                if clave in df['Clave_Busqueda'].values:
                    if clave not in matches_encontrados:
                        matches_encontrados[clave] = linea
                        stats['match_tipo_id'] += 1
                        encontro_match = True
        
        # Estrategia 2: Buscar solo por ID (todas las tiendas con ese número)
        if not encontro_match and numeros:
            if not tipo_txt:
                stats['sin_tipo'] += 1
            
            for num in numeros:
                if num in ids_disponibles:
                    # Encontrar todas las tiendas con ese ID
                    matches = df[df['ID_Numerico'] == num]
                    for _, row in matches.iterrows():
                        clave = row['Clave_Busqueda']
                        if pd.notna(clave) and clave not in matches_encontrados:
                            matches_encontrados[clave] = linea
                            stats['match_solo_id'] += 1
                            encontro_match = True
        
        if not encontro_match:
            stats['sin_match'] += 1
    
    print(f"\n📊 Estadísticas del análisis:")
    print(f"   Líneas con tipo detectado (BA/SC/SM): {stats['con_tipo']}")
    print(f"   Líneas sin tipo detectado: {stats['sin_tipo']}")
    print(f"   Matches perfectos (Tipo+ID): {stats['match_tipo_id']}")
    print(f"   Matches solo por ID: {stats['match_solo_id']}")
    print(f"   Líneas sin match: {stats['sin_match']}")
    print(f"   Total de tiendas encontradas: {len(matches_encontrados)}")
    
    # Crear resultado
    df_resultado = df[df['Clave_Busqueda'].isin(matches_encontrados.keys())].copy()
    df_resultado['Linea_Original_TXT'] = df_resultado['Clave_Busqueda'].map(matches_encontrados)
    
    print(f"\n{'=' * 80}")
    print(f"RESULTADOS")
    print(f"{'=' * 80}")
    print(f"✓ Total de tiendas encontradas: {len(df_resultado)}")
    print(f"✓ Porcentaje de líneas con match: {len(matches_encontrados)/max(len(lineas_txt), 1)*100:.2f}%")
    
    if len(df_resultado) > 0:
        df_resultado.to_excel(archivo_salida, index=False, engine='openpyxl')
        print(f"\n✓ Archivo creado: '{archivo_salida}'")
        
        # Resumen por tipo
        print(f"\n📊 Resumen por tipo de tienda:")
        resumen = df_resultado.groupby('Tipo_Tienda').size().sort_values(ascending=False)
        for tipo, count in resumen.items():
            print(f"   {tipo}: {count} tiendas")
        
        # Preview
        print(f"\n{'=' * 80}")
        print("PREVIEW DE COINCIDENCIAS:")
        print(f"{'=' * 80}")
        
        for i, (_, row) in enumerate(df_resultado.head(20).iterrows(), 1):
            tipo = row['Tipo_Tienda'] if pd.notna(row['Tipo_Tienda']) else '??'
            print(f"\n{i}. {tipo} #{row['ID_Numerico']}")
            print(f"   ID Completo: {row[columna_id]}")
            if 'Nombre de Tienda' in df_resultado.columns:
                print(f"   Nombre: {row['Nombre de Tienda']}")
            print(f"   WhatsApp: {row['Linea_Original_TXT'][:90]}...")
        
        if len(df_resultado) > 20:
            print(f"\n... y {len(df_resultado) - 20} tiendas más")
    else:
        print("\n⚠️  No se encontraron coincidencias")
    
    # Mostrar líneas sin match
    lineas_sin_match = []
    for linea in lineas_txt:
        tipo = normalizar_determinante(linea)
        numeros = extraer_numeros(linea)
        
        tiene_match = False
        if numeros:
            for num in numeros:
                if num in ids_disponibles:
                    tiene_match = True
                    break
        
        if not tiene_match:
            lineas_sin_match.append((linea, tipo, numeros))
    
    if lineas_sin_match:
        print(f"\n{'=' * 80}")
        print(f"LÍNEAS SIN COINCIDENCIA ({len(lineas_sin_match)}):")
        print(f"{'=' * 80}")
        for linea, tipo, nums in lineas_sin_match[:15]:
            print(f"  - {linea[:70]}...")
            print(f"    Tipo: {tipo if tipo else '❌'} | IDs: {nums}")
    
    print(f"\n{'=' * 80}\n")

if __name__ == "__main__":
    archivo_txt = "01_Chat_WA.txt"
    archivo_excel = "tienda.xlsx"
    columna_id = "TIENDA ID_CUBO"
    archivo_salida = "coincidencias.xlsx"
    
    buscar_por_id(
        archivo_txt=archivo_txt,
        archivo_excel=archivo_excel,
        columna_id=columna_id,
        archivo_salida=archivo_salida
    )