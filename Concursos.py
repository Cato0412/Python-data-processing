import pandas as pd
import openpyxl
import re

def normalizar_texto(texto):
    """Normaliza texto para comparación: minúsculas, sin acentos, sin espacios extra"""
    if pd.isna(texto):
        return ""
    texto = str(texto).lower().strip()
    texto = re.sub(r'\s+', ' ', texto)
    replacements = {'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ú': 'u', 'ñ': 'n'}
    for old, new in replacements.items():
        texto = texto.replace(old, new)
    return texto

def buscar_coincidencias(archivo_txt, archivo_excel, columna_busqueda, 
                        archivo_salida='coincidencias.xlsx', 
                        tipo_busqueda='parcial'):
    """
    Busca coincidencias entre un archivo TXT y una columna de Excel.
    """
    
    print("=" * 70)
    print("BUSCADOR DE COINCIDENCIAS TXT-EXCEL")
    print("=" * 70)
    
    # Leer el archivo TXT
    try:
        with open(archivo_txt, 'r', encoding='utf-8') as f:
            valores_txt = [linea.strip() for linea in f if linea.strip()]
        print(f"\n✓ Archivo TXT leído: {len(valores_txt)} valores encontrados")
        print(f"  Primeros valores:")
        for v in valores_txt[:3]:
            print(f"    - {v[:80]}")
    except Exception as e:
        print(f"\n✗ Error al leer el archivo TXT: {e}")
        return
    
    # Leer el archivo Excel
    try:
        df = pd.read_excel(archivo_excel)
        print(f"\n✓ Archivo Excel leído: {len(df)} filas, {len(df.columns)} columnas")
        print(f"  Columnas: {list(df.columns)}")
    except Exception as e:
        print(f"\n✗ Error al leer el archivo Excel: {e}")
        return
    
    # Verificar que la columna existe
    if columna_busqueda not in df.columns:
        print(f"\n✗ Error: La columna '{columna_busqueda}' no existe")
        return
    
    print(f"\n  Primeros valores en '{columna_busqueda}':")
    for v in df[columna_busqueda].dropna().head(3):
        print(f"    - {str(v)[:80]}")
    
    print(f"\n{'=' * 70}")
    print(f"Tipo de búsqueda: {tipo_busqueda.upper()}")
    print(f"{'=' * 70}")
    
    df_resultado = pd.DataFrame()
    coincidencias_detalle = []
    valores_encontrados_set = set()
    
    if tipo_busqueda == 'exacta':
        df[columna_busqueda] = df[columna_busqueda].astype(str)
        valores_txt_str = [str(v) for v in valores_txt]
        df_resultado = df[df[columna_busqueda].isin(valores_txt_str)]
        valores_encontrados_set = set(df_resultado[columna_busqueda].unique())
        
    else:  # Búsqueda parcial
        df['_temp_norm'] = df[columna_busqueda].apply(normalizar_texto)
        
        print("\nBuscando coincidencias...")
        for i, valor_txt in enumerate(valores_txt):
            if (i + 1) % 50 == 0:
                print(f"  Procesados: {i + 1}/{len(valores_txt)}")
            
            valor_norm = normalizar_texto(valor_txt)
            
            # Buscar coincidencias
            matches = df[df['_temp_norm'].str.contains(valor_norm, na=False, regex=False)]
            
            if len(matches) > 0:
                df_resultado = pd.concat([df_resultado, matches], ignore_index=True)
                valores_encontrados_set.add(valor_txt)
                
                for _, row in matches.iterrows():
                    coincidencias_detalle.append({
                        'Valor_TXT': valor_txt,
                        'Valor_Excel': row[columna_busqueda]
                    })
        
        if '_temp_norm' in df_resultado.columns:
            df_resultado = df_resultado.drop(columns=['_temp_norm'])
        
        df_resultado = df_resultado.drop_duplicates()
    
    print(f"\n{'=' * 70}")
    print(f"RESULTADOS")
    print(f"{'=' * 70}")
    print(f"✓ Filas con coincidencias: {len(df_resultado)}")
    print(f"✓ Valores del TXT encontrados: {len(valores_encontrados_set)}/{len(valores_txt)}")
    print(f"✓ Porcentaje de éxito: {len(valores_encontrados_set)/len(valores_txt)*100:.2f}%")
    
    if len(df_resultado) > 0:
        df_resultado.to_excel(archivo_salida, index=False, engine='openpyxl')
        print(f"\n✓ Archivo creado: '{archivo_salida}'")
        
        print(f"\n{'=' * 70}")
        print("PREVIEW:")
        print(f"{'=' * 70}")
        print(df_resultado.head(5).to_string(index=False))
        
        if tipo_busqueda == 'parcial' and coincidencias_detalle:
            print(f"\n{'=' * 70}")
            print("EJEMPLOS DE COINCIDENCIAS:")
            print(f"{'=' * 70}")
            for det in coincidencias_detalle[:5]:
                print(f"\n  TXT: {det['Valor_TXT'][:70]}")
                print(f"   ↓")
                print(f"  Excel: {det['Valor_Excel']}")
    else:
        print("\n⚠ No se encontraron coincidencias")
        print("\nSugerencias:")
        print("  1. Verifica que los valores del TXT coincidan con el Excel")
        print("  2. Prueba con tipo_busqueda='exacta' si son valores cortos")
        print("  3. Revisa si hay caracteres especiales o formato diferente")
    
    valores_no_encontrados = [v for v in valores_txt if v not in valores_encontrados_set]
    
    if valores_no_encontrados:
        print(f"\n{'=' * 70}")
        print(f"SIN COINCIDENCIA ({len(valores_no_encontrados)}):")
        print(f"{'=' * 70}")
        for v in valores_no_encontrados[:5]:
            print(f"  - {v[:70]}")
        if len(valores_no_encontrados) > 5:
            print(f"  ... y {len(valores_no_encontrados) - 5} más")
    
    print(f"\n{'=' * 70}\n")

# ============================================================================
# CONFIGURACIÓN
# ============================================================================

if __name__ == "__main__":
    archivo_txt = "dump_total.txt"
    archivo_excel = "tienda.xlsx"
    columna = "Nombre de Tienda"
    archivo_salida = "coincidencias.xlsx"
    
    # CAMBIA AQUÍ EL TIPO DE BÚSQUEDA:
    # 'parcial' = busca si el texto del TXT está contenido en el Excel (flexible)
    # 'exacta' = el valor debe ser idéntico
    tipo = 'parcial'
    
    buscar_coincidencias(
        archivo_txt=archivo_txt,
        archivo_excel=archivo_excel,
        columna_busqueda=columna,
        archivo_salida=archivo_salida,
        tipo_busqueda=tipo
    )