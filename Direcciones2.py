import pandas as pd
import requests
import time
from urllib.parse import quote

# ============= CONFIGURACIÓN =============
archivo_excel = 'Rut.xlsx'  # Cambia por tu archivo
API_KEY = 'AIzaSyAFwJ8jwqD8311nkfQ9yTWHcBhUp6Cs-xg'  # Pon tu API Key de Google Cloud aquí
DELAY_ENTRE_PETICIONES = 0.05  # Segundos entre peticiones (ajustar según tu cuota)
# =========================================

# Leer el archivo Excel
print("📂 Leyendo archivo Excel...")
df = pd.read_excel(archivo_excel)

print(f"✓ Archivo cargado: {len(df)} registros encontrados")
print(f"✓ Columnas disponibles: {list(df.columns)}\n")

# Función para crear la dirección completa desde los campos del Excel
def crear_direccion_completa(row):
    """
    Construye la dirección completa para geocodificar
    """
    partes = []
    
    # Agregar Calle y Número
    if pd.notna(row.get('Calle')):
        calle = str(row['Calle']).strip()
        if calle:
            partes.append(calle)
    
    if pd.notna(row.get('Numero')):
        numero = str(row['Numero']).strip()
        if numero and numero != '0':
            partes.append(numero)
    
    # Agregar Colonia
    if pd.notna(row.get('Colonia')):
        colonia = str(row['Colonia']).strip()
        if colonia:
            partes.append(colonia)
    
    # Agregar Municipio
    if pd.notna(row.get('Municipio')):
        municipio = str(row['Municipio']).strip()
        if municipio:
            partes.append(municipio)
    
    # Agregar Estado
    if pd.notna(row.get('Estado')):
        estado = str(row['Estado']).strip()
        if estado:
            partes.append(estado)
    
    # Agregar C.P.
    if pd.notna(row.get('C.P.')):
        cp = str(row['C.P.']).strip()
        if cp:
            partes.append(cp)
    
    # Agregar País
    if pd.notna(row.get('País')):
        pais = str(row['País']).strip()
        if pais:
            partes.append(pais)
    else:
        partes.append('México')  # Default
    
    return ', '.join(partes) if partes else ''

# Función para geocodificar usando Google Maps Geocoding API
def geocodificar_direccion(direccion, api_key):
    """
    Obtiene las coordenadas y dirección formateada de una dirección usando Google Geocoding API
    """
    if not direccion:
        return None, None, None, 'Sin dirección'
    
    url = 'https://maps.googleapis.com/maps/api/geocode/json'
    params = {
        'address': direccion,
        'key': api_key,
        'region': 'mx'  # Priorizar resultados de México
    }
    
    try:
        response = requests.get(url, params=params, timeout=10)
        data = response.json()
        
        if data['status'] == 'OK' and len(data['results']) > 0:
            location = data['results'][0]['geometry']['location']
            formatted_address = data['results'][0]['formatted_address']
            return location['lat'], location['lng'], formatted_address, 'OK'
        else:
            return None, None, None, data['status']
    except requests.exceptions.Timeout:
        return None, None, None, 'TIMEOUT'
    except Exception as e:
        return None, None, None, f'ERROR: {str(e)}'

# Crear dirección completa para cada registro
print("🔨 Construyendo direcciones...")
df['Direccion_Completa'] = df.apply(crear_direccion_completa, axis=1)

# Inicializar columnas para los resultados
df['Latitud_Google'] = None
df['Longitud_Google'] = None
df['Direccion_Formateada_Google'] = None
df['Status_Geocodificacion'] = None
df['Google_Maps_URL'] = None

print(f"\n🌍 Iniciando geocodificación con Google Maps API...")
print(f"Total de registros a procesar: {len(df)}\n")

# Contadores
exitosos = 0
fallidos = 0

# Geocodificar cada dirección
for idx, row in df.iterrows():
    direccion = row['Direccion_Completa']
    nombre_tienda = row.get('Nombre de Tienda', f'Registro {idx+1}')
    
    if not direccion:
        print(f"⚠️  [{idx+1}/{len(df)}] {nombre_tienda}: Sin dirección para geocodificar")
        df.at[idx, 'Status_Geocodificacion'] = 'SIN_DIRECCION'
        fallidos += 1
        continue
    
    # Geocodificar
    lat, lng, formatted_addr, status = geocodificar_direccion(direccion, API_KEY)
    
    # Guardar resultados
    df.at[idx, 'Latitud_Google'] = lat
    df.at[idx, 'Longitud_Google'] = lng
    df.at[idx, 'Direccion_Formateada_Google'] = formatted_addr
    df.at[idx, 'Status_Geocodificacion'] = status
    
    # Crear URL de Google Maps si se obtuvo coordenadas
    if lat and lng:
        df.at[idx, 'Google_Maps_URL'] = f"https://www.google.com/maps/search/?api=1&query={lat},{lng}"
        exitosos += 1
        print(f"✅ [{idx+1}/{len(df)}] {nombre_tienda}: Geocodificado exitosamente")
    else:
        fallidos += 1
        print(f"❌ [{idx+1}/{len(df)}] {nombre_tienda}: {status}")
    
    # Pausa entre peticiones para respetar límites de la API
    if idx < len(df) - 1:  # No hacer pausa en el último
        time.sleep(DELAY_ENTRE_PETICIONES)

# Guardar resultados en Excel con hipervínculos
print("\n💾 Guardando resultados...")
output_file = 'tiendas_geocodificadas.xlsx'

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='Tiendas')
    
    worksheet = writer.sheets['Tiendas']
    
    # Hacer clickeables los enlaces de Google Maps
    if 'Google_Maps_URL' in df.columns:
        url_col_idx = df.columns.get_loc('Google_Maps_URL') + 1
        
        for idx, url in enumerate(df['Google_Maps_URL'], start=2):
            if url and pd.notna(url):
                cell = worksheet.cell(row=idx, column=url_col_idx)
                cell.hyperlink = url
                cell.style = 'Hyperlink'

print(f"✅ Datos guardados en '{output_file}'")

# Guardar también en CSV
csv_file = 'tiendas_geocodificadas.csv'
df.to_csv(csv_file, index=False, encoding='utf-8-sig')
print(f"✅ Datos guardados en '{csv_file}'")

# Estadísticas finales
print(f"\n{'='*60}")
print(f"📊 RESUMEN DE GEOCODIFICACIÓN:")
print(f"{'='*60}")
print(f"Total de registros procesados: {len(df)}")
print(f"✅ Exitosos: {exitosos} ({exitosos/len(df)*100:.1f}%)")
print(f"❌ Fallidos: {fallidos} ({fallidos/len(df)*100:.1f}%)")

# Mostrar desglose de errores si hay fallidos
if fallidos > 0:
    print(f"\n📋 Desglose de errores:")
    status_counts = df['Status_Geocodificacion'].value_counts()
    for status, count in status_counts.items():
        if status != 'OK':
            print(f"   {status}: {count}")

print(f"\n✨ ¡Proceso completado!")