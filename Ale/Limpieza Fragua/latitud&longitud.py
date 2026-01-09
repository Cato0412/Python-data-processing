import pandas as pd
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut, GeocoderServiceError
import time
from datetime import datetime

# Configuración
INPUT_FILE = 'Plantilla.xlsx'  # Nombre de tu archivo
OUTPUT_FILE = 'direcciones_geocodificadas.xlsx'

# Inicializar geocodificador (Nominatim requiere un user_agent único)
geolocator = Nominatim(user_agent="mi_app_geocoding_2024", timeout=10)

def construir_direccion(row):
    """Construye la dirección completa a partir de las columnas"""
    componentes = []
    
    if pd.notna(row.get('CALLE')):
        componentes.append(str(row['CALLE']))
    
    if pd.notna(row.get('NUMERO')):
        componentes.append(str(row['NUMERO']))
    
    if pd.notna(row.get('COLONIA')):
        componentes.append(str(row['COLONIA']))
    
    if pd.notna(row.get('CODIGO_POSTAL')):
        componentes.append(str(row['CODIGO_POSTAL']))
    
    if pd.notna(row.get('MUNICIPIO')):
        componentes.append(str(row['MUNICIPIO']))
    
    if pd.notna(row.get('ESTADO')):
        componentes.append(str(row['ESTADO']))
    
    componentes.append("México")
    
    return ', '.join(componentes)

def geocodificar_direccion(direccion, intento=1, max_intentos=3):
    """Obtiene lat/lng de una dirección con reintentos"""
    try:
        location = geolocator.geocode(direccion, country_codes='mx')
        
        if location:
            return {
                'latitud': location.latitude,
                'longitud': location.longitude,
                'direccion_formateada': location.address,
                'estado': 'exitoso'
            }
        else:
            return {
                'latitud': None,
                'longitud': None,
                'direccion_formateada': None,
                'estado': 'no_encontrado'
            }
    except (GeocoderTimedOut, GeocoderServiceError) as e:
        if intento < max_intentos:
            print(f"  Error de conexión, reintentando ({intento}/{max_intentos})...")
            time.sleep(3)
            return geocodificar_direccion(direccion, intento + 1, max_intentos)
        else:
            return {
                'latitud': None,
                'longitud': None,
                'direccion_formateada': None,
                'estado': f'error: {str(e)}'
            }
    except Exception as e:
        return {
            'latitud': None,
            'longitud': None,
            'direccion_formateada': None,
            'estado': f'error: {str(e)}'
        }

def procesar_excel():
    """Procesa el archivo Excel completo"""
    print(f"Leyendo archivo {INPUT_FILE}...")
    df = pd.read_excel(INPUT_FILE)
    
    print(f"Total de registros: {len(df)}")
    print("\nColumnas encontradas:", df.columns.tolist())
    
    # Crear nuevas columnas
    df['DIRECCION_COMPLETA'] = ''
    df['LATITUD'] = None
    df['LONGITUD'] = None
    df['DIRECCION_OSM'] = ''
    df['ESTADO_GEOCODIFICACION'] = ''
    
    # Contadores
    exitosos = 0
    fallidos = 0
    inicio = datetime.now()
    
    print("\nIniciando geocodificación...")
    print("⚠️ Nota: Nominatim tiene límite de 1 request/segundo")
    print("   Con 2777 direcciones tomará aprox. 46 minutos\n")
    
    for idx, row in df.iterrows():
        # Construir dirección
        direccion = construir_direccion(row)
        df.at[idx, 'DIRECCION_COMPLETA'] = direccion
        
        # Geocodificar
        nombre = row.get('NOMBRE_SUCURSAL', 'Sin nombre')
        print(f"[{idx + 1}/{len(df)}] {nombre[:40]}...")
        resultado = geocodificar_direccion(direccion)
        
        # Guardar resultados
        df.at[idx, 'LATITUD'] = resultado['latitud']
        df.at[idx, 'LONGITUD'] = resultado['longitud']
        df.at[idx, 'DIRECCION_OSM'] = resultado['direccion_formateada']
        df.at[idx, 'ESTADO_GEOCODIFICACION'] = resultado['estado']
        
        if resultado['estado'] == 'exitoso':
            exitosos += 1
        else:
            fallidos += 1
            print(f"  ⚠️ No se pudo geocodificar")
        
        # Pausa obligatoria de 1 segundo (política de uso de Nominatim)
        time.sleep(1)
        
        # Guardar progreso cada 50 registros
        if (idx + 1) % 50 == 0:
            df.to_excel(OUTPUT_FILE, index=False)
            tiempo_transcurrido = datetime.now() - inicio
            estimado_restante = (tiempo_transcurrido / (idx + 1)) * (len(df) - idx - 1)
            print(f"\n💾 Progreso guardado: {idx + 1}/{len(df)} registros")
            print(f"   Tiempo transcurrido: {tiempo_transcurrido}")
            print(f"   Tiempo estimado restante: {estimado_restante}\n")
    
    # Guardar archivo final
    df.to_excel(OUTPUT_FILE, index=False)
    
    # Resumen
    duracion = datetime.now() - inicio
    print("\n" + "="*60)
    print("RESUMEN")
    print("="*60)
    print(f"✅ Exitosos: {exitosos}")
    print(f"❌ Fallidos: {fallidos}")
    print(f"📊 Tasa de éxito: {(exitosos/len(df)*100):.1f}%")
    print(f"⏱️ Tiempo total: {duracion}")
    print(f"📁 Archivo guardado: {OUTPUT_FILE}")
    print("="*60)

if __name__ == "__main__":
    try:
        procesar_excel()
    except FileNotFoundError:
        print(f"❌ Error: No se encontró el archivo '{INPUT_FILE}'")
    except Exception as e:
        print(f"❌ Error inesperado: {str(e)}")
        import traceback
        traceback.print_exc()