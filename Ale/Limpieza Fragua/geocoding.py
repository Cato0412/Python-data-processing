import pandas as pd
import requests
import time
from typing import Tuple, Optional

# Configuración
API_KEY = ''  # Reemplaza con tu API key de Google
INPUT_FILE = r'C:\Users\lapmxdf558\Desktop\PYTHON\Limpieza Fragua\direcciones.xlsx'  # Ruta completa del archivo
OUTPUT_FILE = r'C:\Users\lapmxdf558\Desktop\PYTHON\Limpieza Fragua\direcciones_geocodificadas.xlsx'  # Ruta completa de salida
DELAY_SECONDS = 0.1  # Delay entre peticiones para no exceder límites

def geocode_address(address: str, api_key: str) -> Tuple[Optional[float], Optional[float], str]:
    """
    Geocodifica una dirección usando Google Geocoding API.
    
    Returns:
        Tuple con (latitud, longitud, status)
    """
    base_url = "https://maps.googleapis.com/maps/api/geocode/json"
    
    params = {
        'address': address,
        'key': api_key
    }
    
    try:
        response = requests.get(base_url, params=params)
        data = response.json()
        
        if data['status'] == 'OK':
            location = data['results'][0]['geometry']['location']
            return location['lat'], location['lng'], 'OK'
        else:
            return None, None, data['status']
            
    except Exception as e:
        return None, None, f'ERROR: {str(e)}'

def main():
    # Leer el archivo Excel
    print(f"Leyendo archivo {INPUT_FILE}...")
    df = pd.read_excel(INPUT_FILE)
    
    # Verificar que existe la columna 'direcciones'
    if 'direcciones' not in df.columns:
        print("Error: No se encontró la columna 'direcciones' en el archivo.")
        print(f"Columnas disponibles: {df.columns.tolist()}")
        return
    
    # Crear columnas para resultados
    df['latitud'] = None
    df['longitud'] = None
    df['geocoding_status'] = None
    
    total = len(df)
    print(f"\nIniciando geocodificación de {total} direcciones...")
    print("Esto puede tomar varios minutos...\n")
    
    # Procesar cada dirección
    for idx, row in df.iterrows():
        address = row['direcciones']
        
        # Saltar si la dirección está vacía
        if pd.isna(address) or str(address).strip() == '':
            df.at[idx, 'geocoding_status'] = 'DIRECCION_VACIA'
            continue
        
        # Geocodificar
        lat, lng, status = geocode_address(str(address), API_KEY)
        
        # Guardar resultados
        df.at[idx, 'latitud'] = lat
        df.at[idx, 'longitud'] = lng
        df.at[idx, 'geocoding_status'] = status
        
        # Mostrar progreso cada 100 direcciones
        if (idx + 1) % 100 == 0:
            print(f"Progreso: {idx + 1}/{total} direcciones procesadas ({((idx + 1)/total*100):.1f}%)")
        
        # Delay para respetar límites de la API
        time.sleep(DELAY_SECONDS)
    
    # Guardar resultado
    print(f"\nGuardando resultados en {OUTPUT_FILE}...")
    df.to_excel(OUTPUT_FILE, index=False)
    
    # Estadísticas
    successful = (df['geocoding_status'] == 'OK').sum()
    failed = total - successful
    
    print("\n" + "="*50)
    print("RESUMEN")
    print("="*50)
    print(f"Total de direcciones: {total}")
    print(f"Geocodificadas exitosamente: {successful}")
    print(f"Fallidas: {failed}")
    print(f"\nArchivo guardado: {OUTPUT_FILE}")
    print("="*50)

if __name__ == "__main__":

    main()