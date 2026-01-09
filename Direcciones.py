import pandas as pd
import requests
import time
from datetime import datetime

# ============= CONFIGURACIÓN =============
GOOGLE_API_KEY = "AIzaSyAFwJ8jwqD8311nkfQ9yTWHcBhUp6Cs-xg"  # Coloca aquí tu API key de Google Cloud
INPUT_FILE = "Rut.xlsx"      # Nombre de tu archivo Excel de entrada
OUTPUT_FILE = "Rut_direcciones_completas.xlsx"  # Archivo de salida

# ============= FUNCIONES =============

def extract_address_components(address_components):
    """Extrae componentes de dirección de forma más completa"""
    components = {
        "street": "",
        "number": "",
        "neighborhood": "",
        "city": "",
        "municipality": "",
        "state": "",
        "postal_code": "",
        "country": ""
    }
    
    for component in address_components:
        types = component["types"]
        long_name = component.get("long_name", "")
        
        # Calle
        if "route" in types and not components["street"]:
            components["street"] = long_name
        # Número
        elif "street_number" in types and not components["number"]:
            components["number"] = long_name
        # Colonia/Barrio
        elif ("sublocality_level_1" in types or "sublocality" in types or 
              "neighborhood" in types) and not components["neighborhood"]:
            components["neighborhood"] = long_name
        # Ciudad
        elif "locality" in types and not components["city"]:
            components["city"] = long_name
        # Municipio (alternativa a ciudad)
        elif "administrative_area_level_2" in types and not components["municipality"]:
            components["municipality"] = long_name
        # Estado
        elif "administrative_area_level_1" in types and not components["state"]:
            components["state"] = long_name
        # Código Postal
        elif "postal_code" in types and not components["postal_code"]:
            components["postal_code"] = long_name
        # País
        elif "country" in types and not components["country"]:
            components["country"] = long_name
    
    # Si no hay ciudad pero hay municipio, usar municipio
    if not components["city"] and components["municipality"]:
        components["city"] = components["municipality"]
    
    return components


def parse_formatted_address(formatted_address):
    """
    Intenta extraer componentes adicionales del formatted_address
    cuando los address_components no son suficientes
    """
    result = {
        "calle_numero": "",
        "colonia": "",
        "ciudad_cp": ""
    }
    
    # El formatted_address suele venir como:
    # "Calle Número, Colonia, CP Ciudad, Estado, País"
    parts = [p.strip() for p in formatted_address.split(',')]
    
    if len(parts) >= 1:
        result["calle_numero"] = parts[0]
    if len(parts) >= 2:
        result["colonia"] = parts[1]
    if len(parts) >= 3:
        result["ciudad_cp"] = parts[2]
    
    return result


def geocode_place(nombre_tienda, determinante, pais, api_key):
    """
    Geocodifica un lugar usando Google Geocoding API
    """
    # Construir query de búsqueda
    query = f"{nombre_tienda} {determinante}, {pais}"
    
    url = "https://maps.googleapis.com/maps/api/geocode/json"
    params = {
        "address": query,
        "key": api_key,
        "language": "es"  # Forzar respuestas en español
    }
    
    try:
        response = requests.get(url, params=params, timeout=10)
        response.raise_for_status()
        data = response.json()
        
        if data["status"] == "OK" and len(data["results"]) > 0:
            result = data["results"][0]
            location = result["geometry"]["location"]
            formatted_address = result.get("formatted_address", "")
            
            # Extraer componentes estructurados
            components = extract_address_components(result.get("address_components", []))
            
            # Si falta información, intentar parsear el formatted_address
            parsed = parse_formatted_address(formatted_address)
            
            # Combinar información
            street = components["street"] or ""
            number = components["number"] or ""
            
            # Si no hay calle/número separados, intentar extraer del formatted_address
            if not street and parsed["calle_numero"]:
                # Intentar separar calle y número del primer componente
                calle_numero_parts = parsed["calle_numero"].split()
                if len(calle_numero_parts) > 1 and calle_numero_parts[-1].isdigit():
                    number = calle_numero_parts[-1]
                    street = " ".join(calle_numero_parts[:-1])
                else:
                    street = parsed["calle_numero"]
            
            neighborhood = components["neighborhood"] or parsed.get("colonia", "")
            city = components["city"] or ""
            state = components["state"] or ""
            postal_code = components["postal_code"] or ""
            
            return {
                "status": "SUCCESS",
                "direccion_completa": formatted_address,
                "calle": street,
                "numero": number,
                "colonia": neighborhood,
                "ciudad": city,
                "municipio": components["municipality"] or city,
                "estado": state,
                "codigo_postal": postal_code,
                "pais": components["country"],
                "latitud": location["lat"],
                "longitud": location["lng"],
                "place_id": result.get("place_id", ""),
                "error": None
            }
        else:
            error_msg = data.get("status", "No encontrada")
            if data.get("error_message"):
                error_msg += f": {data['error_message']}"
            
            return {
                "status": "NOT_FOUND",
                "direccion_completa": None,
                "calle": None,
                "numero": None,
                "colonia": None,
                "ciudad": None,
                "municipio": None,
                "estado": None,
                "codigo_postal": None,
                "pais": None,
                "latitud": None,
                "longitud": None,
                "place_id": None,
                "error": error_msg
            }
            
    except requests.exceptions.RequestException as e:
        return {
            "status": "ERROR",
            "direccion_completa": None,
            "calle": None,
            "numero": None,
            "colonia": None,
            "ciudad": None,
            "municipio": None,
            "estado": None,
            "codigo_postal": None,
            "pais": None,
            "latitud": None,
            "longitud": None,
            "place_id": None,
            "error": str(e)
        }


def process_excel(input_file, output_file, api_key, delay=0.15):
    """
    Procesa el archivo Excel y geocodifica todas las direcciones
    """
    print(f"📂 Leyendo archivo: {input_file}")
    
    # Leer Excel
    try:
        df = pd.read_excel(input_file)
    except FileNotFoundError:
        print(f"❌ Error: No se encontró el archivo '{input_file}'")
        print(f"   Asegúrate de que el archivo esté en la misma carpeta que el script")
        return
    
    # Verificar que existan las columnas necesarias
    required_columns = ["Nombre de Tienda", "Determinante", "País"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        print(f"❌ Error: Faltan columnas en el Excel: {missing_columns}")
        print(f"   Columnas encontradas: {list(df.columns)}")
        return
    
    total = len(df)
    print(f"✓ Archivo cargado: {total} registros")
    print(f"   Columnas: {list(df.columns)}")
    print(f"\n🚀 Iniciando geocodificación...")
    print(f"   Delay entre peticiones: {delay} segundos")
    print("-" * 75)
    
    # Inicializar/actualizar columnas
    column_mapping = {
        "Calle": "calle",
        "Numero": "numero", 
        "Latitud": "latitud",
        "Longitud": "longitud",
        "Municipio": "municipio",
        "Colonia": "colonia",
        "C.P.": "codigo_postal",
        "Estado": "estado"
    }
    
    # Columnas adicionales
    if "direccion_completa" not in df.columns:
        df["direccion_completa"] = None
    if "status_geocoding" not in df.columns:
        df["status_geocoding"] = None
    if "error_geocoding" not in df.columns:
        df["error_geocoding"] = None
    if "place_id" not in df.columns:
        df["place_id"] = None
    
    success_count = 0
    failed_count = 0
    partial_count = 0  # Cuando tiene coordenadas pero faltan datos
    start_time = time.time()
    
    # Procesar cada fila
    for idx, row in df.iterrows():
        # Extraer datos
        nombre_tienda = str(row.get("Nombre de Tienda", "")).strip()
        determinante = str(row.get("Determinante", "")).strip()
        pais = str(row.get("País", "")).strip()
        
        if not nombre_tienda or not pais:
            print(f"[{idx+1}/{total}] ⚠️  Registro {idx+1}: Datos incompletos")
            df.at[idx, "status_geocoding"] = "INCOMPLETE"
            df.at[idx, "error_geocoding"] = "Nombre de Tienda o País vacío"
            failed_count += 1
            continue
        
        # Geocodificar
        result = geocode_place(nombre_tienda, determinante, pais, api_key)
        
        # Guardar resultados
        df.at[idx, "status_geocoding"] = result["status"]
        df.at[idx, "direccion_completa"] = result["direccion_completa"]
        df.at[idx, "place_id"] = result["place_id"]
        df.at[idx, "error_geocoding"] = result["error"]
        
        # Actualizar columnas existentes
        for excel_col, result_key in column_mapping.items():
            if excel_col in df.columns:
                df.at[idx, excel_col] = result[result_key]
        
        # Mostrar progreso
        if result["status"] == "SUCCESS":
            # Contar cuántos campos están completos
            campos_completos = sum([
                bool(result["calle"]),
                bool(result["numero"]),
                bool(result["colonia"]),
                bool(result["codigo_postal"])
            ])
            
            if campos_completos >= 3:
                success_count += 1
                status_icon = "✓"
            else:
                partial_count += 1
                status_icon = "◐"  # Parcial
            
            ciudad = result["ciudad"] or result["municipio"] or "N/A"
            info = f"{ciudad}"
            if result["calle"]:
                info = f"{result['calle']} {result['numero']}, {ciudad}"
            
            print(f"[{idx+1}/{total}] {status_icon} {nombre_tienda[:30]:<30} → {info[:40]}")
        else:
            failed_count += 1
            error = result["error"] or "No encontrada"
            print(f"[{idx+1}/{total}] ✗ {nombre_tienda[:30]:<30} → {error[:40]}")
        
        # Guardar progreso cada 50 registros
        if (idx + 1) % 50 == 0:
            df.to_excel(output_file, index=False)
            elapsed = time.time() - start_time
            avg_time = elapsed / (idx + 1)
            remaining = (total - idx - 1) * avg_time
            print(f"\n💾 Progreso guardado: {idx+1}/{total} ({(idx+1)/total*100:.1f}%)")
            print(f"   ✓ Completos: {success_count} | ◐ Parciales: {partial_count} | ✗ Fallidos: {failed_count}")
            print(f"   ⏱️  Tiempo: {elapsed/60:.1f} min | ⏳ Restante: {remaining/60:.1f} min\n")
        
        # Delay entre peticiones
        time.sleep(delay)
    
    # Guardar resultado final
    df.to_excel(output_file, index=False)
    
    # Resumen final
    elapsed = time.time() - start_time
    print("\n" + "=" * 75)
    print("✅ PROCESO COMPLETADO")
    print("=" * 75)
    print(f"Total procesados: {total}")
    print(f"✓ Completos: {success_count} ({success_count/total*100:.1f}%)")
    print(f"◐ Parciales (solo coordenadas): {partial_count} ({partial_count/total*100:.1f}%)")
    print(f"✗ Fallidos: {failed_count} ({failed_count/total*100:.1f}%)")
    print(f"Tiempo total: {elapsed/60:.1f} minutos")
    print(f"Tiempo promedio: {elapsed/total:.2f} seg/dirección")
    print(f"\n📁 Archivo guardado: {output_file}")
    print("\n💡 CONSEJOS:")
    print("   • Revisa 'status_geocoding' = 'NOT_FOUND' para direcciones no encontradas")
    print("   • Las direcciones parciales tienen coordenadas pero datos incompletos")
    print("   • Usa 'place_id' para búsquedas más precisas en Google Maps")


# ============= EJECUCIÓN PRINCIPAL =============

if __name__ == "__main__":
    print("=" * 75)
    print("        GEOCODIFICADOR MASIVO CON GOOGLE CLOUD API (MEJORADO)")
    print("=" * 75)
    print(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    
    # Verificar que se haya configurado la API key
    if GOOGLE_API_KEY == "TU_API_KEY_AQUI":
        print("❌ ERROR: Debes configurar tu GOOGLE_API_KEY en el script")
        print("   Edita la línea 8 y coloca tu API key de Google Cloud")
        print("\n📖 Para obtener tu API Key:")
        print("   1. Ve a: https://console.cloud.google.com/")
        print("   2. Crea un proyecto (si no tienes)")
        print("   3. Habilita 'Geocoding API'")
        print("   4. Ve a 'Credenciales' → Crear credenciales → API Key")
        exit(1)
    
    # Procesar archivo
    process_excel(INPUT_FILE, OUTPUT_FILE, GOOGLE_API_KEY, delay=0.15)