"""
Validador de Fotos Selfie en Excel
Lee un archivo Excel, busca la columna de fotos selfie y valida si hay personas en cada imagen
"""

import cv2
import pandas as pd
import requests
from io import BytesIO
import numpy as np
from datetime import datetime
import os
from urllib.parse import urlparse

class ValidadorFotosSelfie:
    def __init__(self):
        """Inicializa múltiples detectores para mayor precisión"""
        # Detector principal (rostros frontales)
        cascade_path = cv2.data.haarcascades + 'haarcascade_frontalface_default.xml'
        self.detector_frontal = cv2.CascadeClassifier(cascade_path)
        
        # Detector alternativo (más sensible)
        alt_cascade = cv2.data.haarcascades + 'haarcascade_frontalface_alt2.xml'
        self.detector_alt = cv2.CascadeClassifier(alt_cascade)
        
        # Detector de perfil
        profile_cascade = cv2.data.haarcascades + 'haarcascade_profileface.xml'
        self.detector_perfil = cv2.CascadeClassifier(profile_cascade)
    
    def descargar_imagen_desde_url(self, url):
        """Descarga una imagen desde una URL"""
        try:
            # Agregar timeout para evitar esperas largas
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            
            # Convertir a imagen OpenCV
            imagen_bytes = BytesIO(response.content)
            imagen_array = np.asarray(bytearray(imagen_bytes.read()), dtype=np.uint8)
            imagen = cv2.imdecode(imagen_array, cv2.IMREAD_COLOR)
            
            return imagen
        except Exception as e:
            return None
    
    def detectar_personas_en_imagen(self, imagen):
        """Detecta personas en una imagen usando múltiples métodos"""
        if imagen is None:
            return {"personas": 0, "detectado": False, "error": "Imagen no válida"}
        
        try:
            # Convertir a escala de grises
            gris = cv2.cvtColor(imagen, cv2.COLOR_BGR2GRAY)
            
            # Mejorar contraste para mejor detección
            gris = cv2.equalizeHist(gris)
            
            rostros_encontrados = set()
            
            # Método 1: Detector frontal estándar (más estricto)
            rostros1 = self.detector_frontal.detectMultiScale(
                gris,
                scaleFactor=1.05,  # Más sensible (antes 1.1)
                minNeighbors=3,     # Menos estricto (antes 5)
                minSize=(20, 20),   # Acepta rostros más pequeños
                flags=cv2.CASCADE_SCALE_IMAGE
            )
            for (x, y, w, h) in rostros1:
                rostros_encontrados.add((x, y, w, h))
            
            # Método 2: Detector alternativo (más sensible)
            rostros2 = self.detector_alt.detectMultiScale(
                gris,
                scaleFactor=1.05,
                minNeighbors=2,     # Muy sensible
                minSize=(20, 20),
                flags=cv2.CASCADE_SCALE_IMAGE
            )
            for (x, y, w, h) in rostros2:
                # Evitar duplicados cercanos
                if not self._es_duplicado((x, y, w, h), rostros_encontrados):
                    rostros_encontrados.add((x, y, w, h))
            
            # Método 3: Detector de perfil (izquierda)
            rostros3 = self.detector_perfil.detectMultiScale(
                gris,
                scaleFactor=1.05,
                minNeighbors=3,
                minSize=(20, 20)
            )
            for (x, y, w, h) in rostros3:
                if not self._es_duplicado((x, y, w, h), rostros_encontrados):
                    rostros_encontrados.add((x, y, w, h))
            
            # Método 4: Probar con imagen volteada para perfil derecho
            gris_flip = cv2.flip(gris, 1)
            rostros4 = self.detector_perfil.detectMultiScale(
                gris_flip,
                scaleFactor=1.05,
                minNeighbors=3,
                minSize=(20, 20)
            )
            ancho_img = gris.shape[1]
            for (x, y, w, h) in rostros4:
                # Convertir coordenadas de vuelta
                x_real = ancho_img - x - w
                if not self._es_duplicado((x_real, y, w, h), rostros_encontrados):
                    rostros_encontrados.add((x_real, y, w, h))
            
            num_personas = len(rostros_encontrados)
            
            return {
                "personas": num_personas,
                "detectado": num_personas > 0,
                "error": None,
                "rostros": list(rostros_encontrados)
            }
        except Exception as e:
            return {"personas": 0, "detectado": False, "error": str(e)}
    
    def _es_duplicado(self, nuevo_rostro, rostros_existentes, umbral=0.3):
        """Verifica si un rostro ya fue detectado (evita duplicados)"""
        x1, y1, w1, h1 = nuevo_rostro
        
        for (x2, y2, w2, h2) in rostros_existentes:
            # Calcular superposición
            x_izq = max(x1, x2)
            y_arr = max(y1, y2)
            x_der = min(x1 + w1, x2 + w2)
            y_aba = min(y1 + h1, y2 + h2)
            
            if x_der > x_izq and y_aba > y_arr:
                area_interseccion = (x_der - x_izq) * (y_aba - y_arr)
                area1 = w1 * h1
                area2 = w2 * h2
                area_union = area1 + area2 - area_interseccion
                
                iou = area_interseccion / area_union if area_union > 0 else 0
                
                if iou > umbral:
                    return True
        
        return False
    
    def procesar_excel(self, ruta_excel, nombre_columna="Foto Selfie (Obligatoria)"):
        """
        Procesa el archivo Excel y valida todas las fotos
        
        Args:
            ruta_excel: Ruta al archivo Excel
            nombre_columna: Nombre de la columna que contiene las URLs de fotos
        """
        print(f"\n{'='*80}")
        print(f"VALIDADOR DE FOTOS SELFIE - ANÁLISIS DE EXCEL")
        print(f"{'='*80}\n")
        
        try:
            # Leer el archivo Excel
            print(f"📂 Leyendo archivo: {ruta_excel}")
            df = pd.read_excel(ruta_excel)
            print(f"✅ Archivo cargado: {len(df)} filas encontradas\n")
            
            # Buscar la columna de fotos selfie
            if nombre_columna not in df.columns:
                print(f"❌ Error: No se encontró la columna '{nombre_columna}'")
                print(f"\nColumnas disponibles en el archivo:")
                for i, col in enumerate(df.columns, 1):
                    print(f"  {i}. {col}")
                return None
            
            print(f"✅ Columna '{nombre_columna}' encontrada\n")
            print(f"{'='*80}")
            print(f"INICIANDO ANÁLISIS DE FOTOS")
            print(f"{'='*80}\n")
            
            # Crear columnas para resultados
            df['Personas_Detectadas'] = 0
            df['Tiene_Personas'] = 'No validado'
            df['Estado_Validacion'] = 'Pendiente'
            df['Fecha_Validacion'] = ''
            df['Metodo_Deteccion'] = ''
            
            resultados_resumen = {
                'total_filas': len(df),
                'con_url': 0,
                'sin_url': 0,
                'validadas_ok': 0,
                'sin_personas': 0,
                'errores': 0
            }
            
            # Procesar cada fila
            for idx, row in df.iterrows():
                fila_num = idx + 2  # +2 porque Excel empieza en 1 y tiene encabezado
                url = row[nombre_columna]
                
                print(f"Fila {fila_num}: ", end="")
                
                # Verificar si hay URL
                if pd.isna(url) or str(url).strip() == '':
                    print("⚠️  Sin foto (celda vacía)")
                    df.at[idx, 'Estado_Validacion'] = 'Sin foto'
                    df.at[idx, 'Tiene_Personas'] = 'N/A'
                    resultados_resumen['sin_url'] += 1
                    continue
                
                resultados_resumen['con_url'] += 1
                
                # Descargar imagen
                print(f"⬇️  Descargando... ", end="")
                imagen = self.descargar_imagen_desde_url(str(url))
                
                if imagen is None:
                    print("❌ Error al descargar")
                    df.at[idx, 'Estado_Validacion'] = 'Error de descarga'
                    df.at[idx, 'Tiene_Personas'] = 'Error'
                    resultados_resumen['errores'] += 1
                    continue
                
                # Detectar personas
                resultado = self.detectar_personas_en_imagen(imagen)
                
                if resultado['error']:
                    print(f"❌ Error: {resultado['error']}")
                    df.at[idx, 'Estado_Validacion'] = f"Error: {resultado['error']}"
                    df.at[idx, 'Tiene_Personas'] = 'Error'
                    resultados_resumen['errores'] += 1
                else:
                    num_personas = resultado['personas']
                    df.at[idx, 'Personas_Detectadas'] = num_personas
                    df.at[idx, 'Fecha_Validacion'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    df.at[idx, 'Metodo_Deteccion'] = 'Múltiple (mejorado)'
                    
                    if num_personas > 0:
                        print(f"✅ {num_personas} persona(s) detectada(s)")
                        df.at[idx, 'Tiene_Personas'] = 'Sí'
                        df.at[idx, 'Estado_Validacion'] = 'Válida'
                        resultados_resumen['validadas_ok'] += 1
                    else:
                        print(f"⚠️  NO se detectaron personas")
                        df.at[idx, 'Tiene_Personas'] = 'No'
                        df.at[idx, 'Estado_Validacion'] = 'Sin personas'
                        resultados_resumen['sin_personas'] += 1
            
            # Guardar resultados
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nombre_base = os.path.splitext(os.path.basename(ruta_excel))[0]
            carpeta = os.path.dirname(ruta_excel)
            ruta_resultado = os.path.join(carpeta, f"{nombre_base}_validado_{timestamp}.xlsx")
            
            df.to_excel(ruta_resultado, index=False)
            
            # Mostrar resumen
            self.mostrar_resumen(resultados_resumen, ruta_resultado)
            
            return df
            
        except Exception as e:
            print(f"\n❌ Error al procesar el archivo Excel: {str(e)}")
            return None
    
    def mostrar_resumen(self, resultados, ruta_resultado):
        """Muestra un resumen de los resultados"""
        print(f"\n{'='*80}")
        print(f"RESUMEN DE VALIDACIÓN")
        print(f"{'='*80}\n")
        
        print(f"📊 Total de filas procesadas: {resultados['total_filas']}")
        print(f"\n📸 Fotos encontradas:")
        print(f"   • Con URL de foto: {resultados['con_url']}")
        print(f"   • Sin URL (vacías): {resultados['sin_url']}")
        
        print(f"\n✅ Resultados de validación:")
        print(f"   • Fotos con personas detectadas: {resultados['validadas_ok']}")
        print(f"   • Fotos sin personas: {resultados['sin_personas']}")
        print(f"   • Errores al procesar: {resultados['errores']}")
        
        if resultados['con_url'] > 0:
            porcentaje_validas = (resultados['validadas_ok'] / resultados['con_url']) * 100
            print(f"\n📈 Porcentaje de fotos válidas: {porcentaje_validas:.1f}%")
        
        print(f"\n💾 Archivo de resultados guardado en:")
        print(f"   {ruta_resultado}")
        
        print(f"\n{'='*80}\n")

def main():
    """Función principal"""
    # Ruta fija del archivo Excel
    ruta_excel = r"C:\Users\lapmxdf558\Documents\Archivos Alejandro\Genomma Mayoreo\Rutero Mayoreo\Rutero Enero\Asistencia"
    
    print("="*80)
    print("VALIDADOR AUTOMÁTICO DE FOTOS SELFIE")
    print("="*80)
    
    # Buscar archivos Excel en la carpeta
    print(f"\n📂 Buscando archivos Excel en: {ruta_excel}\n")
    
    if not os.path.exists(ruta_excel):
        print(f"❌ Error: La carpeta no existe")
        return
    
    archivos_excel = [f for f in os.listdir(ruta_excel) if f.endswith(('.xlsx', '.xls')) and not f.startswith('~')]
    
    if not archivos_excel:
        print("❌ No se encontraron archivos Excel en la carpeta")
        return
    
    print(f"Archivos Excel encontrados:")
    for i, archivo in enumerate(archivos_excel, 1):
        print(f"  {i}. {archivo}")
    
    # Seleccionar archivo
    if len(archivos_excel) == 1:
        archivo_seleccionado = archivos_excel[0]
        print(f"\n✅ Procesando único archivo: {archivo_seleccionado}")
    else:
        try:
            seleccion = int(input(f"\nSelecciona el número de archivo (1-{len(archivos_excel)}): "))
            if 1 <= seleccion <= len(archivos_excel):
                archivo_seleccionado = archivos_excel[seleccion - 1]
            else:
                print("❌ Selección inválida")
                return
        except ValueError:
            print("❌ Entrada inválida")
            return
    
    ruta_completa = os.path.join(ruta_excel, archivo_seleccionado)
    
    # Crear validador y procesar
    validador = ValidadorFotosSelfie()
    validador.procesar_excel(ruta_completa)

if __name__ == "__main__":
    main()