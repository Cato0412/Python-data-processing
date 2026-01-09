import os
import pywhatkit as kit
import pyautogui
import time
import pandas as pd

# ---- CONFIGURACIÓN ----
ARCHIVO_EXCEL = r"C:\Users\lapmxdf558\Downloads\Efectividad\Efectividad_Data_12 Nov.xlsx"
HOJA = "Sin CheckInOut"

# Leer el archivo Excel (la fila 0 tiene los encabezados reales)
print("Leyendo archivo Excel...")
df = pd.read_excel(ARCHIVO_EXCEL, sheet_name=HOJA, header=1)

# Mostrar las primeras filas para verificar
print("\nPrimeras filas del archivo:")
print(df.head())
print(f"\nColumnas encontradas: {list(df.columns)}")
print(f"\nTotal de registros: {len(df)}")

# Limpiar datos y crear diccionario de contactos
contactos = {}
for index, row in df.iterrows():
    # Usar 'Nombre Promotor' que es donde están los nombres
    if 'Nombre Promotor' in df.columns:
        nombre = str(row['Nombre Promotor']).strip()
        
        # Si el nombre es 'nan', saltar este registro
        if nombre == 'nan' or nombre == '':
            continue
        
        # Recolectar todos los números válidos de esta persona
        numeros_validos = []
        
        # Revisar las columnas de números disponibles
        columnas_numeros = []
        if 'Numeros' in df.columns:
            columnas_numeros.append('Numeros')
        if 'cel empresa' in df.columns:
            columnas_numeros.append('cel empresa')
        if 'Numero casa' in df.columns:
            columnas_numeros.append('Numero casa')
        
        for col in columnas_numeros:
            celular = str(row[col]).strip()
            
            # Validar que el número no esté vacío, sea NaN o sea 0
            if celular and celular != 'nan' and celular != '' and celular != '0':
                # Asegurarse de que el número tenga el formato correcto
                if not celular.startswith('+'):
                    celular = '+52' + celular.replace(' ', '').replace('-', '')
                
                # Evitar duplicados
                if celular not in numeros_validos:
                    numeros_validos.append(celular)
        
        # Si hay al menos un número válido, agregar al diccionario
        if numeros_validos:
            contactos[nombre] = numeros_validos
    else:
        print("Error: No se encontró la columna 'Nombre Promotor'")
        print(f"Columnas disponibles: {list(df.columns)}")
        break

print(f"\nContactos válidos encontrados: {len(contactos)}")
print("\nLista de contactos:")
for nombre, numeros in contactos.items():
    print(f"  - {nombre}:")
    for i, num in enumerate(numeros, 1):
        print(f"      {i}. {num}")


# Función para enviar mensaje
def enviar_whatsapp(numero, mensaje):
    try:
        kit.sendwhatmsg_instantly(numero, mensaje, wait_time=10, tab_close=False)
        time.sleep(7)
        pyautogui.press("enter")  # enviar el texto
        time.sleep(3)
    except Exception as e:
        print(f"Error enviando a {numero}: {e}")


# ---- MENSAJE DE SEGUIMIENTO SOBRE CAPACITACIÓN ----
def crear_mensaje(nombre):
    return f"""Buen día {nombre},

Nos hemos dado cuenta que no has estado registrando tu asistencia en Retail Optics. 

Por favor, podrías ayudarnos respondiendo las siguientes preguntas:

1. ¿Ya fuiste capacitado en el uso de la aplicación Retail Optics?

2. ¿Has tenido algún problema técnico con la aplicación? (no carga, no abre, errores, etc.)

3. ¿Cuál es el motivo principal por el que no estás capturando tu asistencia?

Es importante que nos compartas esta información para poder apoyarte y resolver cualquier inconveniente.

Por favor, reporta cualquier incidencia directamente a tu supervisor.

Saludos!"""


# ---- Envío de mensajes ----
print("\n--- INICIANDO ENVÍO DE MENSAJES DE SEGUIMIENTO ---\n")

respuesta = input("¿Deseas proceder con el envío de mensajes? (s/n): ")

if respuesta.lower() == 's':
    for nombre, numeros in contactos.items():
        mensaje = crear_mensaje(nombre)
        
        # Enviar a todos los números de esta persona
        for i, numero in enumerate(numeros, 1):
            print(f"\nEnviando mensaje a {nombre} - Número {i}/{len(numeros)} ({numero})...")
            enviar_whatsapp(numero, mensaje)
            time.sleep(5)  # Pausa entre mensajes para evitar bloqueos
            print(f"✓ Mensaje enviado a {nombre} ({numero})")
    
    print("\n--- ENVÍO COMPLETADO ---")
else:
    print("\nEnvío cancelado por el usuario.")