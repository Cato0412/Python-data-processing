import os
import pywhatkit as kit
import pyautogui
import time
import pandas as pd

# ---- CONFIGURACIÓN ----
ARCHIVO_EXCEL = r"C:\Users\lapmxdf558\Downloads\Efectividad\llamadas.xlsx"
HOJA = "llamadas"  # Ajusta el nombre de la hoja si es diferente

# Leer el archivo Excel
print("Leyendo archivo Excel...")
df = pd.read_excel(ARCHIVO_EXCEL, sheet_name=HOJA, header=3)  # header=3 porque los datos empiezan en la fila 4

# Mostrar las primeras filas para verificar
print("\nPrimeras filas del archivo:")
print(df.head())
print(f"\nColumnas encontradas: {list(df.columns)}")
print(f"\nTotal de registros: {len(df)}")

# Limpiar datos y crear diccionario de contactos
contactos = {}
for index, row in df.iterrows():
    # Obtener los datos de cada columna
    ruta = str(row['Ruta']).strip() if 'Ruta' in df.columns else ''
    usuario = str(row['Usuario']).strip() if 'Usuario' in df.columns else ''
    promotor = str(row['Promotor']).strip() if 'Promotor' in df.columns else ''
    numeros = str(row['Numeros']).strip() if 'Numeros' in df.columns else ''
    
    # Validar que el nombre del promotor sea válido
    if promotor == 'nan' or promotor == '' or pd.isna(promotor):
        continue
    
    # Validar que el número sea válido
    if numeros == 'nan' or numeros == '' or numeros == '0' or pd.isna(numeros):
        continue
    
    # Formatear el número
    if not numeros.startswith('+'):
        numeros = '+52' + numeros.replace(' ', '').replace('-', '')
    
    # Guardar en el diccionario con la información completa
    contactos[promotor] = {
        'ruta': ruta,
        'usuario': usuario,
        'numero': numeros
    }

print(f"\nContactos válidos encontrados: {len(contactos)}")
print("\nLista de contactos:")
for nombre, datos in contactos.items():
    print(f"  - {nombre} (Ruta: {datos['ruta']}, Usuario: {datos['usuario']})")
    print(f"      Número: {datos['numero']}")


# Función para enviar mensaje
def enviar_whatsapp(numero, mensaje):
    try:
        kit.sendwhatmsg_instantly(numero, mensaje, wait_time=10, tab_close=False)
        time.sleep(7)
        pyautogui.press("enter")  # enviar el texto
        time.sleep(3)
    except Exception as e:
        print(f"Error enviando a {numero}: {e}")


# ---- MENSAJE PERSONALIZADO ----
def crear_mensaje(nombre, ruta):
    return f"""Hola buen día, {nombre}, de la ruta {ruta} le hablo de Retail Optics y no se ve reflejada su efectividad correctamente 
    ¿Esta iniciando y cerrando sesión todos los días?
    ¿Esta haciendo Check in con su respectivo Check out en sus tiendas?
    """


# ---- Envío de mensajes ----
print("\n--- INICIANDO ENVÍO DE MENSAJES ---\n")

respuesta = input("¿Deseas proceder con el envío de mensajes? (s/n): ")

if respuesta.lower() == 's':
    for nombre, datos in contactos.items():
        mensaje = crear_mensaje(nombre, datos['ruta'])
        numero = datos['numero']
        
        print(f"\nEnviando mensaje a {nombre} (Ruta: {datos['ruta']}) - {numero}...")
        enviar_whatsapp(numero, mensaje)
        time.sleep(5)  # Pausa entre mensajes para evitar bloqueos
        print(f"✓ Mensaje enviado a {nombre} ({numero})")
    
    print("\n--- ENVÍO COMPLETADO ---")
else:
    print("\nEnvío cancelado por el usuario.")