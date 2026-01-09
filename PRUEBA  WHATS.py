import os
import glob
import pywhatkit as kit
import pyautogui
import time
import pyperclip

# Carpeta con archivos
RUTA = r"C:\\Users\\lapmxdf558\\Documents\\JUAN\\BONOS PY\\ASISTENCIA\\SUP"
archivos = glob.glob(os.path.join(RUTA, "*.xlsx"))

# Diccionario: nombre → número de WhatsApp
contactos = {
    "Preza Parra Ivan De Jesus": "+525573969007",
    "Ortiz Heredia Eder": "+525573969007",
    "Ibarra Medina Ana Leticia": "+525573969007",
    "Nuñez Monarrez Elsa Guadalupe": "+525573969007"
}

# Recorremos los archivos


for archivo in archivos:
    nombre = os.path.splitext(os.path.basename(archivo))[0]  # Ej: "JuanPerez"
    nombrec=os.path.abspath(archivo)

    if nombre in contactos:
        numero = contactos[nombre]
        mensaje = f"Hola {nombre}!!!!, Te adjunto el archivo de asistencia y efectividad del día de ayer, el cual muestra la efectividad de tu equipo, el personal que no aparece en el archivo quiere decir que no asistio el dia de ayer, pido tu apoyo para que a esa gente faltante se le haga la observacion acerca de su captura diaria , de lo contrario eso los perjudicara en su efectividad y consecuentemente en sus pagos. Cualquier duda al respecto pueden acudir al ejecutivo Luis Arturo Hernandes al numero 5512345678. Saludos!!! "
        
        print(f"Enviando a {nombre} ({numero})...")
        
        # Enviar mensaje (aquí se abre WhatsApp Web)
        kit.sendwhatmsg_instantly(numero, mensaje, wait_time=7, tab_close=False)
        
        time.sleep(5)  # Espera a que se cargue bien WhatsApp Web

        # Simular click en el ícono de clip (coordenadas de tu pantalla, hay que ajustarlas)
        pyautogui.click(x=923, y=940)  
        time.sleep(1)

        # Click en "Documentos" (también ajustar coordenadas según tu pantalla)
        pyautogui.click(x=905, y=434)
       
        time.sleep(1)

        
        # Escribir la ruta del archivo
        pyperclip.copy(nombrec)        # Copia la ruta al portapapeles
        time.sleep(0.5)
        pyautogui.hotkey("ctrl", "v")  # Pega la ruta en WhatsApp
        time.sleep(1)
        pyautogui.press("enter")  # Selecciona el archivo
        time.sleep(1)

        # Enviar
        pyautogui.press("enter")

        time.sleep(3)

    else:
        print(f"No se encontró número para {nombre}")
