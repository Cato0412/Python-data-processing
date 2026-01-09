#BLOQUE ENVIO DE WHATSAPP
#----------------------------------------------------------------------------------------------
import os
import glob
import pywhatkit as kit
import pyautogui
import pyperclip
import time
import pandas as pd

# Carpeta con archivos
RUTA = r"C:\Users\lapmxdf558\Documents\JUAN\BONOS PY\ASISTENCIA\SUP"
archivos = glob.glob(os.path.join(RUTA, "*.xlsx"))

# Diccionario: nombre → número de WhatsApp
contactos = {
    "Preza Parra Ivan De Jesus": "+525580069038",
    "Ortiz Heredia Eder": "+525580582033",
    "Ibarra Medina Ana Leticia": "+525579905497",
    "Nuñez Monarrez Elsa Guadalupe": "+525579472585",
    "Almora Villanueva Yu-Liang Misuko":"+525580582209",
    "Alvarez Sanchez Sergio Alejandro":"+525580582032",
    "Andrade Garcia Jorge Omar":"+525579909897",
    "Aquino Bojado Luz Adriana":"+525579996979",
    "Barragan Robles Maria Del Carmen":"+525580034977",
    "Cervantes Silva Ivan":"+525641787829",
    "Garza Torres Nallely Yazmin":"+525580033829",
    "Hernandez Gonzalez Jose Ivan":"+525549417761",
    "Martinez Montantes Lazaro":"+525580030158",
    "Moreno Flores Luz Andrea":"+525548611479",
    "Perez Rios Luis Alberto":"+525579474061",
    "Reyes Vazquez Cipriano":"+525625538793",
    "Rodriguez Cerda Maria Angelica":"+525541900030",
    "Solis Lucas Pedro":"+525578679426",
    "Vacante_Guanajuato_301":"+525579997033",
    "Vela Morales Nestor Adan":"+525579482048",
    "Covarrubias Rodriguez Sergio Adrian":"+525579473307"

}


# Crear diccionario con nombres de archivo normalizados
archivos_nombres = {
    os.path.splitext(os.path.basename(a))[0].strip().lower(): a for a in archivos
}

# ---- Tabla comparativa para depuración ----
comparacion = []
for nombre in contactos.keys():
    norm = nombre.strip().lower()
    tiene_archivo = "✅ Sí" if norm in archivos_nombres else "❌ No"
    comparacion.append([nombre, contactos[nombre], tiene_archivo])

df = pd.DataFrame(comparacion, columns=["Nombre Contacto", "Número", "Archivo encontrado"])
print("\nomparación contactos vs archivos:\n")
print(df.to_string(index=False))


# Función para enviar mensaje con o sin archivo
def enviar_whatsapp(numero, mensaje, archivo_ruta=None):
    try:
        kit.sendwhatmsg_instantly(numero, mensaje, wait_time=10, tab_close=False)
        time.sleep(7)
        pyautogui.press("enter")  # enviar el texto

        if archivo_ruta and os.path.exists(archivo_ruta):
            pyautogui.click(x=923, y=940)  # Clip (ajustar coordenadas)
            time.sleep(1)
            pyautogui.click(x=905, y=434)  # Documentos (ajustar coordenadas)
            time.sleep(1)
            pyperclip.copy(archivo_ruta)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(1)
            pyautogui.press("enter")  # Selecciona archivo
            time.sleep(1)
            pyautogui.press("enter")  # Envía archivo
            time.sleep(3)

    except Exception as e:
        print(f"Error enviando a {numero}: {e}")


# ---- Envío de mensajes ----
for nombre, numero in contactos.items():
    nombre_norm = nombre.strip().lower()

    if nombre_norm in archivos_nombres:  # Hay archivo
        archivo = archivos_nombres[nombre_norm]
        mensaje = (
            f"Buen día a traves de este número se enviaran los reportes de asistencia, estos mensajes"
            " se encuentran programados automáticamente, por lo que cualquier duda pueden acudir al "
            "ejecutivo Luis Arturo Hernandez al número 5532099064."
            f"Hola {nombre}!!!!, Te adjunto el archivo de asistencia y efectividad del día de ayer, "
            "el cual muestra la efectividad de tu equipo a traves de lo registrado en Retail Optics. El personal que no aparece en el archivo no tiene registros en la aplicación. "
            "Te pido apoyo para que a esa gente faltante se le haga la observación sobre su captura diaria. "
            "De lo contrario eso los perjudicará en su efectividad y consecuentemente en sus pagos. "
            "Cualquier duda comunicarse con Arturo Hernandez al numero adjunto "
            "Saludos!!!"
        )
        
        print(f"Enviando a {nombre} ({numero}) con archivo...")
        enviar_whatsapp(numero, mensaje, archivo_ruta=archivo)

    else:  # No hay archivo
        mensaje = (
            f"Buen día a traves de este número se enviaran los reportes de asistencia, estos mensajes"
            " se encuentran programados automaticamente, por lo que cualquier duda pueden acudir al "
            "ejecutivo Luis Arturo Hernandez al número 5532099064."
            "Cualquier duda comunicarse con Arturo Hernandez al numero adjunto"
            f"Hola {nombre}, no hubo asistencia de tu equipo registrada el día de ayer."
            )
        print(f"Enviando mensaje de 'no hubo asistencia' a {nombre} ({numero})...")
        enviar_whatsapp(numero, mensaje)
