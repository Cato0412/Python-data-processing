import os
import re
from PyPDF2 import PdfReader, PdfWriter

# 📂 Rutas
ruta_pdf = r"C:\Users\lapmxdf558\Documents\JUAN\OTROS\DIANA\RECIBOS.pdf"
carpeta_salida = r"C:\Users\lapmxdf558\Documents\JUAN\OTROS\DIANA\RESULTADOS"
os.makedirs(carpeta_salida, exist_ok=True)

# 📖 Abrir PDF
with open(ruta_pdf, "rb") as archivo:
    lector = PdfReader(archivo)
    total_paginas = len(lector.pages)
    print(f"Total de páginas: {total_paginas}")

    for i in range(total_paginas):
        # 🔍 Extraer texto de la página
        pagina = lector.pages[i]
        texto = pagina.extract_text() or ""
        
        # 🧩 Buscar folio y nombre
        patron = r"\b(\d{5,6})\s+([A-ZÁÉÍÓÚÑ ]+?)\s+C\.P\."
        coincidencia = re.search(patron, texto, re.DOTALL)

        if coincidencia:
            folio = coincidencia.group(1).strip()
            nombre_extraido = coincidencia.group(2).strip()
            # 🔥 Limpiar caracteres inválidos
            nombre_limpio = re.sub(r'[\\/*?:"<>|\r\n]+', ' ', nombre_extraido)
            nombre_limpio = re.sub(r'\s+', ' ', nombre_limpio).strip()
            nombre_final = f"{nombre_limpio}"
        else:
            nombre_final = f"pagina_{i+1}"

        # Crear PDF con una sola página
        escritor = PdfWriter()
        escritor.add_page(pagina)

        # Guardar archivo
        nombre_archivo = os.path.join(carpeta_salida, f"{nombre_final}.pdf")
        with open(nombre_archivo, "wb") as salida:
            escritor.write(salida)

        print(f"Página {i+1} guardada como: {nombre_archivo}")

print("✅ ¡PDF separado y nombrado con folio y nombre!")
