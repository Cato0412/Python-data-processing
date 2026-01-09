import pandas as pd
import matplotlib.pyplot as plt

# Leer archivo
df = pd.read_excel(r'C:\Users\lapmxdf558\Documents\JUAN\HC\HC PLANTILLA.xlsx')

# Seleccionar Top 10 por sueldo más alto
df_top10 = df.sort_values(by="SUELDO DIARIO", ascending=False).head(10)

# Columnas
nombres = df_top10["NOMBRE COMPLETO"]
SD = df_top10["SUELDO DIARIO"]

# Gráfica
plt.figure(figsize=(10, 6))
plt.bar(nombres, SD, color='red')
plt.title("Top 10 Sueldos más Altos")
plt.xlabel("Nombre")
plt.ylabel("Sueldo")
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.show()
