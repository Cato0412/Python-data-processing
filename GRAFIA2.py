import pandas as pd
import matplotlib.pyplot as plt

# Leer archivo
df1 = pd.read_excel(r'C:\Users\lapmxdf1558\Documents\JUAN\HC\HC PLANTILLA.xlsx')

# Agrupar por DEPARTAMENTO (puede ser cualquier columna categórica)
departamentos = df1['ESTADO'].value_counts()

# Gráfico de pastel
plt.figure(figsize=(10, 10))
plt.pie(
    departamentos.values,           # valores (cantidad de empleados por depto)
    labels=departamentos.index,     # etiquetas (nombres de los deptos)
    autopct='%1.0f%%',              # mostrar porcentaje con 1 decimal
    startangle=90,                  # empieza desde arriba
    colors=plt.cm.Paired.colors     # colores automáticos bonitos
)

plt.title('Distribución de Empleados por Estado')
plt.axis('equal')  # hace el gráfico circular
plt.show()
