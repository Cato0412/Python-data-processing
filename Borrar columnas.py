import pandas as pd

# Ruta del archivo Excel
archivo_entrada = 'Inv-Compilados.xlsx'
archivo_salida = 'archivo_limpio.xlsx'

# Leer el archivo Excel
df = pd.read_excel(archivo_entrada)

# Mostrar columnas originales
print(f"Columnas originales ({len(df.columns)}):")
print(df.columns.tolist())

# Filtrar columnas que NO contengan "Fecha Respuesta"
columnas_mantener = [col for col in df.columns if 'Fecha Respuesta' not in str(col)]

# Crear DataFrame con las columnas filtradas
df_limpio = df[columnas_mantener]

# Mostrar columnas eliminadas
columnas_eliminadas = [col for col in df.columns if col not in columnas_mantener]
print(f"\nColumnas eliminadas ({len(columnas_eliminadas)}):")
print(columnas_eliminadas)

print(f"\nColumnas restantes ({len(df_limpio.columns)}):")
print(df_limpio.columns.tolist())

# Guardar el archivo limpio
df_limpio.to_excel(archivo_salida, index=False)

print(f"\nArchivo guardado como: {archivo_salida}")