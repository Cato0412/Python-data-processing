import pandas as pd

def safe_drop(df, cols):
    """
    Borra columnas de un DataFrame sin dar error si no existen.
    - df: DataFrame de entrada
    - cols: lista de columnas a eliminar
    """
    # Filtrar solo columnas que existen
    cols_to_drop = [c for c in cols if c in df.columns]
    return df.drop(columns=cols_to_drop)

# Ejemplo de uso
df = pd.DataFrame({
    "a": [1, 2],
    "b": [3, 4]
})

# Intentar borrar columnas (aunque alguna no exista)
df = safe_drop(df, ["b", "x", "y"])
print(df)
