"""
Scripts para limpieza y extracción de datos de direcciones en Excel
Autor: Asistente Claude
Fecha: 2025

Requisitos:
pip install pandas openpyxl
"""

import pandas as pd
import re
import os
from pathlib import Path


class AddressProcessor:
    """Clase para procesar direcciones en archivos Excel"""
    
    def __init__(self):
        self.patterns_to_remove = ['No.', 'No', 'N0.', 'N0', '#']
    
    def extract_street_name(self, address):
        """Extrae solo el nombre de la calle (antes del primer número)"""
        if pd.isna(address) or not isinstance(address, str):
            return address
        
        # Buscar el primer número
        match = re.search(r'\d+', address)
        
        if match:
            # Extraer todo antes del primer número
            street_name = address[:match.start()].strip()
            
            # Eliminar patrones no deseados del final
            for pattern in self.patterns_to_remove:
                street_name = re.sub(rf'{re.escape(pattern)}\s*$', '', street_name, flags=re.IGNORECASE).strip()
            
            return street_name
        
        return address.strip()
    
    def extract_first_number(self, address):
        """Extrae el primer número de la dirección"""
        if pd.isna(address) or not isinstance(address, str):
            return ''
        
        # Eliminar N0 para evitar confusión
        clean_address = re.sub(r'N0\.?\s*', '', address, flags=re.IGNORECASE)
        
        # Buscar el primer número
        match = re.search(r'\d+', clean_address)
        
        return match.group(0) if match else ''
    
    def clean_address_until_number(self, address):
        """Limpia la dirección dejando solo hasta el primer número (incluido)"""
        if pd.isna(address) or not isinstance(address, str):
            return address
        
        # Buscar el primer número
        match = re.search(r'\d+', address)
        
        if match:
            # Extraer desde el inicio hasta después del número
            return address[:match.end()].strip()
        
        return address.strip()
    
    def process_file(self, input_file, output_file, column_name, operation='extract_street'):
        """
        Procesa un archivo Excel aplicando la operación especificada
        
        Args:
            input_file: Ruta del archivo de entrada
            output_file: Ruta del archivo de salida
            column_name: Nombre de la columna a procesar
            operation: Tipo de operación ('extract_street', 'extract_number', 'clean_until_number')
        """
        try:
            # Leer el archivo Excel
            df = pd.read_excel(input_file)
            
            # Verificar que la columna existe
            if column_name not in df.columns:
                raise ValueError(f"La columna '{column_name}' no existe en el archivo")
            
            # Aplicar la operación correspondiente
            if operation == 'extract_street':
                df[column_name] = df[column_name].apply(self.extract_street_name)
            elif operation == 'extract_number':
                df[column_name] = df[column_name].apply(self.extract_first_number)
            elif operation == 'clean_until_number':
                df[column_name] = df[column_name].apply(self.clean_address_until_number)
            else:
                raise ValueError(f"Operación '{operation}' no reconocida")
            
            # Guardar el archivo procesado
            df.to_excel(output_file, index=False)
            print(f"✓ Archivo procesado exitosamente: {output_file}")
            
            return df
            
        except Exception as e:
            print(f"✗ Error al procesar {input_file}: {str(e)}")
            return None
    
    def process_multiple_files(self, input_folder, output_folder, column_name, operation='extract_street'):
        """
        Procesa múltiples archivos Excel en una carpeta
        
        Args:
            input_folder: Carpeta con archivos de entrada
            output_folder: Carpeta para archivos de salida
            column_name: Nombre de la columna a procesar
            operation: Tipo de operación
        """
        # Crear carpeta de salida si no existe
        Path(output_folder).mkdir(parents=True, exist_ok=True)
        
        # Buscar todos los archivos Excel
        excel_files = list(Path(input_folder).glob('*.xlsx')) + list(Path(input_folder).glob('*.xls'))
        
        if not excel_files:
            print(f"No se encontraron archivos Excel en {input_folder}")
            return
        
        print(f"Encontrados {len(excel_files)} archivos para procesar\n")
        
        results = []
        for file in excel_files:
            output_file = Path(output_folder) / f"{file.stem}_{operation}{file.suffix}"
            result = self.process_file(str(file), str(output_file), column_name, operation)
            if result is not None:
                results.append((file.name, len(result)))
        
        print(f"\n✓ Procesamiento completado: {len(results)}/{len(excel_files)} archivos")


# ====== SCRIPTS DE USO DIRECTO ======

def extract_streets(input_file, output_file, column_name='E'):
    """Script simple para extraer nombres de calles"""
    processor = AddressProcessor()
    processor.process_file(input_file, output_file, column_name, 'extract_street')


def extract_numbers(input_file, output_file, column_name='E'):
    """Script simple para extraer números"""
    processor = AddressProcessor()
    processor.process_file(input_file, output_file, column_name, 'extract_number')


def clean_addresses(input_file, output_file, column_name='E'):
    """Script simple para limpiar direcciones hasta el número"""
    processor = AddressProcessor()
    processor.process_file(input_file, output_file, column_name, 'clean_until_number')


# ====== EJEMPLO DE USO ======

if __name__ == "__main__":
    """
    Ejemplos de uso de los scripts
    """
    
    # Crear instancia del procesador
    processor = AddressProcessor()
    
    print("=== PROCESADOR DE DIRECCIONES ===\n")
    
    # OPCIÓN 1: Procesar un solo archivo
    print("OPCIÓN 1: Procesar un archivo individual")
    print("-" * 50)
    
    # Ejemplo 1: Extraer nombres de calles
    # extract_streets('mi_archivo.xlsx', 'calles_extraidas.xlsx', column_name='E')
    
    # Ejemplo 2: Extraer números
    # extract_numbers('mi_archivo.xlsx', 'numeros_extraidos.xlsx', column_name='E')
    
    # Ejemplo 3: Limpiar hasta el número
    # clean_addresses('mi_archivo.xlsx', 'direcciones_limpias.xlsx', column_name='E')
    
    
    # OPCIÓN 2: Procesar múltiples archivos en una carpeta
    print("\nOPCIÓN 2: Procesar múltiples archivos")
    print("-" * 50)
    
    # Ejemplo: Procesar todos los archivos de una carpeta
    # processor.process_multiple_files(
    #     input_folder='./archivos_entrada',
    #     output_folder='./archivos_salida',
    #     column_name='E',
    #     operation='extract_street'  # o 'extract_number' o 'clean_until_number'
    # )
    
    
    # OPCIÓN 3: Uso manual con DataFrame
    print("\nOPCIÓN 3: Uso manual con DataFrames")
    print("-" * 50)
    print("# Cargar archivo")
    print("df = pd.read_excel('archivo.xlsx')")
    print("\n# Extraer calles")
    print("df['E'] = df['E'].apply(processor.extract_street_name)")
    print("\n# Extraer números")
    print("df['E'] = df['E'].apply(processor.extract_first_number)")
    print("\n# Guardar")
    print("df.to_excel('resultado.xlsx', index=False)")
    
    
    print("\n\n=== INSTRUCCIONES ===")
    print("1. Descomenta las líneas de ejemplo que necesites")
    print("2. Ajusta las rutas de archivos y nombres de columnas")
    print("3. Ejecuta el script: python limpieza_direcciones.py")
    print("\nOperaciones disponibles:")
    print("  - extract_street: Extrae solo el nombre de la calle")
    print("  - extract_number: Extrae solo el primer número")
    print("  - clean_until_number: Deja la dirección hasta el primer número")