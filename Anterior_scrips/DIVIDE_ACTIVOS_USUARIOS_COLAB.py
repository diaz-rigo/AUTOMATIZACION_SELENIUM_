# Instalar xlsxwriter
!pip install xlsxwriter

import os
import pandas as pd

# Subir archivo
from google.colab import files
uploaded = files.upload()

# Seleccionar el archivo subido
file_path = list(uploaded.keys())[0]
# file_path = '/content/Captura Insercion Masiva (1).xlsx'

# Cargar el archivo Excel
df = pd.read_excel(file_path, sheet_name='Activos')

# Crear carpeta para almacenar los excels si no existe
output_folder = 'output_excels'
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Filtrar por cada usuario basado en el campo de USUARIO
usuarios = df['USUARIO'].unique()

for usuario in usuarios:
    # Filtrar los datos por el USUARIO
    df_usuario = df[df['USUARIO'] == usuario]
    
    # Crear un archivo Excel para cada usuario en la carpeta 'output_excels'
    output_file = os.path.join(output_folder, f'{usuario}.xlsx')
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    
    # Filtrar por cada entrada y crear una hoja por cada entrada
    entradas = df_usuario['ENTRADA'].unique()
    for entrada in entradas:
        # Eliminar el '.0' si es un número decimal sin parte fraccionaria
        if isinstance(entrada, float) and entrada.is_integer():
            entrada_str = str(int(entrada))  # Convertir a entero si es flotante con .0
        else:
            entrada_str = str(entrada)  # Convertir a cadena
        
        # Guardar los datos en la hoja con el nombre correcto
        df_entrada = df_usuario[df_usuario['ENTRADA'] == entrada]
        df_entrada.to_excel(writer, sheet_name=entrada_str, index=False)
    
    # Guardar el archivo Excel
    writer.close()

print("Archivos Excel generados en la carpeta 'output_excels'.")
