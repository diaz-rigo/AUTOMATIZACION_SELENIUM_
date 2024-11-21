import requests
import pandas as pd
from io import BytesIO

# Definir el filtro de entrada
ENTRADA_FILTRO = 52881

# URL del archivo Google Sheets modificado para exportar como Excel
SHEET_URL = "https://docs.google.com/spreadsheets/d/1ZdZnlu78HdxVv_zh8hw4KUQLBWnjpHdA/export?format=xlsx"

# Descargar el archivo Excel desde la URL
response = requests.get(SHEET_URL)

# Verificar si la descarga fue exitosa
if response.status_code == 200:
    # Leer el archivo Excel directamente desde la memoria utilizando BytesIO
    df = pd.read_excel(BytesIO(response.content), sheet_name=str(ENTRADA_FILTRO))

    # Verificar el contenido de la columna 'predio_id'
    print("Valores únicos en 'predio_id':")
    print(df['predio_id'].unique())  # Para revisión de valores

    # ---- FILTRO 1: Predio no encontrado ----
    filtro_1 = (df['no_folio'] == 'Predio no encontrado') & \
               (df['indicador_duplicado'] == 'Encontrado') & \
               (df['cantidad_duplicados'] == 1) & \
               (df['predio_id'].isna())  # Filtrar sólo valores nulos

    df_filtrado_1 = df[filtro_1]
    print("::::::::::::::::::::::::SIN PREDIO::::::::::::::::::::::::::::::::::::::::::::")

    print("\nFiltrado 1: 'Predio no encontrado, Encontrado, 1 duplicado, predio_id = NaN'")
    print(f"Cantidad de filas encontradas: {len(df_filtrado_1)}")
    if not df_filtrado_1.empty:
        print(df_filtrado_1)
    else:
        print("No se encontraron filas que coincidan con la primera condición.")

    # ---- FILTRO 2: Antecedente no encontrado ----
    filtro_2 = (df['no_folio'] == 'Antecedente no encontrado') & \
               (df['indicador_duplicado'] == 'Duplicado') & \
               (df['cantidad_duplicados'] >= 2) & \
               (df['predio_id'].isna()) 

    df_filtrado_2 = df[filtro_2]

    print("::::::::::::::::::::::::DUPLICADOS::::::::::::::::::::::::::::::::::::::::::::")
    print("\nFiltrado 2: 'Antecedente no encontrado, Duplicadoduplicados, predio_id = NaN'")
    print(f"Cantidad de filas encontradas: {len(df_filtrado_2)}")
    if not df_filtrado_2.empty:
        print(df_filtrado_2)
    else:
        print("No se encontraron filas que coincidan con la segunda condición.")

else:
    print(f"Error al descargar el archivo: {response.status_code}")
