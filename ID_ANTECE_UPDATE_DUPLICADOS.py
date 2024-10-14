import psycopg2
import pandas as pd

# Definir rango para filtrar los datos
RANGO_INICIO = 1
RANGO_FIN = 3

# Conexión a la base de datos PostgreSQL
conn = psycopg2.connect(
    dbname="erpp",
    user="angel",
    password="ErppHgo&2024",
    host="100.66.168.122",
    port="5432"
)
cursor = conn.cursor()

# Ruta del archivo Excel
file_path = '/Users/arturo/Documents/SCRIP_ UPDATE_DUPLICADO/predios_con_antecedentes_52879UPDATE.xlsx'

# Cargar el archivo Excel usando Pandas
df = pd.read_excel(file_path)

# Filtrar el DataFrame según los criterios del rango en la columna 'CARGA'
df_filtrado = df[(df['CARGA'] >= RANGO_INICIO) & (df['CARGA'] <= RANGO_FIN)]

# Crear listas para almacenar los datos antiguos y actualizados
datos_anteriores = []
datos_actualizados = []

# Iterar sobre las filas filtradas
for index, row in df_filtrado.iterrows():
    # Extraer los valores necesarios del Excel
    anio = int(row['AÑO'])  # Convertir a entero para evitar el .0
    volumen = int(row['VOLUMEN'])  # Convertir a entero
    secciones_por_oficina_id = 55  # Asumiendo un valor fijo para la oficina
    libro_bis = 0  # Asumimos valor fijo 0 para 'libro_bis'
    num_libro = int(row['TOMO'])  # Convertir a entero
    INSCRIPCION = row['INSCRIPCION']  # Columna 'INSCRIPCION'
    
    # Segunda consulta: prelacion_ante
    id_pre_ante = int(row['id'])  # Convertir a entero
    predio_id = row['predio_id']  # Columna 'predio_id'
    antecedente_id = row['antecedente_id']  # Columna 'antecedente_id'

    # Realizar la consulta a la tabla prelacion_ante
    query_prelacion = f"""
    SELECT id, anio, documento, libro_bis, predio_id, volumen, libro, prelacion_id 
    FROM prelacion_ante 
    WHERE id = {id_pre_ante}
    """
    cursor.execute(query_prelacion)
    result_prelacion = cursor.fetchone()

    if result_prelacion:  # Verificar si la consulta retorna resultados
        # Guardar los datos antiguos antes de la actualización
        datos_anteriores.append({
            "id_pre_ante": result_prelacion[0],
            "anio": result_prelacion[1],
            "documento": result_prelacion[2],
            "volumen": result_prelacion[5],
            "libro": result_prelacion[6]
        })
        
        # Comparar valores y ejecutar actualizaciones si es necesario
        if anio != result_prelacion[1] or volumen != result_prelacion[5] or num_libro != result_prelacion[6]:
            print(f"Actualizando prelacion_ante ID {id_pre_ante}...")
            
            # Consultar el ID del libro en la tabla 'libro'
            query_libro = f"""
            SELECT * 
            FROM libro 
            WHERE anio = {anio} 
            AND volumen = '{volumen}' 
            AND secciones_por_oficina_id = {secciones_por_oficina_id} 
            AND libro_bis = '{libro_bis}' 
            AND num_libro = '{num_libro}'
            """
            cursor.execute(query_libro)
            result_libro = cursor.fetchone()  # Obtenemos una fila de la consulta
            if result_libro:
                ID_LIBRO = result_libro[0]
                
                # Actualizar la tabla prelacion_ante
                update_prelacion_ante = f"""
                    UPDATE prelacion_ante
                    SET anio = {anio}, documento = '{INSCRIPCION}', volumen = '{volumen}', libro = '{num_libro}'
                    WHERE id = {id_pre_ante};
                """
                cursor.execute(update_prelacion_ante)

                # Actualizar la tabla antecedente
                update_antecedente = f"""
                UPDATE antecedente 
                SET documento = '{INSCRIPCION}', libro_id = '{ID_LIBRO}'
                WHERE id = {antecedente_id};
                """
                cursor.execute(update_antecedente)

                # Guardar los datos actualizados después de la actualización
                datos_actualizados.append({
                    "id_pre_ante": id_pre_ante,
                    "anio": anio,
                    "documento": INSCRIPCION,
                    "volumen": volumen,
                    "libro": num_libro
                })
            else:
                print(f"No se encontró un libro para anio {anio}, volumen {volumen}, num_libro {num_libro}")

# Guardar los cambios en la base de datos
conn.commit()

# Crear DataFrames con los datos anteriores y actualizados
df_anteriores = pd.DataFrame(datos_anteriores)
df_actualizados = pd.DataFrame(datos_actualizados)

# Guardar los DataFrames en un archivo Excel con dos hojas
output_file_path = '/Users/arturo/Documents/SCRIP_ UPDATE_DUPLICADO/UPDATEpredios_con_antsultas_Libro_Prelacion.xlsx'
with pd.ExcelWriter(output_file_path) as writer:
    df_anteriores.to_excel(writer, sheet_name='Datos Anteriores', index=False)
    df_actualizados.to_excel(writer, sheet_name='Datos Actualizados', index=False)

# Cerrar la conexión a la base de datos
cursor.close()
conn.close()

print(f"Los resultados se han guardado en {output_file_path}")
