import psycopg2
from psycopg2 import sql
import pandas as pd

ENTRADA_ = 52879

# Función para obtener los datos desde la base de datos y exportar a Excel
def obtener_antecedente(consecutivo, anio):
    try:
        with psycopg2.connect(
            dbname="erpp",
            user="angel",
            password="ErppHgo&2024",
            host="100.66.168.122",
            port="5432"
        ) as conn:
            with conn.cursor() as cursor:
                # Paso 1: Obtener el prelacion_id basado en el consecutivo y anio
                query_prelacion = sql.SQL(""" 
                    SELECT id 
                    FROM prelacion 
                    WHERE consecutivo = %s AND anio = %s; 
                """)
                cursor.execute(query_prelacion, (consecutivo, anio))
                prelacion_id = cursor.fetchone()

                if prelacion_id is None:
                    print("No se encontró ninguna prelación con el consecutivo y año proporcionados.")
                    return

                prelacion_id = prelacion_id[0]  # Extraer el id

                # Paso 2: Obtener los predios junto con los datos de "libro"
                query_predio = sql.SQL(""" 
                    SELECT pa.id, pa.anio, pa.documento, pa.libro_bis, pa.predio_id, pa.volumen, pa.libro,
                        p.antecedente_id,
                        a.*, pa.libro_id,
                        l.*  -- Seleccionar todas las columnas de la tabla libro
                    FROM prelacion_ante pa
                    LEFT JOIN predio_ante p ON p.predio_id = pa.predio_id
                    LEFT JOIN antecedente a ON a.id = p.antecedente_id
                    LEFT JOIN libro l ON l.id = a.libro_id  -- Hacemos un LEFT JOIN con la tabla libro
                    WHERE pa.prelacion_id = %s 
                    AND (pa.anio, pa.documento, pa.libro_bis, pa.volumen, pa.libro) IN (
                        SELECT anio, documento, libro_bis, volumen, libro
                        FROM prelacion_ante
                        WHERE prelacion_id = %s 
                        GROUP BY anio, documento, libro_bis, volumen, libro
                        HAVING COUNT(*) > 1
                    )
                    ORDER BY pa.documento;
                """)
                cursor.execute(query_predio, (prelacion_id, prelacion_id))
                resultados = cursor.fetchall()  # Obtener todos los resultados

                # Obtener los nombres de las columnas directamente desde el cursor
                columnas = [desc[0] for desc in cursor.description]

                # Convertir a DataFrame
                df = pd.DataFrame(resultados, columns=columnas)
                
                # Exportar a Excel
                nombre_archivo = f"predios_con_antecedentes_{consecutivo}.xlsx"
                df.to_excel(nombre_archivo, index=False)
                print(f"Los datos se han exportado a '{nombre_archivo}'")

    except Exception as e:
        print(f"Ocurrió un error: {e}")

# Llamada a la función
obtener_antecedente(ENTRADA_, 2024)
