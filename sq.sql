Columnas del archivo Excel: Index(['TOMO', 'LIBRO', 'SESION ', 'AÑO', 'VOLUMEN', 'INSCRIPCION', 'CLAVE',
       'ID_FOLIO', 'Unnamed: 8', 'CARGA'],
      dtype='object')
Consulta prelacion: 
        SELECT * FROM prelacion p 
        WHERE p.consecutivo = %s AND anio = %s;
     con valores (52872, 2024)
prelacion----- [(2350612, 2024, 52872, None, None, False, None, True, None, None, datetime.datetime(2024, 9, 20, 10, 49, 45, 670000), datetime.datetime(2024, 9, 20, 10, 49, 45, 670000), None, None, None, None, 10, None, 'ENTRADA+SISTEMA+CONTINO', None, None, None, None, None, None, None, None, None, None, None, None, 1, None, None, None, 11, 1, None, 4, None, 51721, None, 51721, None, None, 0, None, 61, None, None, None, None, None, None, False, None, 1, False, True, True, 0, None, None, None, None, 'sA8ZcX')]
Consulta documento: 
        SELECT * FROM prelacion_ante pa 
        WHERE prelacion_id = %s AND documento = %s AND anio = %s AND libro = %s AND volumen = %s;
     con valores (2350612, 3, 1975, 0, 1)
---antecedente (572419, 1975, '3', None, '0', None, False, False, None, None, None, 11, None, None, 2350612, 1, 4, None, '1', '0', None, None)
Encontrado
Consulta predio: 
        SELECT id, no_folio FROM predio WHERE id = %s;
     con valor predio_id = None
predio.. None





















-----------------------------------------------------------------------------------
Consulta prelacion: 
        SELECT * FROM prelacion p 
        WHERE p.consecutivo = %s AND anio = %s;
     con valores (52872, 2024)
prelacion----- [(2350612, 2024, 52872, None, None, False, None, True, None, None, datetime.datetime(2024, 9, 20, 10, 49, 45, 670000), datetime.datetime(2024, 9, 20, 10, 49, 45, 670000), None, None, None, None, 10, None, 'ENTRADA+SISTEMA+CONTINO', None, None, None, None, None, None, None, None, None, None, None, None, 1, None, None, None, 11, 1, None, 4, None, 51721, None, 51721, None, None, 0, None, 61, None, None, None, None, None, None, False, None, 1, False, True, True, 0, None, None, None, None, 'sA8ZcX')]
Consulta documento: 
        SELECT * FROM prelacion_ante pa 
        WHERE prelacion_id = %s AND documento = %s AND anio = %s AND libro = %s AND volumen = %s;
     con valores (2350612, 2, 1975, 91, 1)
---antecedente None
Duplicado
-----------------------------------------------------------------------------------
Archivo Excel actualizado y guardado con color para duplicados