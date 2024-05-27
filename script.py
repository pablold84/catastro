import pandas as pd
from sqlalchemy import create_engine
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
import os

# Configura tus parámetros de conexión a PostgreSQL
conn_params = {
    'host': 'cartofs.seresco.red',
    'dbname': 'SEGIPSA',
    'user': 'Segipsa',
    'password': 'Segipas24',
    'port': 5432
}

# Crear una cadena de conexión utilizando SQLAlchemy
conn_str = f"postgresql+psycopg2://{conn_params['user']}:{conn_params['password']}@{conn_params['host']}:{conn_params['port']}/{conn_params['dbname']}"

# Crear el motor de conexión
engine = create_engine(conn_str)

# Definir consultas SQL y campos correspondientes
consultas_campos = {
    "campo_1": {
        "consulta": "SELECT del FROM cabrales.segipsa_placo where foto_croquis='CAMPO/FOTOS-CAMPO-REDUCIDAS/DESCARGA-1/1- (33).jpg'",
        "campo": "del",
        "texto_agregar": "GERENCIA-MUNICIPIO: ",
        "celda_destino": "A7"
    },
    "campo_2": {
        "consulta": "SELECT exp, control, anio FROM cabrales.segipsa_placo where foto_croquis='CAMPO/FOTOS-CAMPO-REDUCIDAS/DESCARGA-1/1- (33).jpg'",
        "campo": "expediente",
        "texto_agregar": "Nº EXPEDIENTE: ",
        "celda_destino": "P7"
    }
}

# Obtiene la ruta al directorio actual donde se encuentra el script
dir_actual = os.path.dirname(__file__)

# Combina la ruta al directorio actual con la ruta relativa al archivo
file_path = os.path.join(dir_actual, "modelo6", "origen", "ficheros", "FICHA_RESUMEN_PLACO23 (Modelo6).xlsx")

# Cargar el archivo Excel existente
book = load_workbook(file_path)

# Iterar sobre las consultas, obtener datos y escribir en Excel
for nombre_campo, detalles in consultas_campos.items():
    consulta_sql = detalles["consulta"]
    texto_agregar = detalles["texto_agregar"]
    celda_destino = detalles["celda_destino"]

    # Leer los datos de la base de datos PostgreSQL en un DataFrame de pandas
    df = pd.read_sql_query(consulta_sql, engine)

    if nombre_campo == "campo_2":
        # Convertir los valores enteros a cadenas antes de concatenarlos
        df["expediente"] = df["exp"].astype(str) + "." + df["control"].astype(str) + "/" + df["anio"].astype(str)
        texto_agregar += df["expediente"].iloc[0]#agrego la cadena antes compuesta

    # Seleccionar la hoja en la que quieres escribir los datos
    ws = book["FICHA RESUMEN PLACO"]  # Nombre de la hoja de tu plantilla

    # Obtener las coordenadas de la celda destino
    dest_col, dest_row = column_index_from_string(celda_destino[0]), int(celda_destino[1:])

    # Escribe el texto concatenado y los datos en la celda especificada
    if nombre_campo == "campo_2":
        ws.cell(row=dest_row, column=dest_col, value=texto_agregar)
    else:
        for idx, row in df.iterrows():
            for col, value in enumerate(row, start=1):
                ws.cell(row=dest_row, column=dest_col, value=texto_agregar + str(value))
                dest_col += 1

# Guardar los cambios en el archivo Excel
book.save(file_path)

print("Datos escritos exitosamente en el archivo existente")
