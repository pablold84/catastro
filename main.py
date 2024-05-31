import os
import shutil
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import column_index_from_string
import pandas as pd
from sqlalchemy import create_engine
import tkinter as tk
from tkinter import filedialog, messagebox

# Función para configurar la conexión a la base de datos
def obtener_conexion():
    conn_params = {
        'host': 'cartofs.seresco.red',
        'dbname': 'SEGIPSA',
        'user': 'Segipsa',
        'password': 'Segipas24',
        'port': 5432
    }
    conn_str = f"postgresql+psycopg2://{conn_params['user']}:{conn_params['password']}@{conn_params['host']}:{conn_params['port']}/{conn_params['dbname']}"
    return create_engine(conn_str)

# Función para escribir datos iniciales en el Excel
def escribir_datos_iniciales(ws, schema, refcat_value):
    # Conectar a la base de datos
    engine = obtener_conexion()

    # Obtener todos los registros de la tabla DATOS_INICIALES filtrando por REFCAT
    df_datos_iniciales = pd.read_sql_query(f'SELECT * FROM "{schema}"."DATOS_INICIALES" WHERE "REFCAT" = %s', engine, params=(refcat_value,))

    # Verificar si se encontraron registros
    if df_datos_iniciales.empty:
        messagebox.showerror("Error", f"No se encontraron registros con REFCAT = {refcat_value}")
        return False

    # Campos y celdas de destino para DATOS_INICIALES
    mapeo_campos = {
        "CARGO": "A",
        "ORD_CONS": "B",
        "ES": "C",
        "PLA": "D",
        "PUE": "E",
        "TIPOL": "F",
        "CAT_PREDO": "G",
        "DES": "H",
        "SUP_LOCAL": "I",
        "U_CONS": "J",
        "AP_CO_CO": "K",
        "ANY_ANTIG": "L"
    }

    # Escribir los datos de DATOS_INICIALES en el Excel
    start_row = 5
    for idx, row in df_datos_iniciales.iterrows():
        for campo, col in mapeo_campos.items():
            if campo in row:
                valor_campo = row[campo]
                dest_col = column_index_from_string(col)
                dest_row = start_row + idx
                ws.cell(row=dest_row, column=dest_col, value=valor_campo)

    return True

# Función para escribir datos de SAUCE en el Excel
def escribir_datos_sauce(ws):
    # Definir el nombre del archivo CSV
    archivo_csv = "modelo6/origen/ficheros/sauce/33_217210_24.csv"

    # Diccionario para almacenar las líneas de cada sección
    secciones = {"FINCAS": [], "EXPEDIENTE": [], "CONSTRUCCIONES": [], "UNIDADES CONSTRUCTIVAS": []}

    # Bandera para indicar la sección actual
    seccion_actual = None

    # Leer el archivo CSV
    with open(archivo_csv, 'r') as file:
        for line in file:
            # Verificar si la línea está vacía (contiene solo el retorno de carro)
            if line.strip() == "":
                # Si la línea está vacía, cambiar a la siguiente sección
                seccion_actual = None
                continue

            # Detectar la sección actual
            if line.startswith("EXPEDIENTE"):
                seccion_actual = "EXPEDIENTE"
                continue
            elif line.startswith("FINCAS"):
                seccion_actual = "FINCAS"
                continue
            elif line.startswith("CONSTRUCCIONES"):
                seccion_actual = "CONSTRUCCIONES"
                continue
            elif line.startswith("UNIDADES CONSTRUCTIVAS"):
                seccion_actual = "UNIDADES CONSTRUCTIVAS"
                continue

            # Almacenar las líneas en la sección actual si no está vacía
            if seccion_actual in secciones and line.strip() != "":
                secciones[seccion_actual].append(line.strip())

    # Extraer datos de CONSTRUCCIONES
    construcciones_data = []
    for construccion in secciones["CONSTRUCCIONES"]:
        campos_construccion = construccion.split(";")
        cargo = campos_construccion[29]
        orden = campos_construccion[11]

        pcatastral1 = campos_construccion[9]
        pcatastral2 = campos_construccion[10]

        escalera = campos_construccion[13]
        planta = campos_construccion[14]
        puerta = campos_construccion[15]
        tipologia = campos_construccion[19]
        categ_pred = campos_construccion[20]
        destino = campos_construccion[18]
        superficie_total = campos_construccion[26]
        unidad_const_dest = campos_construccion[16]
        aa_antiguedad = campos_construccion[23]
        aa_reforma = campos_construccion[22]

        # Obtener el coeficiente de conservación de UNIDADES CONSTRUCTIVAS
        coef_conservacion = "N/A"
        for unidad in secciones["UNIDADES CONSTRUCTIVAS"]:
            campos_unidad = unidad.split(";")
            if pcatastral1 == campos_unidad[2] and pcatastral2 == campos_unidad[3] and unidad_const_dest == campos_unidad[4]:
                coef_conservacion = campos_unidad[16]
                break


        construcciones_data.append([cargo, orden, escalera, planta, puerta, tipologia, categ_pred, destino, superficie_total, unidad_const_dest, coef_conservacion, aa_antiguedad, aa_reforma])

    # Omitir el primer elemento (encabezado) de construcciones_data
    construcciones_data = construcciones_data[1:]

    # Definir mapeo_campos_sauce
    mapeo_campos_sauce = {
        "CARGO": "O",
        "ORDEN": "P",
        "ESCALERA": "Q",
        "PLANTA": "R",
        "PUERTA": "S",
        "TIPOLOGIA": "T",
        "CATEG_PRED": "U",
        "DESTINO": "V",
        "SUPERFICIE_TOTAL": "W",
        "UNIDAD_CONST": "X",
        "COEF_CONSERVACION": "Y",
        "AA_ANTIGUEDAD": "Z",
        "AA_REFORMA": "AA"
    }

    # Escribir los datos en el archivo
    start_row = 5  # Corresponde a la fila 5 en Excel
    for idx, construccion in enumerate(construcciones_data):
        for campo, col in mapeo_campos_sauce.items():
            valor_campo = construccion[list(mapeo_campos_sauce.keys()).index(campo)]
            ws[col + str(start_row + idx)] = valor_campo

   

# Función para seleccionar el directorio de salida
def seleccionar_directorio():
    directorio = filedialog.askdirectory()
    if directorio:
        entry_output_dir.delete(0, tk.END)
        entry_output_dir.insert(0, directorio)

# Función para seleccionar el archivo de plantilla
def seleccionar_archivo():
    archivo = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])
    if archivo:
        entry_template_file.delete(0, tk.END)
        entry_template_file.insert(0, archivo)

# Función para comparar y resaltar diferencias
def comparar_y_resaltar(ws):
    # Obtener el rango de celdas a comparar (desde A5 hasta la última fila escrita en los datos iniciales)
    max_row = ws.max_row
    max_col_iniciales = ws.max_column
    max_col_sauce = max_col_iniciales + 14  # Ajuste para comenzar desde la columna "O" de SAUCE
    
    for row in range(5, max_row + 1):
        for col in range(1, max_col_iniciales + 1):
            cell_iniciales = ws.cell(row=row, column=col)
            cell_sauce = ws.cell(row=row, column=col + 14)  # Ajuste para alinear con la columna "O" de SAUCE
            
            # Obtener los valores de las celdas como texto
            valor_celda_iniciales = str(cell_iniciales.value)
            valor_celda_sauce = str(cell_sauce.value)
            
            # Comparar los valores de las celdas
            if valor_celda_iniciales != valor_celda_sauce:
                # Resaltar el texto en rojo en la parte de SAUCE
                cell_sauce.font = Font(color="FF0000")  # Rojo




# Función para ejecutar ambos procesos y guardar el archivo Excel
def ejecutar_procesos():
    output_dir = entry_output_dir.get()
    template_file = entry_template_file.get()
    schema = entry_schema.get()
    refcat_value = entry_refcat.get()

    if not output_dir or not template_file or not schema or not refcat_value:
        messagebox.showerror("Error", "Por favor, complete todos los campos")
        return

    # Crear el directorio de salida si no existe
    os.makedirs(output_dir, exist_ok=True)

    # Usar el campo REFCAT como identificador en el nombre del archivo
    file_name = f"FICHA_RESUMEN_PLACO23_{refcat_value}.xlsx"
    file_path = os.path.join(output_dir, file_name)

    # Copiar el archivo de plantilla
    shutil.copyfile(template_file, file_path)

    # Cargar el archivo copiado
    book = load_workbook(file_path)
    ws_sauce = book["SAUCE"]

    # Escribir datos iniciales
    if escribir_datos_iniciales(ws_sauce, schema, refcat_value):
        # Escribir datos de SAUCE
        escribir_datos_sauce(ws_sauce)
        # Comparar y resaltar diferencias
        comparar_y_resaltar(ws_sauce)
        # Guardar el libro de Excel
        book.save(file_path)
        messagebox.showinfo("Éxito", f"Datos escritos exitosamente en el archivo {file_path}")
    else:
        # Eliminar el archivo creado si no se encontraron datos iniciales
        os.remove(file_path)

# Crear la ventana principal
root = tk.Tk()
root.title("Generador de Modelo 6")

# Crear y colocar los elementos de la interfaz gráfica
label_output_dir = tk.Label(root, text="Directorio de Salida:")
label_output_dir.grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
entry_output_dir = tk.Entry(root, width=50)
entry_output_dir.grid(row=0, column=1, padx=10, pady=5)
button_output_dir = tk.Button(root, text="Seleccionar", command=seleccionar_directorio)
button_output_dir.grid(row=0, column=2, padx=10, pady=5)

label_template_file = tk.Label(root, text="Archivo de Plantilla:")
label_template_file.grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)
entry_template_file = tk.Entry(root, width=50)
entry_template_file.grid(row=1, column=1, padx=10, pady=5)
button_template_file = tk.Button(root, text="Seleccionar", command=seleccionar_archivo)
button_template_file.grid(row=1, column=2, padx=10, pady=5)

label_schema = tk.Label(root, text="Esquema:")
label_schema.grid(row=2, column=0, padx=10, pady=5, sticky=tk.W)
entry_schema = tk.Entry(root, width=50)
entry_schema.grid(row=2, column=1, padx=10, pady=5)

label_refcat = tk.Label(root, text="REFCAT:")
label_refcat.grid(row=3, column=0, padx=10, pady=5, sticky=tk.W)
entry_refcat = tk.Entry(root, width=50)
entry_refcat.grid(row=3, column=1, padx=10, pady=5)

button_ejecutar = tk.Button(root, text="Ejecutar", command=ejecutar_procesos)
button_ejecutar.grid(row=4, column=1, padx=10, pady=20)

# Ejecutar la ventana principal
root.mainloop()
