import pandas as pd
import io

# Definir el nombre del archivo CSV
archivo_csv = "modelo6/origen/ficheros/sauce/33_217210_24.csv"

# Diccionario para almacenar las líneas de cada sección
secciones = {"FINCAS": [], "EXPEDIENTE": [], "CONSTRUCCIONES":[], "UNIDADES CONSTRUCTIVAS":[]}

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
        campos_unidad = unidad.split(";") #asigno si coincide refCat y unidad constructiva
        if pcatastral1 == campos_unidad[2] and pcatastral2 == campos_unidad[3] and unidad_const_dest == campos_unidad[4]:
            coef_conservacion = campos_unidad[16]
            break

    construcciones_data.append([cargo, orden, escalera, planta, puerta, tipologia, categ_pred, destino, superficie_total, unidad_const_dest, coef_conservacion, aa_antiguedad, aa_reforma])

# Imprimir los resultados combinados
for construccion in construcciones_data:
    print("\t".join(construccion))