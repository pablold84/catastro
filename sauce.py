import pandas as pd

# Especificar la ruta del archivo CSV
archivo_csv = 'modelo6/origen/ficheros/sauce/33_217210_24.CSV'

# Leer el archivo CSV
df = pd.read_csv(archivo_csv, sep=';', skiprows=1, nrows=1, usecols=['VPD', 'EJERCICIO', 'NUMERO', 'CONTROL'])

# Mostrar el DataFrame resultante
print(df)
