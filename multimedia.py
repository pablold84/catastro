import openpyxl
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
import os

# Nombre del archivo Excel (sin extensión)
file_name = '001917000UN58H'

# Ruta del archivo Excel existente
excel_path = os.path.abspath(os.path.join('modelo6/origen/ficheros/multimedia', f"{file_name}.xlsx"))

# Rutas de las imágenes
image_dir = os.path.abspath('modelo6/origen/ficheros/multimedia/multimedia')
image_png_path = os.path.join(image_dir, f"{file_name}.png")
image_jpg_path = os.path.join(image_dir, f"{file_name}.jpg")

# Dimensiones deseadas para las imágenes (en píxeles)
desired_width = 700
desired_height = 700

# Función para redimensionar y guardar las imágenes
def resize_image(image_path, output_path, width, height):
    with PILImage.open(image_path) as img:
        resized_img = img.resize((width, height), PILImage.Resampling.LANCZOS)
        resized_img.save(output_path)

# Verificar que los archivos existen y redimensionarlos
resized_png_path = os.path.join(image_dir, f"{file_name}_resized.png")
resized_jpg_path = os.path.join(image_dir, f"{file_name}_resized.jpg")

if os.path.exists(image_png_path):
    resize_image(image_png_path, resized_png_path, desired_width, desired_height)
else:
    print(f"No se encontró la imagen PNG en la ruta: {image_png_path}")

if os.path.exists(image_jpg_path):
    resize_image(image_jpg_path, resized_jpg_path, desired_width, desired_height)
else:
    print(f"No se encontró la imagen JPG en la ruta: {image_jpg_path}")

# Si ambos archivos redimensionados existen, continuar con el proceso
if os.path.exists(resized_png_path) and os.path.exists(resized_jpg_path):
    # Abrir el archivo Excel existente
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    # Añadir la primera imagen redimensionada (PNG) a la celda C5
    img_png = Image(resized_png_path)
    img_png.anchor = 'C5'  # Anclar la imagen a la celda C5
    ws.add_image(img_png)

    # Añadir la segunda imagen redimensionada (JPG) a la celda R5
    img_jpg = Image(resized_jpg_path)
    img_jpg.anchor = 'R5'  # Anclar la imagen a la celda R5
    ws.add_image(img_jpg)

    # Guardar el libro en la misma ruta sin perder el formato
    wb.save(excel_path)
    print(f"Imágenes redimensionadas insertadas y archivo Excel guardado en: {excel_path}")
else:
    print("No se pudo continuar, una o ambas imágenes no se encontraron.")
