import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image as PILImage
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import logging

#Script que recorre un directorio padre, en busca de ficheros excel, jpg, png, insertando los últimos dentro de la hoja CROQUIS del excel.

# Configurar logging
logging.basicConfig(filename='discrepancias.log', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Función para redimensionar y guardar las imágenes
def resize_image(image_path, output_path, width, height):
    with PILImage.open(image_path) as img:
        resized_img = img.resize((width, height), PILImage.Resampling.LANCZOS)
        resized_img.save(output_path)

# Función para añadir imágenes al archivo Excel
def add_images_to_excel(excel_path, png_path, jpg_path, desired_width, desired_height):
    resized_png_path = os.path.join(os.path.dirname(png_path), 'resized_' + os.path.basename(png_path))
    resized_jpg_path = os.path.join(os.path.dirname(jpg_path), 'resized_' + os.path.basename(jpg_path))
    
    if os.path.exists(png_path):
        resize_image(png_path, resized_png_path, desired_width, desired_height)
    else:
        print(f"No se encontró la imagen PNG en la ruta: {png_path}")
        return
    
    if os.path.exists(jpg_path):
        resize_image(jpg_path, resized_jpg_path, desired_width, desired_height)
    else:
        print(f"No se encontró la imagen JPG en la ruta: {jpg_path}")
        return
    
    # Abrir el archivo Excel existente
    wb = openpyxl.load_workbook(excel_path)
    
    # Verificar si la hoja "CROQUIS" existe
    if "CROQUIS" not in wb.sheetnames:
        print(f"La hoja 'CROQUIS' no existe en el archivo Excel: {excel_path}")
        return
    
    ws = wb["CROQUIS"]
    
    # Añadir la primera imagen redimensionada (PNG) a la celda C5
    img_png = OpenpyxlImage(resized_png_path)
    img_png.anchor = 'C5'  # Anclar la imagen a la celda C5
    ws.add_image(img_png)
    
    # Añadir la segunda imagen redimensionada (JPG) a la celda R5
    img_jpg = OpenpyxlImage(resized_jpg_path)
    img_jpg.anchor = 'R5'  # Anclar la imagen a la celda R5
    ws.add_image(img_jpg)
    
    # Guardar el libro en la misma ruta sin perder el formato
    wb.save(excel_path)
    print(f"Imágenes redimensionadas insertadas y archivo Excel guardado en: {excel_path}")

# Función para procesar las carpetas en el directorio principal
def process_folders(main_directory, progress, progress_label):
    folders = [folder_name for folder_name in os.listdir(main_directory) if os.path.isdir(os.path.join(main_directory, folder_name))]
    total_folders = len(folders)
    
    for index, folder_name in enumerate(folders):
        folder_path = os.path.join(main_directory, folder_name)
        excel_path = os.path.join(folder_path, f"{folder_name}.xlsx")
        png_path = os.path.join(folder_path, f"{folder_name}.png")
        jpg_path = os.path.join(folder_path, f"{folder_name}.jpg")
        
        # Verificar si los nombres de los archivos coinciden con el nombre de la carpeta
        discrepancias = []
        if not os.path.exists(excel_path):
            discrepancias.append(f"Falta archivo Excel: {excel_path}")
        if not os.path.exists(png_path):
            discrepancias.append(f"Falta archivo PNG: {png_path}")
        if not os.path.exists(jpg_path):
            discrepancias.append(f"Falta archivo JPG: {jpg_path}")
        
        if discrepancias:
            for discrepancia in discrepancias:
                logging.info(f"Discrepancia en carpeta {folder_name}: {discrepancia}")
            messagebox.showwarning("Discrepancia encontrada", f"Se encontraron discrepancias en la carpeta {folder_name}. Verifique el archivo de log para más detalles.")
            continue

        add_images_to_excel(excel_path, png_path, jpg_path, 700, 700)
        
        # Actualizar la barra de progreso
        progress['value'] = (index + 1) / total_folders * 100
        progress_label.config(text=f"Procesando carpeta {index + 1} de {total_folders}")
        root.update_idletasks()

# Función para seleccionar el directorio principal
def select_directory():
    main_directory = filedialog.askdirectory()
    if main_directory:
        process_folders(main_directory, progress, progress_label)
        messagebox.showinfo("Completado", "Se han procesado todas las carpetas.")

# Crear la interfaz de usuario con tkinter
root = tk.Tk()
root.title("Procesador de Imágenes en Excel")
root.geometry("400x200")

frame = tk.Frame(root, padx=20, pady=20)
frame.pack(expand=True, fill=tk.BOTH)

label = tk.Label(frame, text="Seleccione el directorio principal:")
label.pack(pady=(0, 10))

select_button = tk.Button(frame, text="Seleccionar Directorio", command=select_directory)
select_button.pack(pady=(0, 20))

progress_label = tk.Label(frame, text="")
progress_label.pack()

progress = ttk.Progressbar(frame, orient="horizontal", length=300, mode="determinate")
progress.pack(pady=(0, 10))

root.mainloop()
