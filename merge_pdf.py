import os
import logging
from tkinter import Tk, Button, Label, filedialog, messagebox, ttk
from pypdf import PdfMerger

# Configuración del logging
logging.basicConfig(filename='registro_pdf.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def merge_pdfs_by_refcat(directory, progress_bar, total_label, current_label):
    # Crear un diccionario para almacenar listas de archivos por refcat
    refcat_dict = {}

    logging.info(f"Iniciando proceso en el directorio: {directory}")

    # Recorrer los archivos en el directorio
    file_count = 0
    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            file_count += 1
            refcat = filename[:14]
            if refcat not in refcat_dict:
                refcat_dict[refcat] = []
            refcat_dict[refcat].append(filename)

    total_files = file_count // 3  # Dividir el total entre 3
    total_label.config(text=f"Total de archivos: {total_files}")

    # Definir el orden de los sufijos
    order = [
        "_FICHA_RESUMEN_PLACO.pdf",
        "_FichaResumen_SAUCE.pdf",
        "_FichaResumen_CROQUIS.pdf"
    ]

    # Unir los PDFs por refcat y eliminar los originales
    for index, (refcat, pdf_files) in enumerate(refcat_dict.items()):
        # Calcular el progreso
        progress = (index + 1) * 100 / len(refcat_dict)
        progress_bar["value"] = progress
        root.update_idletasks()  # Actualizar la interfaz gráfica

        current_label.config(text=f"Procesando archivo {index+1} de {total_files}")

        # Ordenar los archivos según el orden especificado
        sorted_files = sorted(pdf_files, key=lambda x: order.index(next(filter(lambda suffix: x.endswith(suffix), order), None)))

        merger = PdfMerger()
        for pdf in sorted_files:
            pdf_path = os.path.join(directory, pdf)
            merger.append(pdf_path)
        output_path = os.path.join(directory, f"{refcat}_FichaResumen.pdf")
        merger.write(output_path)
        merger.close()

        logging.info(f"Creado archivo combinado: {output_path}")

        # Eliminar los archivos originales después de combinar
        for pdf in sorted_files:
            pdf_path = os.path.join(directory, pdf)
            os.remove(pdf_path)
            logging.info(f"Eliminado archivo original: {pdf_path}")

    logging.info(f"Proceso completado en el directorio: {directory}")

def select_directory_and_merge():
    directory = filedialog.askdirectory()
    if directory:
        # Configurar la barra de progreso
        progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
        progress_bar.pack(pady=10)

        # Etiquetas para mostrar el número total de archivos y el archivo actual
        total_label = Label(root, text="", font=("Arial", 10))
        total_label.pack(pady=5)

        current_label = Label(root, text="", font=("Arial", 10))
        current_label.pack(pady=5)

        # Fusionar PDFs y actualizar la barra de progreso
        merge_pdfs_by_refcat(directory, progress_bar, total_label, current_label)

        messagebox.showinfo("Proceso Completado", f"PDFs combinados y originales eliminados en {directory}")

        # Eliminar los elementos de la interfaz gráfica después de completar el proceso
        progress_bar.destroy()
        total_label.destroy()
        current_label.destroy()

# Configurar la interfaz gráfica
root = Tk()
root.title("Fusionador de PDFs por Refcat")
root.geometry("400x250")  # Aumentar el alto de la interfaz

label = Label(root, text="Selecciona el directorio con los PDFs a combinar:", font=("Arial", 12))
label.pack(pady=10)

select_button = Button(root, text="Seleccionar Directorio", command=select_directory_and_merge, font=("Arial", 12))
select_button.pack(pady=10)

root.mainloop()
