import os
import pandas as pd
from reportlab.lib.pagesizes import landscape, letter
from reportlab.platypus import SimpleDocTemplate, Table, PageBreak
from reportlab.lib import colors

def excel_to_pdf(excel_path, pdf_path):
    # Leer el archivo Excel
    df = pd.read_excel(excel_path, sheet_name=None)

    # Crear un documento PDF con orientación apaisada
    pdf = SimpleDocTemplate(pdf_path, pagesize=landscape(letter))
    elements = []

    # Definir las hojas de interés
    sheets_of_interest = ["FICHA RESUMEN PLACO", "SAUCE", "CROQUIS"]

    for sheet_name in sheets_of_interest:
        if sheet_name in df:
            data = df[sheet_name]
            if not data.empty:
                # Crear una tabla y agregarla al documento
                table_data = [data.columns.tolist()] + data.values.tolist()
                table = Table(table_data, repeatRows=1)
                table.setStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ])
                elements.append(table)
                elements.append(PageBreak())

    pdf.build(elements)

def process_directory(directory_path):
    for filename in os.listdir(directory_path):
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            excel_path = os.path.join(directory_path, filename)
            pdf_filename = os.path.splitext(filename)[0] + ".pdf"
            pdf_path = os.path.join(directory_path, pdf_filename)
            excel_to_pdf(excel_path, pdf_path)

# Especifica el directorio que contiene los archivos Excel
directory_path = "modelo6/origen/ficheros/iniciales/salida"

process_directory(directory_path)
