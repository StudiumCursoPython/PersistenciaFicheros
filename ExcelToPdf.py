""" 
Curso Python empresa de 'Lenguaje de Programación Python'

Autor: José Antonio Calvo López

Fecha: Noviembre 2023

"""

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import pandas as pd

def excel_to_pdf(excel_file, pdf_file):
    # Lectura de datos de Excel
    df = pd.read_excel(excel_file)

    # Creación de archivo PDF
    c = canvas.Canvas(pdf_file, pagesize=letter)
    width, height = letter

    # Definir algunos estilos básicos
    c.setFont("Helvetica", 10)
    line_height = 14

    # Agregar los datos del DataFrame a la página PDF
    #Para el margen izquierdo
    x_offset = 50
    #Para el margen superior
    y_offset = height - 40  

    # Se dubujan las cabeceras de la columnas
    for i, col in enumerate(df.columns):
        c.drawString(x_offset + i * 100, y_offset, col)

    # Dibujar las filas de datos, no interesa usar los índices (idx)

    for idx, row in df.iterrows():
        y_offset -= line_height
        for i, value in enumerate(row):
            c.drawString(x_offset + i * 100, y_offset, str(value))

        # Añadir una nueva página si se acaba el espacio
        if y_offset < 50:
            c.showPage()
            y_offset = height - 40
            c.setFont("Helvetica", 10)

    # Guardar el PDF
    c.save()

# Uso de la función
excel_to_pdf('Peliculas.xlsx', 'salidaPeliculas.pdf')

