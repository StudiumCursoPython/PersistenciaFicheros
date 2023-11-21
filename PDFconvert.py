import tkinter as tk
from tkinter import filedialog
from tkinter import *
from tkinter.ttk import *
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph
from PIL import Image, ImageTk
import PyPDF2
import os
import sys

def seleccionar_archivos():
    rutas_archivos = filedialog.askopenfilenames(filetypes=[("Documentos de Word o PDF", "*.docx;*.pdf")])
    if rutas_archivos:
        # Si solo se seleccionó un archivo, se procesa directamente.
        if len(rutas_archivos) == 1:
            ruta_archivo = rutas_archivos[0]
            ruta_salida = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
            if ruta_salida:
                if ruta_archivo.lower().endswith(".docx"):
                    convertir_word_a_pdf(ruta_archivo, ruta_salida)
                elif ruta_archivo.lower().endswith(".pdf"):
                    copiar_pdf(ruta_archivo, ruta_salida)
        else:
            opcion_combinar = tk.messagebox.askyesno("Combinar PDFs", "¿Deseas combinar los archivos en un único PDF?")
            if opcion_combinar:
                ruta_salida = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
                if ruta_salida:
                    combinar_archivos(rutas_archivos, ruta_salida)
            else:
                for ruta_archivo in rutas_archivos:
                    ruta_salida = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
                    if ruta_salida:
                        if ruta_archivo.lower().endswith(".docx"):
                            convertir_word_a_pdf(ruta_archivo, ruta_salida)
                        elif ruta_archivo.lower().endswith(".pdf"):
                            copiar_pdf(ruta_archivo, ruta_salida)

def combinar_archivos(rutas_archivos, ruta_salida):
    escritor_pdf = PyPDF2.PdfWriter()

    for ruta_archivo in rutas_archivos:
        if ruta_archivo.lower().endswith(".docx"):
            pdf_temporal = ruta_archivo.replace(".docx", ".pdf")
            convertir_word_a_pdf(ruta_archivo, pdf_temporal)
            lector_pdf = PyPDF2.PdfReader(pdf_temporal)
            os.remove(pdf_temporal)
        else:
            lector_pdf = PyPDF2.PdfReader(ruta_archivo)

        for numero_pagina in range(len(lector_pdf.pages)):
            pagina = lector_pdf.pages[numero_pagina]
            escritor_pdf.add_page(pagina)

    with open(ruta_salida, "wb") as archivo_salida:
        escritor_pdf.write(archivo_salida)

    print("Se ha creado el archivo PDF combinado en:", ruta_salida)

def convertir_word_a_pdf(ruta_entrada, ruta_salida):
    documento = Document(ruta_entrada)
    # Uso de la librería reportlab
    pdf = SimpleDocTemplate(ruta_salida, pagesize=letter)
    estilos = getSampleStyleSheet()
    contenido = []

    for parrafo in documento.paragraphs:
        contenido.append(Paragraph(parrafo.text, estilos["Normal"]))

    pdf.build(contenido)
    print("Se ha creado el archivo PDF en:", ruta_salida)

def copiar_pdf(ruta_entrada, ruta_salida):
    with open(ruta_entrada, "rb") as archivo_entrada:
        lector_pdf = PyPDF2.PdfReader(archivo_entrada)
        escritor_pdf = PyPDF2.PdfWriter()

        for numero_pagina in range(len(lector_pdf.pages)):
            pagina = lector_pdf.pages[numero_pagina]
            escritor_pdf.add_page(pagina)

        with open(ruta_salida, "wb") as archivo_salida:
            escritor_pdf.write(archivo_salida)

    print("Se ha creado el archivo PDF en:", ruta_salida)

# Crear la ventana de la aplicación
aplicacion = tk.Tk()
aplicacion.title("Conversor de Documentos a PDF")
aplicacion.geometry("400x300")  # Tamaño de la ventana
aplicacion.resizable(False, False)

# Obtener la ruta de acceso a los recursos incluidos en el archivo para los archivos empaquetados
ruta_recursos = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))

# Otra forma de obtener los recursos 
# ruta_recursos = getattr(sys, 'os.path.dirname(os.path.abspath(__file__))', os.path.dirname(os.path.abspath(__file__)))

# Cargar las imágenes
icono = PhotoImage(file=os.path.join(ruta_recursos, "Studium.png"))
# Función que construye la ruta completa al archivo de imagen "Studium.png" válida para los sistemas W y Linux

foto = PhotoImage(file=os.path.join(ruta_recursos, "pdfImagen.png"))

# Ajuste del tamaño de la imagen del botón
ancho_imagen = 50
alto_imagen = 50
foto = foto.subsample(foto.width() // ancho_imagen, foto.height() // alto_imagen)

# Crear una etiqueta para la imagen de fondo
imagen_fondo = Image.open(os.path.join(ruta_recursos, "ImgCurso.png"))
# Tamaño para la imagen de fondo
imagen_fondo = imagen_fondo.resize((400, 300))
imagen_fondo = ImageTk.PhotoImage(imagen_fondo)
etiqueta_fondo = Label(aplicacion, image=imagen_fondo)
etiqueta_fondo.place(x=0, y=0, relwidth=1, relheight=1)

# Botón para seleccionar documentos de Word o PDF y personalización del botón
boton_seleccionar = Button(aplicacion, text="Seleccionar documentos de Word o PDF", command=seleccionar_archivos, image=foto, compound=BOTTOM)
boton_seleccionar.pack(pady=110)

# Ejecutar la aplicación
aplicacion.iconphoto(True, icono)
aplicacion.mainloop()
