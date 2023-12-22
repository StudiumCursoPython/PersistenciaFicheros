""" 
Curso Python empresa de 'Lenguaje de Programación Python'

Autor: José Antonio Calvo López

Fecha: Noviembre 2023

"""

import tkinter as tk
from tkinter import filedialog
from tkinter import *
from tkinter.ttk import *
from docx import Document
from PIL import Image, ImageTk
import PyPDF2
import os
import sys

def seleccionar_archivos():
    rutas_archivos = filedialog.askopenfilenames(filetypes=[("Documentos PDF", "*.pdf")])
    for ruta_pdf in rutas_archivos:
        ruta_docx = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Documento de Word", "*.docx")])
        if ruta_docx:
            convertir_pdf_a_word(ruta_pdf, ruta_docx)

def convertir_pdf_a_word(ruta_pdf, ruta_docx):
    lector_pdf = PyPDF2.PdfReader(ruta_pdf)
    documento_word = Document()

    for pagina in lector_pdf.pages:
        texto = pagina.extract_text()
        if texto:
            documento_word.add_paragraph(texto)

    documento_word.save(ruta_docx)
    print("Se ha creado el archivo .docx en:", ruta_docx)

# Crear la ventana de la aplicación
aplicacion = tk.Tk()
aplicacion.title("Conversor de de PDF a Docx")
aplicacion.geometry("400x300")  # Tamaño de la ventana
aplicacion.resizable(False, False)

# Obtener la ruta de acceso a los recursos incluidos en el archivo
ruta_recursos = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))

# Crear la ventana de la aplicación
icono = PhotoImage(file=os.path.join(ruta_recursos, "Studium.png"))
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
boton_seleccionar = Button(aplicacion, text="Selección documento PDF", command=seleccionar_archivos, image=foto, compound=BOTTOM)
boton_seleccionar.pack(pady=110)

# Ejecutar la aplicación
aplicacion.iconphoto(True, icono)

# Ejecutar la aplicación
aplicacion.mainloop()