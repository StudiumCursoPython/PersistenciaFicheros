from docx import Document

nombre_original = "Antonio"
nombre_cambio_docx = input("Nuevo Nombre: ")

# Se especifica la ruta del archivo docx que queremos usar
ruta_archivo_docx = r'C:/Users/clja1/Documents/CursoFormacionPyhton/PersistenciaFicheros/EmailPrueba.docx'
doc = Document(ruta_archivo_docx)

# Recorrer los párrafos del documento y reemplazar el nombre original
for paragraph in doc.paragraphs:
    for run in paragraph.runs:
        run.text = run.text.replace(nombre_original, nombre_cambio_docx)

# Se guarda el documento con el nombre actualizado
doc.save(ruta_archivo_docx)

print("El nombre ha sido cambiado en el archivo DOCX.")


