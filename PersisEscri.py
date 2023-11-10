from docx import Document

# Se especifica la ruta del archivo docx que queremos usar
ruta_archivo_docx = r'C:/Users/Pepe/Documents/CursoFormacionPython/PersistenciaFicheros/EmailPrueba.docx'
doc = Document(ruta_archivo_docx)

# Para agregar una nueva frase al documento
nueva_frase_docx = "Formación de Python. Escritura en archivo .docx"
doc.add_paragraph(nueva_frase_docx)

# Guardar los cambios
doc.save(ruta_archivo_docx)
print(f'Se ha agregado la frase "{nueva_frase_docx}" al archivo docx.')

# Se especifica la ruta del archivo txt que queremos usar
ruta_archivo_txt = r'C:/Users/Pepe/Documents/CursoFormacionPython/PersistenciaFicheros/EmailPrueba.txt'

# Agregar nueva frase en el documento
nueva_frase_txt = "Formación de Python. Escritura en archivo .txt"

# Abre el archivo TXT en modo escritura (agrega contenido al final)
with open(ruta_archivo_txt, 'a', encoding='utf-8') as archivo_txt:
    # Agrega la nueva frase al archivo
    archivo_txt.write(nueva_frase_txt + '\n')

print(f'Se ha agregado la frase "{nueva_frase_txt}" al archivo TXT.')
