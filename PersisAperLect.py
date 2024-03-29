""" 
Curso Python empresa de 'Lenguaje de Programación Python'

Autor: José Antonio Calvo López

Fecha: Noviembre 2023

"""

from docx import Document

# Tratamiento de apertura y de lectura de un archivo .docx
# Se especifica la ruta del archivo docx que queremos usar
# En ruta relativa
ruta_archivo_docx = r'EmailPrueba.docx'

# Creación un objeto Document a partir del archivo DOCX del cual la variable ruta_archivo_docx ya ha cogido el valor
doc = Document(ruta_archivo_docx)

# Bucle for a través de los párrafos del documento para darle el valor a cada uno e imprimir su contenido por consola
for paragraph in doc.paragraphs:
    print(paragraph.text)

# Tratamiento de apertura y de lectura de un archivo .txt

# Se especifica la ruta del archivo txt que queremos usar
# En ruta absoluta
ruta_archivo_txt = r'C:/Users/Pepe/Documents/CursoFormacionPython/PersistenciaFicheros/EmailPrueba.txt'

# Abre el archivo TXT en modo lectura
with open(ruta_archivo_txt, 'r', encoding='utf-8') as archivo_txt:
    contenido = archivo_txt.read()
# El método read() se usa para leer todo el contenido de una vez, también se puede recorrer con un bucle for
# Imprime por consola el contenido del archivo TXT
print(contenido)




