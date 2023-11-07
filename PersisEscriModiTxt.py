#Importante la codificación (tildes #,?) y que sea la misma que está guardada UTF-8

nombre_original = "Antonio"
nombre_cambio_txt = input("Nuevo Nombre: ")

# Se especifica la ruta del archivo txt que queremos usar
ruta_archivo_txt = r'C:/Users/clja1/Documents/CursoFormacionPyhton/PersistenciaFicheros/EmailPrueba.txt'

# Leer el contenido del archivo
with open(ruta_archivo_txt, 'r', encoding='utf-8') as file:
    contenido = file.read()

# con el método replace el nombre original se sustituye por el nuevo
contenido_actualizado = contenido.replace(nombre_original, nombre_cambio_txt)

# Guardar el contenido actualizado en el archivo
with open(ruta_archivo_txt, 'w', encoding='utf-8') as file:
    file.write(contenido_actualizado)

print("El nombre ha sido cambiado en el archivo TXT.")



