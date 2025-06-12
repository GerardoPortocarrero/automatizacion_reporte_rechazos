# Borrar registros de archivo
def delete_log():
    with open("log.txt", "w") as f:
        pass  # Esto borra el archivo (modo 'w' lo trunca)

# Escribir registros en el archivo
def write_log(text):
    with open("log.txt", "a", encoding="utf-8") as f:
        f.write(f"{text}\n")