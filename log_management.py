import os

# Borrar registros de archivo
def delete_log(project_address):
    with open(os.path.join(project_address, "log.txt"), "w") as f:
        pass  # Esto borra el archivo (modo 'w' lo trunca)

# Escribir registros en el archivo
def write_log(project_address, text):
    with open(os.path.join(project_address, "log.txt"), "a", encoding="utf-8") as f:
        f.write(f"{text}\n")

# Mostrar registros de archivo
def read_log(project_address, ruta_txt):
    with open(os.path.join(project_address, ruta_txt), 'r', encoding='utf-8') as archivo:
        lineas = [linea.rstrip() for linea in archivo.readlines()]

    if not lineas:
        print("El archivo está vacío.")
        return

    # Calcular el ancho máximo de las líneas
    ancho_max = max(len(linea) for linea in lineas)

    # Ajustar margen adicional si se desea
    margen = 4
    ancho_total = ancho_max + margen

    # Crear marco superior
    print("\n╔" + "═" * ancho_total + "╗")
    print("║{:^{}}║".format(" REGISTROS ", ancho_total))
    print("╠" + "═" * ancho_total + "╣")

    # Imprimir cada línea dentro del marco
    for linea in lineas:
        print("║ {:<{}} ║".format(linea, ancho_total - 2))

    # Crear marco inferior
    print("╚" + "═" * ancho_total + "╝")