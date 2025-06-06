import os
from bs4 import BeautifulSoup
import pandas as pd

# Hacer calculos para los porcentajes y totales
def make_calculatiosn_for_sedes(locaciones, transportista_updated, ruta_updated, transportista, ruta):
    month, year = "6", "2025"
    fecha_inicio = pd.to_datetime('01/'+str(month)+'/'+str(year), format='%d/%m/%Y')
    fecha_fin = pd.to_datetime('01/'+str(int(month)+1)+'/'+str(year), format='%d/%m/%Y')
    ruta = ruta[(ruta[date_string] >= fecha_inicio) & (ruta[date_string] < fecha_fin)]

    for locacion in locaciones:
        # Filtrar el DataFrame por la locación actual
        df_locacion = df_mail[df_mail['Locación'] == locacion]

        # Calcular el porcentaje y total de CF RECH
        if not df_locacion.empty:
            porcentaje = df_locacion['CF RECH'].sum() / df_locacion['CF'].sum() if df_locacion['CF'].sum() > 0 else 0
            total = df_locacion['CF RECH'].sum()
        else:
            porcentaje = 0
            total = 0

        # Actualizar los diccionarios con los resultados
        porcentajes[locacion] = f"{porcentaje:.2%}"
        totales[locacion] = f"{total} CF RECH."

# Funcion para construir los datos de cada bloque
def build_sedes_lista(locaciones):
    # Directorio donde están las imágenes
    image_folder = "mail_report"

    sedes = []

    for loc in locaciones:
        sede = {
            "nombre": loc,
            "porcentaje": '1.0', # porcentajes.get(loc, "0.00%"),
            "total": '2.0', # totales.get(loc, "0.00 CF RECH."),
            "imagenes": [f"RECHAZOS-{loc}-{i}.png" for i in range(1, 4)],
            "detalle": f"DETALLES-{loc}-1.png",
            "evolucion": f"RECHAZOS-{loc}-4.png"
        }
        
        # (Opcional) Verifica si los archivos existen
        missing = [
            fname for fname in sede["imagenes"] + [sede["detalle"], sede["evolucion"]]
            if not os.path.isfile(os.path.join(image_folder, fname))
        ]
        if missing:
            print(f"⚠️ Archivos faltantes para '{loc}': {missing}")

        sedes.append(sede)

    return sedes

# Función para construir un bloque HTML de una sede
def crear_bloque_sede(sede, soup):
    bloque = soup.new_tag("div", style="margin-top: 25px;")

    h3 = soup.new_tag("h3", style="color: #cc0000;")
    h3.string = sede["nombre"]
    bloque.append(h3)

    div_imgs = soup.new_tag("div", style="text-align: center;")
    for img in sede["imagenes"]:
        tag_img = soup.new_tag(
            "img",
            src=img,
            width="230",
            height="175",
            style="margin: 5px; border: 1px solid #ddd; border-radius: 6px;"
        )
        div_imgs.append(tag_img)
    bloque.append(div_imgs)

    p1 = soup.new_tag("p")
    p1.append(BeautifulSoup(f'{sede["nombre"]} alcanzó un <strong style="color: #cc0000;">{sede["porcentaje"]}</strong> de CF RECH.', 'html.parser'))
    bloque.append(p1)

    p2 = soup.new_tag("p")
    p2.append(BeautifulSoup(f'Total: <strong>{sede["total"]}</strong>', 'html.parser'))
    bloque.append(p2)

    bloque.append(soup.new_tag("p", string="Detalle:"))

    div_detalle = soup.new_tag("div", style="text-align: center;")
    img_detalle = soup.new_tag("img", src=sede["detalle"], width="700", style="margin: 10px; border: 1px solid #ddd; border-radius: 6px;")
    div_detalle.append(img_detalle)
    bloque.append(div_detalle)

    bloque.append(soup.new_tag("p", string="Evolución Rechazo – Día"))

    div_evo = soup.new_tag("div", style="text-align: center;")
    img_evo = soup.new_tag("img", src=sede["evolucion"], width="700", style="margin: 10px; border: 1px solid #ddd; border-radius: 6px;")
    div_evo.append(img_evo)
    bloque.append(div_evo)

    bloque.append(soup.new_tag("hr", style="border-top: 1px solid #ccc; margin: 30px 0;"))

    return bloque

# Usar template para rellenar los bloques de sedes
def insert_updated_information(sedes):
    # Leer el HTML desde archivo
    with open("mail_report/template.html", "r", encoding="utf-8") as f:
        html = f.read()

    soup = BeautifulSoup(html, "html.parser")

    # Encuentra el contenedor vacío donde insertar los bloques
    contenedor = soup.find("div", id="bloques-sedes")

    # Insertar cada bloque
    for sede in sedes:
        contenedor.append(crear_bloque_sede(sede, soup))

    # Guardar el HTML actualizado
    with open("mail_report/index.html", "w", encoding="utf-8") as f:
        f.write(str(soup))