import os
from bs4 import BeautifulSoup
import pandas as pd
import polars as pl
from pathlib import Path
import mimetypes
from datetime import datetime
import base64
from email.mime.image import MIMEImage

def embedir_imagenes_en_html(soup, mail, ruta_base_imagenes):
    """
    Inserta imágenes embebidas en el HTML de un correo de Outlook (usando cid),
    reemplazando espacios por guiones bajos en los nombres de archivo.

    Args:
        soup (BeautifulSoup): Contenido HTML del correo como objeto BeautifulSoup.
        mail: Objeto mail de Outlook (MailItem).
        ruta_base_imagenes (str): Ruta donde están almacenadas las imágenes referenciadas con cid.
    """
    img_tags = soup.find_all("img", src=True)
    
    for img_tag in img_tags:
        src = img_tag["src"]
        if src.startswith("cid:"):
            raw_cid = src[4:]
            cid = raw_cid.replace(" ", "_")  # solo cambio de espacios por guiones bajos
            full_path = os.path.join(ruta_base_imagenes, cid)

            if not os.path.isfile(full_path):
                print(f"[!] Imagen no encontrada: {full_path}")
                continue

            # Agrega la imagen como attachment embebido
            attachment = mail.Attachments.Add(full_path)
            attachment.PropertyAccessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid
            )

            # Actualiza el src en el HTML
            img_tag["src"] = f"cid:{cid}"

# Hacer calculos para los porcentajes y totales
def make_calculations_for_sedes(locaciones, transportista_updated, ruta_updated, transportista, ruta, month, year):
    calculations = {}
    # Crear fechas de inicio y fin como objetos datetime de Python
    fecha_inicio = datetime.strptime(f"01/{month}/{year}", "%d/%m/%Y")
    # Aseguramos que si month = 12, se pase al siguiente año
    if int(month) == 12:
        fecha_fin = datetime.strptime(f"01/01/{int(year) + 1}", "%d/%m/%Y")
    else:
        fecha_fin = datetime.strptime(f"01/{int(month)+1}/{year}", "%d/%m/%Y")

    # Convertir la columna a fecha
    transportista_updated = transportista_updated.with_columns(
        pl.col(transportista["date"]).str.strptime(pl.Date, "%d/%m/%Y", strict=False)
    )

    ruta_updated = ruta_updated.with_columns(
        pl.col(ruta["date"]).str.strptime(pl.Date, "%d/%m/%Y", strict=False)
    )

    # Filtrado con fechas
    filtering_transportista = transportista_updated.filter(
        (pl.col(transportista["date"]) >= fecha_inicio) &
        (pl.col(transportista["date"]) < fecha_fin)
    )

    filtering_ruta = ruta_updated.filter(
        (pl.col(ruta["date"]) >= fecha_inicio) &
        (pl.col(ruta["date"]) < fecha_fin)
    )

    for locacion in locaciones:
        # Filtrar usando Polars
        df_transportista = filtering_transportista.filter(pl.col("Locación") == locacion)
        df_ruta = filtering_ruta.filter(pl.col("Locación") == locacion)

        # Sumar columnas específicas
        venta_perdida = df_ruta.select(pl.col("Venta Perdida CF").sum()).item()
        total = df_transportista.select(pl.col("Carga Pvta CF").sum()).item()

        # Calcular porcentaje
        porcentaje = (venta_perdida * 100 / total) if venta_perdida > 0 else 0

        # Agregar resultado
        calculations[locacion] = {
            "locacion": locacion,
            "porcentaje": f"{porcentaje:.2f}%",
            "total_cf": venta_perdida,
            "carga_total": total
        }
    
    return calculations

# Funcion para construir descripcion mensual del html
def month_description(soup, month, total, months):
    # <p>En el mes de <strong>JUNIO</strong> se reportó un total de <strong style="color: #cc0000;">2278.88 CF RECHAZADAS</strong>. Se muestra la meta por localidad en la última columna.</p>
    p1 = soup.new_tag("p")
    p1.append(BeautifulSoup(f'En el mes de <strong>{months[int(month)].upper()}</strong> se reportó un total de <strong style="color: #cc0000;">{total} CF Rechazadas</strong>. Se muestra la meta por localidad en la última columna.', 'html.parser'))

    return p1

# Función para construir un bloque HTML de una sede
# def crear_bloque_sede(sede, soup):
#     bloque = soup.new_tag("div", style="margin-top: 25px;")

#     h3 = soup.new_tag("h3", style="color: #cc0000;")
#     h3.string = sede["nombre"]
#     bloque.append(h3)

#     div_imgs = soup.new_tag("div", style="text-align: center;")
#     for img in sede["imagenes"]:
#         tag_img = soup.new_tag(
#             "img",
#             src=img,
#             width="230",
#             height="175",
#             style="margin: 5px; border: 1px solid #ddd; border-radius: 6px;"
#         )
#         div_imgs.append(tag_img)
#     bloque.append(div_imgs)

#     p1 = soup.new_tag("p")
#     p1.append(BeautifulSoup(f'{sede["nombre"]} alcanzó un <strong style="color: #cc0000;">{sede["porcentaje"]}</strong> de CF RECH.', 'html.parser'))
#     bloque.append(p1)

#     p2 = soup.new_tag("p")
#     p2.append(BeautifulSoup(f'Total de <strong>{sede["total"]}</strong> CF RECH', 'html.parser'))
#     bloque.append(p2) 

#     bloque.append(soup.new_tag("p", string="Detalle:"))

#     div_detalle = soup.new_tag("div", style="text-align: center;")
#     img_detalle = soup.new_tag("img", src=sede["detalle"], width="700", style="margin: 10px; border: 1px solid #ddd; border-radius: 6px;")
#     div_detalle.append(img_detalle)
#     bloque.append(div_detalle)

#     bloque.append(soup.new_tag("p", string="Evolución Rechazo – Día"))

#     div_evo = soup.new_tag("div", style="text-align: center;")
#     img_evo = soup.new_tag("img", src=sede["evolucion"], width="700", style="margin: 10px; border: 1px solid #ddd; border-radius: 6px;")
#     div_evo.append(img_evo)
#     bloque.append(div_evo)

#     bloque.append(soup.new_tag("hr", style="border-top: 1px solid #ccc; margin: 30px 0;"))

#     return bloque
def crear_bloque_sede(sede, soup):
    bloque = soup.new_tag("div", style="margin-top: 25px;")

    h3 = soup.new_tag("h3", style="color: #cc0000;")
    h3.string = sede["nombre"]
    bloque.append(h3)

    div_imgs = soup.new_tag("div", style="text-align: center;")
    for img_filename in sede["imagenes"]:
        cid = f"cid:{img_filename}"
        tag_img = soup.new_tag(
            "img",
            src=cid,
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
    p2.append(BeautifulSoup(f'Total de <strong>{sede["total"]}</strong> CF RECH', 'html.parser'))
    bloque.append(p2)

    bloque.append(soup.new_tag("p", string="Detalle:"))

    div_detalle = soup.new_tag("div", style="text-align: center;")
    img_detalle = soup.new_tag("img", src=f"cid:{sede['detalle']}", width="700", style="margin: 10px; border: 1px solid #ddd; border-radius: 6px;")
    div_detalle.append(img_detalle)
    bloque.append(div_detalle)

    bloque.append(soup.new_tag("p", string="Evolución Rechazo – Día"))

    div_evo = soup.new_tag("div", style="text-align: center;")
    img_evo = soup.new_tag("img", src=f"cid:{sede['evolucion']}", width="700", style="margin: 10px; border: 1px solid #ddd; border-radius: 6px;")
    div_evo.append(img_evo)
    bloque.append(div_evo)

    bloque.append(soup.new_tag("hr", style="border-top: 1px solid #ccc; margin: 30px 0;"))

    return bloque


# Funcion para construir los datos de cada bloque
def build_sedes_lista(locaciones, calculations):
    # Directorio donde están las imágenes
    image_folder = "mail_report"

    sedes = []

    for loc in locaciones:
        sede = {
            "nombre": loc,
            "porcentaje": calculations[loc]['porcentaje'], # porcentajes.get(loc, "0.00%"),
            "total": f'{calculations[loc]['total_cf']:.2f}', # totales.get(loc, "0.00 CF RECH."),
            "imagenes": [f"RECHAZOS-{loc}-{i}.png" for i in range(1, 4)],
            "detalle": f"DETALLES-{loc}-1.png",
            "evolucion": f"RECHAZOS-{loc}-4.png"
        }

        sedes.append(sede)

    return sedes

# Usar template para rellenar los bloques de sedes
def insert_updated_information(sedes, month, year, total, months):
    # Leer el HTML desde archivo
    with open("mail_report/template.html", "r", encoding="utf-8") as f:
        html = f.read()

    soup = BeautifulSoup(html, "html.parser")

    # Insertar descripcion mensual
    contenedor = soup.find("div", id="month-description")
    contenedor.append(month_description(soup, month, total, months))

    # Insertar los bloques por sede
    contenedor = soup.find("div", id="bloques-sedes")
    for sede in sedes:
        contenedor.append(crear_bloque_sede(sede, soup))

    # Guardar el HTML actualizado
    with open("mail_report/index.html", "w", encoding="utf-8") as f:
        f.write(str(soup))