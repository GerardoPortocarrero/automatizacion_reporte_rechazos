import os
from bs4 import BeautifulSoup
from pathlib import Path
import mimetypes
import base64
from email.mime.image import MIMEImage

# Funcion para construir el mes y año del reporte
def report_date(soup, month, year, months):
    # <h2 style="color: #ff3c3c; font-size: 20px; margin: 5px 0 0 0; font-weight: normal;">JUNIO 2025</h2>
    h2 = soup.new_tag(
        "h2",
        style="color: #ff3c3c; font-size: 20px; margin: 5px 0 0 0; font-weight: normal;"
    )
    h2.append(BeautifulSoup(f'{months[int(month)].upper()} {year}', 'html.parser'))

    return h2

# Funcion para construir descripcion mensual del html
def month_description(soup, month, total_rechazos, months):
    # <p>En el mes de <strong>JUNIO</strong> se reportó un total de <strong style="color: #cc0000;">2278.88 CF RECHAZADAS</strong>. Se muestra la meta por localidad en la última columna.</p>
    p1 = soup.new_tag("p")
    p1.append(BeautifulSoup(f'En el mes de <strong>{months[int(month)].upper()}</strong> se reportó un total de <strong style="color: #cc0000;">{total_rechazos:.2f} CF Rechazadas</strong>. Se muestra la meta por localidad en la última columna.', 'html.parser'))

    return p1

# Función para construir un bloque HTML CID de una sede
def crear_bloque_sede(sede, soup):
    bloque = soup.new_tag("div", style="margin-top: 35px; font-family: Arial, sans-serif; font-size: 14px; color: #1a1a1a;")

    # Título de sede
    h3 = soup.new_tag("h3", style="color: #cc0000; font-size: 18px; margin-bottom: 12px;")
    h3.string = sede["nombre"]
    bloque.append(h3)

    # Imágenes resumen pequeñas
    div_imgs = soup.new_tag("div", style="text-align: center;")
    for img_filename in sede["imagenes"]:
        cid = f"cid:{img_filename}"
        tag_img = soup.new_tag(
            "img",
            src=cid,
            width="230",
            height="175",
            style="margin: 6px; border: 1px solid #ccc; border-radius: 6px; box-shadow: 1px 1px 4px rgba(0,0,0,0.1); display: inline-block;"
        )
        div_imgs.append(tag_img)
    bloque.append(div_imgs)

    # Texto resumen
    resumen_textos = [
        f'{sede["nombre"]} alcanzó un <strong style="color: #cc0000;">{sede["porcentaje"]}</strong> de CF RECH.',
        f'Venta Rechazada: <strong style="color: #cc0000;">{sede["total_rechazos"]} CF</strong>.',
        f'Carga Total: <strong>{sede["total_carga"]} CF</strong>.'
    ]
    for txt in resumen_textos:
        p = soup.new_tag("p", style="margin: 4px 0;")
        p.append(BeautifulSoup(txt, "html.parser"))
        bloque.append(p)

    # Título detalle
    bloque.append(soup.new_tag("p", string="Detalle:", style="font-weight: bold; margin-top: 18px;"))

    # Imagen de detalle
    div_detalle = soup.new_tag("div", style="text-align: center;")
    img_detalle = soup.new_tag(
        "img", 
        src=f"cid:{sede['detalle']}", 
        width="715", 
        style="margin: 10px auto; border: 1px solid #ccc; border-radius: 6px; box-shadow: 1px 1px 4px rgba(0,0,0,0.1); display: block;")
    div_detalle.append(img_detalle)
    bloque.append(div_detalle)

    # Título evolución
    bloque.append(soup.new_tag("p", string="Evolución Rechazo – Día", style="font-weight: bold; margin-top: 18px;"))

    # Imagen de evolución
    div_evo = soup.new_tag("div", style="text-align: center;")
    img_evo = soup.new_tag(
        "img", 
        src=f"cid:{sede['evolucion']}", 
        width="715", 
        style="margin: 10px auto; border: 1px solid #ccc; border-radius: 6px; box-shadow: 1px 1px 4px rgba(0,0,0,0.1); display: block;")
    div_evo.append(img_evo)
    bloque.append(div_evo)

    # Separador
    bloque.append(soup.new_tag("hr", style="border-top: 1px solid #ccc; margin: 30px 0;"))

    return bloque

# Funcion para construir los datos de cada bloque
def build_sedes_calculations(locaciones, calculations):
    sedes = []

    for loc in locaciones:
        sede = {
            "nombre": loc,
            "porcentaje": calculations[loc]['porcentaje_cf'], # porcentajes.get(loc, "0.00%"),
            "total_rechazos": f'{calculations[loc]['venta_perdida_cf']:.2f}', # venta perdida cf por localidad
            "total_carga": f'{calculations[loc]['carga_cf']:.2f}', # carga cf por localidad
            "imagenes": [f"RECHAZOS-{loc}-{i}.png" for i in range(1, 4)],
            "detalle": f"DETALLES-{loc}-1.png",
            "evolucion": f"RECHAZOS-{loc}-4.png"
        }

        sedes.append(sede)

    return sedes

# Usar template para rellenar los bloques de sedes
def create_html_report_main(mail_report_folder_address, locaciones, calculations, month, year, total_rechazos, months):
    # Leer el HTML desde archivo
    with open(f'{mail_report_folder_address}/template.html', "r", encoding="utf-8") as f:
        html = f.read()

    soup = BeautifulSoup(html, "html.parser")

    # Insertar subtitulo
    contenedor = soup.find("div", id="report-date")
    contenedor.append(report_date(soup, month, year, months))

    # Insertar descripcion mensual
    contenedor = soup.find("div", id="month-description")
    contenedor.append(month_description(soup, month, total_rechazos, months))

    # Insertar los bloques por sede
    contenedor = soup.find("div", id="bloques-sedes")
    for sede in build_sedes_calculations(locaciones, calculations):
        contenedor.append(crear_bloque_sede(sede, soup))

    # Guardar el HTML actualizado
    with open(f'{mail_report_folder_address}/index.html', "w", encoding="utf-8") as f:
        f.write(str(soup))