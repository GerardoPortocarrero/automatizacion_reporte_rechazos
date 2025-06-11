import os
from bs4 import BeautifulSoup
from pathlib import Path
import mimetypes
import base64
from email.mime.image import MIMEImage

# Embeber imagenes para que se visualice por correo con "cid"
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

# Funcion para construir descripcion mensual del html
def month_description(soup, month, total_rechazos, months):
    # <p>En el mes de <strong>JUNIO</strong> se reportó un total de <strong style="color: #cc0000;">2278.88 CF RECHAZADAS</strong>. Se muestra la meta por localidad en la última columna.</p>
    p1 = soup.new_tag("p")
    p1.append(BeautifulSoup(f'En el mes de <strong>{months[int(month)].upper()}</strong> se reportó un total de <strong style="color: #cc0000;">{total_rechazos} CF Rechazadas</strong>. Se muestra la meta por localidad en la última columna.', 'html.parser'))

    return p1

# Función para construir un bloque HTML de una sede
def crear_bloque_sede_solo_para_html_normal(sede, soup):
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
    p2.append(BeautifulSoup(f'Total de <strong>{sede["total_rechazos"]}</strong> CF RECH', 'html.parser'))
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

# Función para construir un bloque HTML CID de una sede
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
    p2.append(BeautifulSoup(f'Venta Rechazada: <strong>{sede["total_rechazos"]} CF</strong>.', 'html.parser'))
    bloque.append(p2)

    p3 = soup.new_tag("p")
    p3.append(BeautifulSoup(f'Carga Total: <strong>{sede["total_carga"]} CF</strong>.', 'html.parser'))
    bloque.append(p3)

    bloque.append(soup.new_tag("p", string="Detalle:"))

    div_detalle = soup.new_tag("div", style="text-align: center;")
    img_detalle = soup.new_tag("img", src=f"cid:{sede['detalle']}", width="715", style="margin: 10px; border: 1px solid #ddd; border-radius: 6px;")
    div_detalle.append(img_detalle)
    bloque.append(div_detalle)

    bloque.append(soup.new_tag("p", string="Evolución Rechazo – Día"))

    div_evo = soup.new_tag("div", style="text-align: center;")
    img_evo = soup.new_tag("img", src=f"cid:{sede['evolucion']}", width="715", style="margin: 10px; border: 1px solid #ddd; border-radius: 6px;")
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
def insert_updated_information(sedes, month, year, total_rechazos, months):
    # Leer el HTML desde archivo
    with open("mail_report/template.html", "r", encoding="utf-8") as f:
        html = f.read()

    soup = BeautifulSoup(html, "html.parser")

    # Insertar descripcion mensual
    contenedor = soup.find("div", id="month-description")
    contenedor.append(month_description(soup, month, total_rechazos, months))

    # Insertar los bloques por sede
    contenedor = soup.find("div", id="bloques-sedes")
    for sede in sedes:
        contenedor.append(crear_bloque_sede(sede, soup))

    # Guardar el HTML actualizado
    with open("mail_report/index.html", "w", encoding="utf-8") as f:
        f.write(str(soup))