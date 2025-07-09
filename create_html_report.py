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
    table_imgs = soup.new_tag("table", attrs={
        "width": "100%",
        "cellpadding": "0",
        "cellspacing": "0",
        "border": "0",
        "style": """
            margin: 0 auto;
            text-align: center;
            border-collapse: collapse;
        """
    })

    row = soup.new_tag("tr")

    for img_filename in sede["imagenes"]:
        td = soup.new_tag("td", style="""
            padding: 8px;
            text-align: center;
            vertical-align: top;
            width: 33.33%;
        """)

        cid = f"cid:{img_filename}"
        img = soup.new_tag("img", src=cid, style="""
            width: 100%;
            max-width: 100%;
            height: auto;
            border: 1px solid #ccc;
            border-radius: 6px;
            box-shadow: 1px 1px 4px rgba(0,0,0,0.1);
            display: block;
            margin: 0 auto;
        """)

        td.append(img)
        row.append(td)

    table_imgs.append(row)
    bloque.append(table_imgs)

    # Texto resumen
    resumen_html = f"""
    <div style="width: 100%;">
        <table cellpadding="0" cellspacing="0" border="0" width="100%" style="
            margin: 0 auto;
            font-family: Arial, sans-serif;
            background-color: #ffffff;
            border-radius: 12px;
            box-shadow: 0 0 14px rgba(0, 0, 0, 0.08);
            text-align: center;
        ">
            <tr>
                <td style="padding: 0px 10px;>
                    <table width="100%" cellpadding="0" cellspacing="0" style="table-layout: fixed;">
                        <tr>
                            <!-- % CF Rechazada + Carga Total CF -->
                            <td style="padding: 18px 12px;">
                                <div style="color: #666; font-weight: bold; font-size: 20px; line-height: 1.4;">
                                    % CF Rechazada
                                </div>
                                <div style="color: #cc0000; font-size: 32px; font-weight: bold; line-height: 1.3; margin-top: 6px;">
                                    {sede["porcentaje"]}
                                </div>
                            </td>
                            <td style="padding: 18px 12px;">
                                <div style="color: #666; font-weight: bold; font-size: 20px; line-height: 1.4;">
                                    Carga Total CF
                                </div>
                                <div style="color: #2c3e50; font-size: 32px; font-weight: bold; line-height: 1.3; margin-top: 6px;">
                                    {sede["total_carga_cf"]}
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <!-- Venta Rechazada + Carga Total CU -->
                            <td style="padding: 18px 12px;">
                                <div style="color: #666; font-weight: bold; font-size: 20px; line-height: 1.4;">
                                    Venta Rechazada CF
                                </div>
                                <div style="color: #e74c3c; font-size: 32px; font-weight: bold; line-height: 1.3; margin-top: 6px;">
                                    {sede["total_rechazos"]}
                                </div>
                            </td>
                            <td style="padding: 18px 12px;">
                                <div style="color: #666; font-weight: bold; font-size: 20px; line-height: 1.4;">
                                    Carga Total CU
                                </div>
                                <div style="color: #2c3e50; font-size: 32px; font-weight: bold; line-height: 1.3; margin-top: 6px;">
                                    {sede["total_carga_cu"]}
                                </div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </div>
    """

    # Luego insertas esto en tu HTML con BeautifulSoup
    bloque.append(BeautifulSoup(resumen_html, "html.parser"))

    # Título detalle
    bloque.append(soup.new_tag("p", string="Detalle:", style="font-weight: bold; margin-top: 18px;"))

    # Imagen de detalle
    div_detalle = soup.new_tag("div", style="text-align: center;")
    img_detalle = soup.new_tag(
        "img",
        src=f"cid:{sede['detalle']}",
        style="""
            display: block;
            margin: 10px auto;
            width: 100%;
            max-width: 1010px;
            height: auto;
            border: 1px solid #ccc;
            border-radius: 6px;
            box-shadow: 1px 1px 4px rgba(0,0,0,0.1);
        """
    )
    div_detalle.append(img_detalle)
    bloque.append(div_detalle)

    # Título evolución
    bloque.append(soup.new_tag("p", string="Evolución Rechazo – Día", style="font-weight: bold; margin-top: 18px;"))

    # Imagen de evolución
    div_evo = soup.new_tag("div", style="text-align: center;")
    img_evo = soup.new_tag(
        "img", 
        src=f"cid:{sede['evolucion']}", 
        style="""
            display: block;
            margin: 10px auto;
            width: 100%;
            max-width: 1010px;
            height: auto;
            border: 1px solid #ccc;
            border-radius: 6px;
            box-shadow: 1px 1px 4px rgba(0,0,0,0.1);
        """
    )
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
            "total_carga_cf": f'{calculations[loc]['carga_cf']:.2f}', # carga cf por localidad
            "total_carga_cu": f'{calculations[loc]['carga_cu']:.2f}', # carga cf por localidad
            "imagenes": [f"RECHAZOS-{loc}-{i}.png" for i in range(1, 4)],
            "detalle": f"DETALLES-{loc}-1.png",
            "evolucion": f"RECHAZOS-{loc}-4.png"
        }

        sedes.append(sede)

    return sedes

# Usar template para rellenar los bloques de sedes
def create_html_report_main(mail_report_folder_address, locaciones, calculations, month, year, total_rechazos, months):
    # Leer el HTML desde archivo
    with open(os.path.join(mail_report_folder_address, 'template.html'), "r", encoding="utf-8") as f:
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
    with open(os.path.join(mail_report_folder_address, 'index.html'), "w", encoding="utf-8") as f:
        f.write(str(soup))