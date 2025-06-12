import os
import time
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from PIL import Image
from io import BytesIO

# Realiza un click en un botón basado en su aria label
def click_location_button(driver, location_button):
    try:
        print(f"* {location_button} *")        

        # Espera que el DOM esté estable y el botón sea visible
        xpath = f"//*[@aria-label='{location_button}']"
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )

        # Busca el botón justo antes de hacer click
        button = driver.find_element(By.XPATH, xpath)
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", button)
        time.sleep(2)

        # Ejecuta el click
        driver.execute_script("arguments[0].click();", button)
        print(f"[✓] Click realizado en: {location_button}")
        time.sleep(4)

    except Exception as e:
        print(f"[X] No se pudo hacer click en: {location_button}")

# Realiza un click en un botón basado en su aria label (para desclicar el boton previamente clickeado)
def unclick_location_button(driver, location_button):
    try:
        # Espera que el DOM esté estable y el botón sea visible
        xpath = f"//*[@aria-label='{location_button}']"
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )

        # Busca el botón justo antes de hacer unclick
        button = driver.find_element(By.XPATH, xpath)
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", button)
        time.sleep(2)

        # Ejecuta unclick
        driver.execute_script("arguments[0].click();", button)
        print(f"[✓] Unclick realizado en: {location_button}")
        time.sleep(4)

    except Exception as e:
        print(f"[X] No se pudo hacer unclick en: {location_button}")

# Descarga las figuras de un reporte de Power BI
def download_graphics(page_name, page_graphics, driver, locacion=""):
    for id, graphic in page_graphics.items():
        print(f'[*] Buscando el gráfico: {graphic}')
        
        try:
            # Buscar todos los elementos con aria-label dentro de visual-container
            elements_found = driver.find_elements(By.XPATH, f"//visual-container//*[contains(@aria-label, '{graphic}')]")
            
            # Si no se encontraron elementos, lanzar una excepción
            if not elements_found:
                raise Exception(f"[X] No se encontró ningún elemento con el texto: '{graphic}'")

            # Si se encontraron elementos, buscar el primero que contenga visual-container
            for cand in elements_found:
                # Subimos al ancestro visual-container
                try:
                    contenedor = cand.find_element(By.XPATH, "./ancestor::visual-container")
                    grafico = contenedor.find_element(By.CLASS_NAME, "visualContainer")
                    break # No es necesario seguir buscando si ya encontramos uno
                except:
                    continue
            else:
                raise Exception(f"[X] No se encontró visualContainer para: '{graphic}'")

            # Desplazar el gráfico al centro de la pantalla
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", grafico)
            time.sleep(2)

            # Guardar la captura de pantalla
            save_file_as = f'{page_name}{str("-"+locacion) if locacion != "" else locacion}-{id}.png'
            save_file_as = save_file_as.replace(" ", "_")

            # Especificar la ruta de guardado
            path = os.path.join('mail_report', save_file_as)

            # Guardar captura de pantalla
            grafico.screenshot(path)
            print(f'✅ Imagen guardada: {save_file_as}')

        except Exception as e:
            print(f"❌ Error: no se pudo encontrar o capturar: {graphic}")

# Genera un reporte de Power BI por página
def graphics_capture_by_page(locaciones, options, page_report):
    page_name = page_report['page_name']
    page_url = page_report['page_url']
    filter_report_by = page_report['filter_report_by']
    page_graphics = page_report['page_graphics']

    print('\n.-----------------------------------------------------------------------.')

    try:
        driver = webdriver.Chrome(options=options)        
        print(f'[*] Abriendo Reporte de Power BI ({page_name})')

        driver.get(page_url)
        print("[*] Esperando a que cargue la página")

        WebDriverWait(driver, 20) # 20 segundos para que se cargue la pagina
        print("[*] Pagina cargada (9 seg de renderizado) ...")
        time.sleep(15) # 15 segundos adicionales para renderizado de la pagina

        if filter_report_by == 'locacion':
            print('\nBuscar por locacion ...\n')
            for locacion in locaciones:
                click_location_button(driver, locacion)
                download_graphics(page_name, page_graphics, driver, locacion)
                unclick_location_button(driver, locacion)
                print('')
        else:
            print('\nBuscar por mes ...\n')
            download_graphics(page_name, page_graphics, driver)
            print('')

    finally:
        driver.quit()

    print("'-----------------------------------------------------------------------'")

# Captura de graficos de Power Bi por pagina
def take_powerbi_graphics_main(locaciones, PAGES_REPORT):
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--user-data-dir=C:\\Users\\AYACDA23\\AppData\\Local\\Google\\Chrome\\User Data\\Profile 6")

    for page in PAGES_REPORT:
        graphics_capture_by_page(locaciones, options, page)