import os
import re
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PyPDF2 import PdfReader, PdfWriter
import pdfplumber


# Función para registrar errores en el log
def log_error(error_message):
    with open("error_log.txt", "a") as log_file:
        log_file.write(f"{error_message}\n")


# Usar WebDriver Manager para gestionar el ChromeDriver automáticamente
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

try:
    # Abrir la página de login
    driver.get("http://gmas.colmayor.edu.co:8080/gmas/Login.gplus")

    # Iniciar sesión
    user_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='user']"))
    )
    user_field.send_keys("")

    pass_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='pass']"))
    )
    pass_field.send_keys("")
    pass_field.send_keys(Keys.RETURN)

    # Navegar por el menú
    sidebar_element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="sidebar-left"]/div[1]/div[2]/i')
        )
    )
    sidebar_element.click()

    sub_menu_element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, "//*[@id='1004199900']/a/i"))
    )
    actions = ActionChains(driver)
    actions.move_to_element(sub_menu_element).click().perform()

    sub_element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, "//*[@id='1004199915']"))
    )
    actions.move_to_element(sub_element).click().perform()

    sub_element_1 = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, "//*[@id='1004200010']"))
    )
    actions.move_to_element(sub_element_1).click().perform()

    # Ruta de la carpeta
    ruta_carpeta = "D:\Pruebas Gmas"
    patron_numeros = re.compile(r"(\d+)-(\d+)")
    patron_nombre = re.compile(r"^(.*?)-")

    # Listar todos los archivos en la carpeta
    archivos = os.listdir(ruta_carpeta)

    # Iterar sobre los archivos y obtener los números y el nombre del archivo
    for archivo in archivos:
        ruta_completa = os.path.join(ruta_carpeta, archivo)

        if os.path.isfile(ruta_completa):
            match_numeros = patron_numeros.search(archivo)
            if match_numeros:
                numero1 = match_numeros.group(1)  # Primer número
                numero2 = match_numeros.group(2)  # Segundo número
                match_nombre = patron_nombre.match(archivo)
                nombre_archivo = (
                    match_nombre.group(1) if match_nombre else "Nombre no encontrado"
                )
                print(f"Nombre del archivo hasta el guion: {nombre_archivo}")
                print(f"Primer número: {numero1}")
                print(f"Segundo número: {numero2}")

                time.sleep(4)

                try:
                    campo_filtro = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located(
                            (By.XPATH, "//*[@id='idExpediente']")
                        )
                    )
                    campo_filtro.send_keys(numero1)
                    button_filtro = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, "//*[@id='filter']"))
                    )
                    button_filtro.click()
                    time.sleep(3)

                    try:
                        elemento = driver.find_element(
                            By.CSS_SELECTOR, "tr.jqgrow.ui-row-ltr"
                        )
                        if elemento:
                            print(f"El expediente {numero1} ya existe.")
                            driver.refresh()
                            continue

                    except:
                        print(
                            f"El expediente {numero1} no se encontró. Creando nuevo expediente..."
                        )

                        button_crear_expediente = WebDriverWait(driver, 10).until(
                            EC.visibility_of_element_located(
                                (By.XPATH, "//*[@id='bttnAddExpediente']/span[2]")
                            )
                        )
                        actions.move_to_element(
                            button_crear_expediente
                        ).click().perform()
                        time.sleep(4)

                        campo_id_expediente = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located(
                                (By.XPATH, "//*[@id='newrow_idExpediente']")
                            )
                        )
                        campo_id_expediente.send_keys(numero1)

                        campo_nombre_expediente = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located(
                                (By.XPATH, "//*[@id='newrow_dsExpediente']")
                            )
                        )
                        campo_nombre_expediente.send_keys(nombre_archivo)

                        campo_tipo_expediente = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located(
                                (By.XPATH, "//*[@id='newrow_dsExpedienteTipo']")
                            )
                        )
                        campo_tipo_expediente.click()

                        campo_tipo_expediente_egreso = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located(
                                (
                                    By.XPATH,
                                    "//*[@id='newrow_dsExpedienteTipo']/option[2]",
                                )
                            )
                        )
                        campo_tipo_expediente_egreso.click()

                        campo_dependencia = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located(
                                (By.XPATH, "//*[@id='newrow_dsEstructuraOrg']")
                            )
                        )
                        campo_dependencia.click()

                        campo_dependencia_tesoreria = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located(
                                (
                                    By.XPATH,
                                    "//*[@id='newrow_dsEstructuraOrg']/option[2]",
                                )
                            )
                        )
                        campo_dependencia_tesoreria.click()

                        button_guardar = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located(
                                (By.XPATH, "//*[@id='newrow']/td[16]/a[1]/img")
                            )
                        )
                        button_guardar.click()

                        time.sleep(4)

                        codigo_expediente = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located(
                                (By.XPATH, "//td[@aria-describedby='list1_codigo']")
                            )
                        )
                        codigo_texto_expediente = codigo_expediente.text
                        print(f"Código: {codigo_texto_expediente}")

                        button_visualizar = driver.find_element(
                            By.CSS_SELECTOR,
                            "img[src='static/images/view.png'][title='Visualizar/Gestionar Anexos']",
                        )
                        button_visualizar.click()

                        time.sleep(4)

                        input_file = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.NAME, "file"))
                        )

                        input_file.send_keys(ruta_completa)
                        time.sleep(5)

                        carpeta_destino = os.path.join(
                            ruta_carpeta, f"{numero1}-{numero2}"
                        )
                        if not os.path.exists(carpeta_destino):
                            os.makedirs(carpeta_destino)

                        output_pdf_path = os.path.join(
                            carpeta_destino, f"{codigo_texto_expediente}.pdf"
                        )
                        input_pdf_path = "Ref_Cruzada_Fisica.pdf"
                        temp_pdf_path = "resultado.pdf"

                        # Crear un PDF temporal con la dirección
                        c = canvas.Canvas(temp_pdf_path, pagesize=letter)
                        c.setFont("Helvetica-Bold", 5.5)
                        c.setFillColorRGB(1, 0, 0)
                        x_position = 358
                        y_position = 514.5
                        c.drawString(x_position, y_position, codigo_texto_expediente)
                        c.save()

                        # Leer el PDF original y el temporal
                        pdf_reader = PdfReader(input_pdf_path)
                        pdf_writer = PdfWriter()
                        overlay_pdf = PdfReader(temp_pdf_path)

                        for page_num in range(len(pdf_reader.pages)):
                            page = pdf_reader.pages[page_num]
                            if page_num == 0:
                                page.merge_page(overlay_pdf.pages[0])
                            pdf_writer.add_page(page)

                        with open(output_pdf_path, "wb") as output_pdf_file:
                            pdf_writer.write(output_pdf_file)

                        os.remove(temp_pdf_path)
                        print(
                            f"El PDF modificado ha sido guardado en {output_pdf_path}"
                        )
                        time.sleep(5)

                        meses = {
                            "enero": "01",
                            "febrero": "02",
                            "marzo": "03",
                            "abril": "04",
                            "mayo": "05",
                            "junio": "06",
                            "julio": "07",
                            "agosto": "08",
                            "septiembre": "09",
                            "octubre": "10",
                            "noviembre": "11",
                            "diciembre": "12",
                        }

                        # Proceso para obtener la fecha
                        with pdfplumber.open(ruta_completa) as pdf:
                            texto = ""
                            for pagina in pdf.pages:
                                texto += pagina.extract_text() + "\n"

                        for mes, numero_mes in meses.items():
                            if mes in texto:
                                fecha = re.search(
                                    r"(\d{1,2})\s*de\s*{}\s*(\d{4})".format(mes), texto
                                )
                                if fecha:
                                    print(f"Fecha encontrada: {fecha.group(0)}")
                                    break

                except Exception as e:
                    log_error(f"Error al procesar el archivo {archivo}: {str(e)}")
                    print(
                        f"Ocurrió un error al procesar el archivo {archivo}. Detalles en error_log.txt."
                    )

except Exception as e:
    log_error(f"Error general: {str(e)}")
    print("Ocurrió un error general. Detalles en error_log.txt.")

finally:
    driver.quit()
