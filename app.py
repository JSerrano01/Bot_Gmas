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
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PyPDF2 import PdfReader, PdfWriter
import os
import pdfplumber
import tkinter as tk
from tkinter import messagebox  # Importar el submódulo messagebox


# Configuración del log normal
log_file_path = "log.txt"
log_file = open(log_file_path, "w")  # Abrir el archivo en modo de escritura

log_file_path_1 = "log_error.txt"
log_file_error = open(log_file_path_1, "w")  # Abrir el archivo en modo de escritura

# Usar WebDriver Manager para gestionar el ChromeDriver automáticamente
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

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
    EC.element_to_be_clickable((By.XPATH, '//*[@id="sidebar-left"]/div[1]/div[2]/i'))
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
ruta_carpeta = "D:\Bot Gmas\Anexos Comprobantes"

# Expresión regular para extraer los números
patron_numeros = re.compile(r"(\d+)-(\d+)")

# Expresión regular para extraer el nombre del archivo hasta el primer guion
patron_nombre = re.compile(r"^(.*?)-")

# Listar todos los archivos en la carpeta
archivos = os.listdir(ruta_carpeta)

# Iterar sobre los archivos y obtener los números y el nombre del archivo
for archivo in archivos:
    ruta_completa = os.path.join(ruta_carpeta, archivo)

    if os.path.isfile(ruta_completa):
        # Primera pasada: extraer números
        match_numeros = patron_numeros.search(archivo)

        if match_numeros:
            numero1 = match_numeros.group(1)  # Primer número
            numero2 = match_numeros.group(2)  # Segundo número

            # Segunda pasada: obtener el nombre completo hasta el primer guion
            match_nombre = patron_nombre.match(archivo)
            nombre_archivo = (
                match_nombre.group(1) if match_nombre else "Nombre no encontrado"
            )

            # Mostrar los resultados
            print(f"Nombre del archivo hasta el guion: {nombre_archivo}")
            print(f"Primer número: {numero1}")
            print(f"Segundo número: {numero2}")
            log_file.write(f"Nombre del archivo: {nombre_archivo} \n")
            log_file.write(f"Primer número: {numero1} \n")
            log_file.write(f"Segundo número: {numero2} \n")

            # Esperar hasta que el menú se despliegue completamente
            time.sleep(4)

            # Encontrar el campo de expediente
            campo_filtro = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//*[@id='idExpediente']"))
            )
            campo_filtro.send_keys(numero1)

            button_filtro = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//*[@id='filter']"))
            )
            button_filtro.click()

            time.sleep(3)

            try:
                # Intentar encontrar el elemento <tr> por su clase
                elemento = driver.find_element(By.CSS_SELECTOR, "tr.jqgrow.ui-row-ltr")
                if elemento:
                    print(f"El expediente {numero1} ya existe.")
                    log_file.write(f"El expediente {numero1} ya existe. \n")
                    log_file_error.write(f"El expediente {numero1} ya existe. \n")
                    driver.refresh()  # Recargar la página
                    continue  # Pasar al siguiente archivo

            except:
                print(
                    f"El expediente {numero1} no se encontró. Creando nuevo expediente..."
                )
                log_file.write(
                    f"El expediente {numero1} no se encontró. Creando nuevo expediente... \n"
                )

                # Crear un nuevo expediente
                button_crear_expediente = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located(
                        (By.XPATH, "//*[@id='bttnAddExpediente']/span[2]")
                    )
                )
                actions.move_to_element(button_crear_expediente).click().perform()

                time.sleep(4)

                # Crear ID expediente
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

                # Tipo Expediente
                campo_tipo_expediente = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//*[@id='newrow_dsExpedienteTipo']")
                    )
                )
                campo_tipo_expediente.click()

                # Opcion expediente
                campo_tipo_expediente_egreso = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//*[@id='newrow_dsExpedienteTipo']/option[2]")
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
                        (By.XPATH, "//*[@id='newrow_dsEstructuraOrg']/option[2]")
                    )
                )
                campo_dependencia_tesoreria.click()

                # Guardar, Crear Expediente
                button_guardar = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//*[@id='newrow']/td[16]/a[1]/img")
                    )
                )
                button_guardar.click()

                time.sleep(4)

                # Extraer numero de expediente creado
                codigo_expediente = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//td[@aria-describedby='list1_codigo']")
                    )
                )

                # Extraer el texto
                codigo_texto_expediente = codigo_expediente.text

                print(f"Código: {codigo_texto_expediente}")
                log_file.write(f"Código: {codigo_texto_expediente} \n")

                # Encuentra el elemento usando un CSS Selector que combine el 'src' y 'title'
                button_visualizar = driver.find_element(
                    By.CSS_SELECTOR,
                    "img[src='static/images/view.png'][title='Visualizar/Gestionar Anexos']",
                )

                button_visualizar.click()

                time.sleep(4)

                # Ruta completa del archivo a subir
                ruta_archivo = ruta_completa

                # Esperar a que el input esté presente y visible
                input_file = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.NAME, "file"))
                )

                # Subir el archivo
                input_file.send_keys(ruta_archivo)

                time.sleep(5)

                # CREAR CARPETA CON NOMBRE DE RUTA DE PDF

                carpeta_destino = os.path.join("D:\Bot Gmas\Realizados" , f"{numero1}-{numero2}")
                if not os.path.exists(carpeta_destino):
                    os.makedirs(carpeta_destino)

                # Ruta del archivo de salida dentro de la nueva carpeta
                output_pdf_path = os.path.join(carpeta_destino, f"{codigo_texto_expediente}.pdf")

                # Ruta del archivo de entrada y temporal
                input_pdf_path = "D:\Bot Gmas\Plantillas Cruzadas\Ref_Cruzada_Fisica.pdf"
                temp_pdf_path = "resultado.pdf"

                # Definir la dirección como el código del expediente
                direccion = codigo_texto_expediente

                # Crear un PDF temporal con la dirección
                c = canvas.Canvas(temp_pdf_path, pagesize=letter)

                # Ajustar la fuente y color
                c.setFont("Helvetica-Bold", 5.5)  # Cambiar a Helvetica-Bold para texto en negrilla
                c.setFillColorRGB(1, 0, 0)  # Establecer el color rojo

                # Calcular las coordenadas para que el texto quede al lado de "SISTEMA G+"
                x_position = 358  # Ajusta este valor según sea necesario
                y_position = 514.5  # Ajusta este valor según sea necesario

                # Dibujar el texto en la posición calculada
                c.drawString(x_position, y_position, direccion)
                c.save()

                # Leer el PDF original y el temporal
                pdf_reader = PdfReader(input_pdf_path)
                pdf_writer = PdfWriter()
                overlay_pdf = PdfReader(temp_pdf_path)

                # Añadir el contenido del archivo temporal al PDF original
                for page_num in range(len(pdf_reader.pages)):
                    page = pdf_reader.pages[page_num]
                    if page_num == 0:  # Modificar solo la primera página
                        page.merge_page(overlay_pdf.pages[0])
                    pdf_writer.add_page(page)

                # Guardar el nuevo PDF en la carpeta correspondiente
                with open(output_pdf_path, "wb") as output_pdf_file:
                    pdf_writer.write(output_pdf_file)

                # Limpiar el archivo temporal
                os.remove(temp_pdf_path)

                print(f"El PDF modificado ha sido guardado en {output_pdf_path}")
                log_file.write(
                    f"El PDF modificado ha sido guardado en {output_pdf_path} \n"
                )

                time.sleep(5)

                # Analizar el archivo original (ruta_completa) para encontrar la fecha
                # Mapeo de meses
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

                try:
                    with pdfplumber.open(ruta_completa) as pdf:
                        first_page = pdf.pages[0]
                        text = first_page.extract_text()
                    date_match = re.search(r"(\d{1,2}) de (\w+) de (\d{4})", text)

                    if date_match:
                        dia = date_match.group(1)
                        mes = date_match.group(2).lower()  # Convertir a minúsculas para la comparación
                        anio = date_match.group(3)

                        # Obtener el número del mes
                        mes_numero = meses.get(mes)
                        if mes_numero:
                            fecha_formateada = f"{dia}/{mes_numero}/{anio}"
                            print(f"Fecha encontrada y formateada: {fecha_formateada}")
                            log_file.write(
                                f"Fecha encontrada y formateada: {fecha_formateada} \n"
                            )
                        else:
                            print(f"ERROR. Mes '{mes}' no reconocido.")
                            log_file.write(f"ERROR. Mes '{mes}' no reconocido. \n")
                            log_file_error.write(f"ERROR. Mes '{mes}' no reconocido. \n")
                    else:
                        print(f"ERROR. No se encontró una fecha en el archivo {ruta_completa}.")
                        log_file.write(
                            f"ERROR. No se encontró una fecha en el archivo {ruta_completa}. \n"
                        )
                        log_file_error.write(
                            f"ERROR. No se encontró una fecha en el archivo {ruta_completa}. \n"
                        )
                        log_file.write("\n")
                        log_file.write("Continuando con el siguiente archivo...")
                        log_file.write("\n")
                        log_file_error.write("\n")
                        log_file_error.write("Continuando con el siguiente archivo...")
                        log_file_error.write("\n")
                        driver.refresh()
                        continue  # Pasar al siguiente archivo si no se encuentra la fecha

                except Exception as e:
                    print(f"ERROR al leer el archivo {ruta_completa}: {str(e)}")
                    log_file.write(
                        f"ERROR al leer el archivo {ruta_completa}: {str(e)} \n")
                    log_file.write("\n")
                    log_file.write("Continuando con el siguiente archivo...")
                    log_file.write("\n")
                    log_file_error.write(
                        f"ERROR al leer el archivo {ruta_completa}: {str(e)} \n"
                    )
                    log_file_error.write("\n")
                    log_file_error.write("Continuando con el siguiente archivo...")
                    log_file_error.write("\n")
                    driver.refresh()
                    continue  # Pasar al siguiente archivo si no se encuentra la fecha

                # Creacion del segundo PDF de referencia cruzada electronica

                # Ruta del archivo original y el PDF a crear
                input_pdf_path = "D:\Bot Gmas\Plantillas Cruzadas\Ref_Cruzada_Electronica.pdf"  # PDF original
                output_pdf_path = os.path.join(
                    carpeta_destino,
                    f"REFERENCIA CRUZADA COMPROBANTE {numero1} ELECTRONICA.pdf",
                )  # Modificar el nombre según sea necesario

                # Direcciones y sus coordenadas específicas
                direcciones = [
                    (fecha_formateada, 351, 625),  # Usa la fecha encontrada
                    (numero1, 367, 597.4),  # Dirección 2 número 1
                    (numero2, 440, 597.4),  # Dirección 3 número 2
                ]

                # Crear un PDF temporal con las direcciones
                temp_pdf_path = "resultado.pdf"
                c = canvas.Canvas(temp_pdf_path, pagesize=letter)

                # Ajustar la fuente y color
                c.setFont("Helvetica-Bold", 5.5)
                c.setFillColorRGB(1, 0, 0)

                # Dibujar cada dirección en la posición especificada
                for direccion, x_position, y_position in direcciones:
                    c.drawString(x_position, y_position, direccion)

                c.save()

                # Leer el PDF original y el temporal
                pdf_reader = PdfReader(input_pdf_path)
                pdf_writer = PdfWriter()
                overlay_pdf = PdfReader(temp_pdf_path)

                # Añadir el contenido del archivo temporal al PDF original
                for page_num in range(len(pdf_reader.pages)):
                    page = pdf_reader.pages[page_num]
                    if page_num == 0:  # Modificar solo la primera página
                        page.merge_page(overlay_pdf.pages[0])
                    pdf_writer.add_page(page)

                # Guardar el nuevo PDF en la carpeta correspondiente
                with open(output_pdf_path, "wb") as output_pdf_file:
                    pdf_writer.write(output_pdf_file)

                # Limpiar el archivo temporal
                os.remove(temp_pdf_path)

                print(f"El PDF modificado ha sido guardado en {output_pdf_path}")
                log_file.write(
                    f"El PDF modificado ha sido guardado en {output_pdf_path} \n"
                )

                ruta_archivo_1 = os.path.join(
                    carpeta_destino,
                    f"REFERENCIA CRUZADA COMPROBANTE {numero1} ELECTRONICA.pdf"
                )

                # Esperar a que el input esté presente y visible
                input_file = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.NAME, "file"))
                )

                # Subir el archivo
                input_file.send_keys(ruta_archivo_1)

                time.sleep(5)

                # Habilitar cargar masiva
                button_masiva = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//*[@id='bttnEnablelinkMassively']")
                    )
                )
                button_masiva.click()

                time.sleep(5)

                # Selección de todos los archivos (cb_list3 checkbox)
                button_seleccion_archivos = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//*[@id='cb_list3']")  # Se cerró el paréntesis correctamente
                    )
                )
                button_seleccion_archivos.click()

                time.sleep(5)

                # Esperar explícitamente a que el botón sea clickeable
                button_vincular_archivos = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable(
                        (By.ID, "bttnlinkMassively")
                    )
                )

                # Hacer clic en el botón
                button_vincular_archivos.click()

                time.sleep(5)

                driver.refresh()

                # Esperar hasta que el menú se despliegue completamente
                time.sleep(4)

                # Encontrar el campo de expediente
                campo_filtro = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//*[@id='idExpediente']"))
                )
                campo_filtro.send_keys(numero1)

                button_filtro = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//*[@id='filter']"))
                )
                button_filtro.click()

                time.sleep(3)

                # Esperar hasta que el icono sea visible y luego hacer clic
                icono_cerrar_expediente = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located(
                        (By.CSS_SELECTOR, "td[role='gridcell'] a i[title='Cerrar Expediente']")
                    )
                )

                icono_cerrar_expediente.click()

                # Encontrar el campo de expediente
                campo_justificacion = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//*[@id='justificacion']"))
                )
                campo_justificacion.send_keys("TRANSFERENCIA DOCUMENTA 2020/TESORERÍA")

                # Esperar hasta que el botón "Aceptar" sea visible y luego hacer clic
                boton_aceptar_cerrar = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located(
                        (By.XPATH, "//button[span[text()='Aceptar']]")  # Usando XPath para seleccionar por texto
                    )
                )

                boton_aceptar_cerrar.click()

                print("Continuando con el siguiente archivo...")
                print("\n")
                log_file.write("\n")
                log_file.write("Continuando con el siguiente archivo...")
                log_file.write("\n")

                log_file.write("\n")
                log_file_error.write("\n")
                time.sleep(4)
                driver.refresh()

window = tk.Tk()
window.withdraw()  # Oculta la ventana principal
tk.messagebox.showinfo("Proceso completado", "El proceso ha finalizado.")
window.quit()
log_file.close()
driver.quit()
