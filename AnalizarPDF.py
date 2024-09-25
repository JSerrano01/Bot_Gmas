import pdfplumber
import re
import os

# Ruta de la carpeta que contiene los PDFs
folder_path = (
    "C:/xampp/htdocs/Bot_Gmas/pdfs"  # Ajusta esta ruta según tu estructura de carpetas
)

# Patrón de fecha que estás buscando
date_pattern = r"\d{1,2} de \w+ de \d{4}"

# Lista para almacenar logs de errores
logs = []

# Recorrer todos los archivos en la carpeta
for filename in os.listdir(folder_path):
    if filename.endswith(".pdf"):
        pdf_path = os.path.join(folder_path, filename)

        try:
            # Abrir y leer el PDF
            with pdfplumber.open(pdf_path) as pdf:
                first_page = pdf.pages[0]
                text = first_page.extract_text()

            # Buscar la fecha en el texto
            date_match = re.search(date_pattern, text)

            if date_match:
                print(f"Fecha encontrada en {filename}: {date_match.group()}")
            else:
                log_message = f"ERROR. No se encontró una fecha en el archivo {filename}."
                print(log_message)
                logs.append(log_message)

        except Exception as e:
            log_message = f"Error al leer el archivo {filename}: {str(e)}"
            print(log_message)
            logs.append(log_message)

# Guardar los logs en un archivo de texto
log_file_path = os.path.join(folder_path, "logs.txt")
with open(log_file_path, "w", encoding="utf-8") as log_file:
    for log in logs:
        log_file.write(log + "\n")

print("Proceso completado. Logs generados en 'logs.txt'.")
