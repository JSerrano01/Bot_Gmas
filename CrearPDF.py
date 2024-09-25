from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PyPDF2 import PdfReader, PdfWriter
import os

# Ruta del archivo de entrada y salida
input_pdf_path = "11 V1.pdf"
output_pdf_path = "1111_modified_final.pdf"

# Dirección que se va a insertar
direccion = "EXP10234213"

# Crear un PDF temporal con la dirección
temp_pdf_path = "resultado.pdf"
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

# Guardar el nuevo PDF
with open(output_pdf_path, "wb") as output_pdf_file:
    pdf_writer.write(output_pdf_file)

# Limpiar el archivo temporal
os.remove(temp_pdf_path)

print(f"El PDF modificado ha sido guardado en {output_pdf_path}")
