from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PyPDF2 import PdfReader, PdfWriter
import os

# Ruta del archivo de entrada y salida
input_pdf_path = "Ref_Cruzada_Electronica.pdf"
output_pdf_path = "Ref_Cruzada_Electronica_modified_final.pdf"

# Direcciones y sus coordenadas específicas
direcciones = [
    ("10/04/2020", 351, 625),  # Dirección 1 que seia el date_match
    ("82045", 367, 597.4),  # Dirección 2 numero 1
    ("9754", 440, 597.4),  # Dirección 3 numero 2
]

# Crear un PDF temporal con las direcciones
temp_pdf_path = "resultado.pdf"
c = canvas.Canvas(temp_pdf_path, pagesize=letter)

# Ajustar la fuente y color
c.setFont("Helvetica-Bold", 5.5)  # Cambiar a Helvetica-Bold para texto en negrilla
c.setFillColorRGB(1, 0, 0)  # Establecer el color rojo

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

# Guardar el nuevo PDF
with open(output_pdf_path, "wb") as output_pdf_file:
    pdf_writer.write(output_pdf_file)

# Limpiar el archivo temporal
os.remove(temp_pdf_path)

print(f"El PDF modificado ha sido guardado en {output_pdf_path}")
