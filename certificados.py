import os
from docxtpl import DocxTemplate
import pandas as pd
from docx2pdf import convert
import datetime

def generar_certificados(csv_file, template_file, output_folder):
    # Cargar el archivo CSV en un DataFrame
    df = pd.read_csv(csv_file)

    # Recorrer cada fila del DataFrame
    for _, row in df.iterrows():
        current_date = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Obtener los valores de la fila
        name = row['name']
        type = row['type']
        document = row['document']
        course = row['course']
        cant = row['cant']
        day = row['day']
        mont = row['mont']
        year = row['year']
        center = row['center']
        city = row['city']

        # Cargar la plantilla
        doc = DocxTemplate(template_file)

        # Reemplazar los marcadores con los valores de la fila
        context = {'name': name, 'type': type, 'document': document, 'course': course, 'cant': cant, 'day': day, 'mont': mont, 'year': year, 'center': center, 'city': city}
        doc.render(context)

        # Guardar el nuevo documento en la carpeta de certificados
        output_file = os.path.join(output_folder, f'{name}_certificado_{current_date}.docx')
        doc.save(output_file)

    print("¡Certificados generados con éxito!")

# Archivos de entrada
csv_file = 'certificados.csv'
template_file = 'diploma.docx'

# Carpeta de salida
output_folder = 'certificados'

# Crear la carpeta de certificados si no existe
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Generar los certificados
generar_certificados(csv_file, template_file, output_folder)

# Ruta de la carpeta de certificados
certificados_folder = 'certificados'

# Obtener la lista de archivos .docx en la carpeta de certificados
docx_files = [file for file in os.listdir(certificados_folder) if file.endswith('.docx')]

# Convertir los archivos .docx a .pdf
for docx_file in docx_files:
    docx_path = os.path.join(certificados_folder, docx_file)
    pdf_path = os.path.join(certificados_folder, docx_file.replace('.docx', '.pdf'))
    convert(docx_path, pdf_path)

print("¡Certificados convertidos a PDF exitosamente!")

# Eliminar los archivos .docx
for docx_file in docx_files:
    docx_path = os.path.join(certificados_folder, docx_file)
    os.remove(docx_path)

print("¡Archivos .docx eliminados!")
