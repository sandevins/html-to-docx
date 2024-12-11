import sys
from bs4 import BeautifulSoup
from docx import Document

def convert_html_to_docx(input_file, output_file):
    try:
        # Cargar el archivo HTML
        with open(input_file, "r", encoding="utf-8") as file:
            soup = BeautifulSoup(file, "html.parser")

        # Crear un documento DOCX
        doc = Document()

        # Agregar el texto del HTML al documento
        for element in soup.find_all(["p", "h1", "h2", "h3", "h4", "h5", "h6"]):
            if element.name.startswith("h"):
                doc.add_heading(element.get_text(), level=int(element.name[1]))
            elif element.name == "p":
                doc.add_paragraph(element.get_text())

        # Guardar el archivo como DOCX
        doc.save(output_file)
        print(f"Archivo convertido con éxito: {output_file}")
    except Exception as e:
        print(f"Error al convertir el archivo: {e}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Uso: python script.py <archivo_entrada.html> <archivo_salida.docx>")
    else:
        input_file = sys.argv[1]
        output_file = sys.argv[2]
        convert_html_to_docx(input_file, output_file)