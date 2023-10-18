import json
import asyncio
from pyppeteer import launch

async def generate_pdf(html_content, pdf_file):
    browser = await launch()
    page = await browser.newPage()
    await page.setContent(html_content)
    await page.pdf({'path': pdf_file, 'format': 'A4'})
    await browser.close()

async def main():
    # Cargar datos desde el archivo JSON
    with open('/Users/victormontanoserrano/Documents/3DSlicer/SlicerLatinAmerica/TutorialMaker/PDF_creator/user_generates_this_file_from_the_tool.json', 'r') as json_file:
        presentation_data = json.load(json_file)

    # Ruta al directorio de imágenes relativo al JSON
    image_directory = '../../'

    # Crear un archivo HTML para todas las diapositivas
    html_content = '<h1>Portada</h1>'
    html_content += f'<h2>{presentation_data["title"]}</h2>'
    html_content += f'<p>Authors: {", ".join(presentation_data["authors"])}</p>'
    html_content += f'<p>{presentation_data["description"]}</p>'

    for i, slide in enumerate(presentation_data['instructions'], start=1):
        html_content += f"<div style='page-break-before: always;'></div>"
        html_content += f"<h1>Slide {i}</h1>"
        html_content += f"<p>Action: {slide['action']}</p>"
        html_content += f"<p>Steps to follow:</p>"
        for step in slide['steps-to-follow']:
            html_content += f"<p>{step}</p>"
        html_content += f'<img src="{image_directory}{slide["image"]}" alt="Slide Image">'

    # Nombre del archivo PDF de salida
    pdf_file = 'presentation.pdf'

    # Generar el PDF para todas las diapositivas en hojas separadas
    await generate_pdf(html_content, pdf_file)

    print("Se ha generado un archivo PDF con diapositivas en hojas separadas.")

if __name__ == '__main__':
    asyncio.get_event_loop().run_until_complete(main())
