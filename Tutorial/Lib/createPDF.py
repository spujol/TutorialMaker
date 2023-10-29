# This script is designed to automate the process of generating a 
# presentation from data stored in a JSON file. It takes the input 
# data from the JSON file, including presentation details, slide 
# content, and associated images. The script then creates a Markdown 
# file representing the presentation's content, including a title 
# page and individual slides.

# After generating the Markdown file, the script converts it into an 
# HTML file, allowing for a visual representation of the presentation. 
# Additionally, it utilizes the WeasyPrint and Puppeteer libraries to 
# convert the HTML into a PDF, providing a more widely accessible format 
# for the final presentation.

# The script follows a series of steps, including loading JSON data, 
# generating Markdown content, saving both Markdown and HTML files, 
# and ultimately producing a PDF document. It automates this process,
# making it easier to create presentations from structured JSON data.

# This script simplifies the process of converting structured data into 
# a presentation, which can be particularly useful for scenarios where 
# presentations need to be dynamically generated from external data sources.

import json
from weasyprint import HTML
import asyncio
from pyppeteer import launch
import os

# Load data from the JSON file whit edited images.
with open('user_generates_this_file_from_the_tool.json', 'r') as json_file:
    presentation_data = json.load(json_file)

# Name of the output files.
output_md_file = 'presentation.md'
output_html_file = 'presentation.html'  
output_pdf_file = 'presentation.pdf'

# Open the Markdown file in write mode
with open(output_md_file, 'w', encoding='utf-8') as md_file:
    # Link the CSS file in the Markdown header
    md_file.write('<link rel="stylesheet" type="text/css" href="styles.css">\n')
    md_file.write('<meta charset="UTF-8">\n')

    # Title page with HTML
    md_file.write('<div class="title-page">\n')
    md_file.write(f'<h1>{presentation_data["title"]}</h1>\n')
    md_file.write(f'<p class="authors">Authors: {", ".join(presentation_data["authors"])}</p>\n')
    md_file.write(f'<p>{presentation_data["description"]}</p>\n')
    md_file.write('</div>\n')

    # For slides
    for slide in presentation_data['instructions']:
        md_file.write('\n')

        # Wrap all the content of the slide in a div
        md_file.write('<div class="slide" style="page-break-before: always;">\n')

        md_file.write(f'<h2>{slide["action"]}</h2>\n\n')

        md_file.write("<ul class='steps'>\n")
        for step in slide['steps-to-follow']:
            md_file.write(f"<li>{step}</li>\n")
        md_file.write("</ul>\n")

        # Insert the image in Markdown with CSS styling
        md_file.write(f'<img class="screenshot" src="{slide["image"]}">\n\n')

        # Close the slide div
        md_file.write('</div>\n')

# Save the HTML file
with open(output_html_file, 'w', encoding='utf-8') as html_file:
    html_file.write(open(output_md_file, 'r', encoding='utf-8').read())

print(f"Markdown file has been generated at '{output_md_file}'")
print(f"HTML file has been generated at '{output_html_file}'")

# Convert the HTML file to PDF using Puppeteer
async def convert_html_to_pdf():
    browser = await launch()
    page = await browser.newPage()

    await page.goto(f'file://{os.path.abspath(output_html_file)}')  # Load the generated HTML file
    await page.pdf({
    'path': output_pdf_file,
    'format': 'A4',
    'landscape': True
})
    await browser.close()

if __name__ == '__main__':
    loop = asyncio.get_event_loop()
    loop.run_until_complete(convert_html_to_pdf())
    print(f"PDF file has been generated at '{output_pdf_file}'")
