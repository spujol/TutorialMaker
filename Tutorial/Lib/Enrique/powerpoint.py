
import json
import sys
import subprocess
from datetime import datetime

# Install and import external libraries to 3D slicer
import pip
pip.main(['install', 'python-pptx'])
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor


## Functions not directly related to the design and formatting of the power point presentation.

# Function to open and load JSON format metadata to a dictionary in python
def abrir_archivo_json(ruta_json):
    with open(ruta_json, "r") as archivo:
        return json.load(archivo)
def guardar_presentacion(presentacion, nombre_archivo):
    presentacion.save(nombre_archivo)

# Function to automatically open pptx file
def abrir_presentacion(archivo_pptx):
    if sys.platform == 'win32':
        subprocess.run(["start", archivo_pptx], shell=True)
    elif sys.platform == 'darwin':
        subprocess.run(["open", archivo_pptx])
    else:
        subprocess.run(["xdg-open", archivo_pptx])

## Functions directly related to the design and formatting of the power point presentation.

# Function for creating a new Power Point presentation
def crear_presentacion():
    return Presentation()

# Function to creat a new slide 
# prs = is a Power Point object
# layout_index = is the type of slide following this indeces:
# Title Slide - Layout Index: 0
# Title and Content Slide - Layout Index: 1
# Section Header Slide - Layout Index: 2
# Two Content Slide - Layout Index: 3
# Title Slide with Content - Layout Index: 4
# Title Only Slide - Layout Index: 5
# Blank Slide - Layout Index: 6
# Content with Caption Slide - Layout Index: 7
# Picture with Caption Slide - Layout Index: 8
# Title and Vertical Text Slide - Layout Index: 9
# Comparison Slide - Layout Index: 10
# Content with Vertical Text Slide - Layout Index: 11
# Title and Content with Vertical Text Slide - Layout Index: 12
def agregar_diapositiva(prs, layout_index):
    return prs.slides.add_slide(prs.slide_layouts[layout_index])

# Function to add a text box 
# slide = is the slide to which you want to add the text box
# text: The text that will be displayed in the text box.
# width: The width of the text box in PowerPoint units (e.g., Inches).
# height: The height of the text box in PowerPoint units.
# left: The left position of the text box relative to the slide's top-left corner.
# top: The top position of the text box relative to the slide's top-left corner.
# font_size: The font size of the text.
# font_color: The color of the text specified as an RGB tuple.
def agregar_cuadro_texto(slide, text, width, height, left, top, font_size, font_color, background_color=None):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    p = tf.add_paragraph()
    p.text = text
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(font_size)
    p.font.color.rgb = RGBColor(font_color[0], font_color[1], font_color[2])

    if background_color:
        fill = txBox.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(background_color[0], background_color[1], background_color[2])

# Function to add the current date text. 
def agregar_fecha(slide):
    txBox = slide.shapes.add_textbox(Inches(7.5), Inches(6), Inches(2), Inches(0.5))
    tf = txBox.text_frame
    p = tf.add_paragraph()
    run = p.add_run()
    run.text = "Fecha de elaboración:"
    font = run.font
    font.size = Pt(15)
    run.font.color.rgb = RGBColor(137, 137, 137)
    p.alignment = PP_ALIGN.LEFT
    date_text = datetime.now().strftime("%d/%m/%Y")
    p = tf.add_paragraph()
    run = p.add_run()
    run.text = date_text
    font = run.font
    font.size = Pt(15)
    run.font.color.rgb = RGBColor(137, 137, 137)
    p.alignment = PP_ALIGN.CENTER

# Function to create a Tutorial cover page
# prs = is a Power Point object
# title = title of the tutorial (metadata JSON file)
# authors = Author or list of authors (metadata JSON file)
def crear_portada(prs, title, authors): 

    slide = agregar_diapositiva(prs, 6)
    agregar_cuadro_texto(slide, title , Inches(9), Inches(1), Inches(0.75), Inches(1.5), 44, (255, 255, 255), (162, 214, 228))
    if isinstance(authors, list):
        texto_autores = "\n".join(authors)
    else:
        texto_autores = authors
    agregar_cuadro_texto(slide, texto_autores, Inches(9), Inches(1), Inches(0.75), Inches(2.5), 30, (137, 137, 137))
    agregar_fecha(slide)
    slide.shapes.add_picture("C:/Users/Dell/OneDrive - Universidad Autónoma del Estado de México/Documents/3DSlicer/TutorialMaker/Tutorial/Lib/Enrique/Metadatos/3D-Slicer_Logo.jpg" , Inches(0.5), Inches(6.5), Inches(1.6), Inches(0.7))

    return slide

# Function to create a Tutorial acknowledgments page
# prs = is a Power Point object
def crear_agradecimientos(prs):
    slide = agregar_diapositiva(prs, 5)
    title = slide.shapes.title
    title.text = "Agradecimientos"
    slide.shapes.add_picture("C:/Users/Dell/OneDrive - Universidad Autónoma del Estado de México/Documents/3DSlicer/TutorialMaker/Tutorial/Lib/Enrique/Metadatos/Agredemientos.jpg",Inches(2), Inches(2), Inches(6), Inches(4))

    return slide



# Function to create a slide with the content of the tutorial steps
# prs = Power Point object
# tilte = tilte of the main step of the tutorial (metadata JSON file)
# c_step = current step number of the tutorial
# len_step = total steps number of the tutorial 

def crear_paso_tutorial(prs,title, info_step, c_step, len_step):
    
    slide = agregar_diapositiva(prs, 5)
    #Add a centered title
    title_ = slide.shapes.title
    title_.text = info_step["action"]
    title_.text_frame.paragraphs[0].font.size = Pt(36)
    #Add text information 
    txBox = slide.shapes.add_textbox(Inches(0.1), Inches(1.1), Inches(9.5), Inches(1))
    tf = txBox.text_frame
    p = tf.add_paragraph()
    run = p.add_run()
    run.text = info_step["steps-to-follow"][0]
    font = run.font
    font.size = Pt(15)
    run.font.color.rgb = RGBColor(137, 137, 137)
    p.alignment = PP_ALIGN.LEFT
    tf.word_wrap = True
    #Add image's tutorial 
    slide.shapes.add_picture("C:/Users/Dell/OneDrive - Universidad Autónoma del Estado de México/Documents/3DSlicer/TutorialMaker/Tutorial/Lib/Enrique/" + info_step["image"],Inches(1), Inches(2.1), Inches(8), Inches(5))
    #Add footer with the current step number
    agregar_num_paso(slide, title ,c_step,len_step)

    return slide


# Function to add a footer with the current step number
# slide = is the slide to which you want to add the text box
# tilte = name of the tutorial (metadata JSON file)
# c_step = current step number of the tutorial
# f_step = total steps number of the tutorial 

def agregar_num_paso(slide, title ,c_step,f_step):
    txBox = slide.shapes.add_textbox(Inches(4), Inches(6.8), Inches(2), Inches(0.5))
    tf = txBox.text_frame
    p = tf.add_paragraph()
    run = p.add_run()
    run.text = title +  ": Paso " + str(c_step) + " de" + str(f_step)
    font = run.font
    font.size = Pt(15)
    run.font.color.rgb = RGBColor(137, 137, 137)
    p.alignment = PP_ALIGN.CENTER


# Load metadata JSON file
ruta_json = "C:/Users/Dell/OneDrive - Universidad Autónoma del Estado de México/Documents/3DSlicer/TutorialMaker/Tutorial/Lib/Enrique/Metadatos/sample_edited.json"
data = abrir_archivo_json(ruta_json)

# Tutorial generation
prs = crear_presentacion()
crear_portada(prs, data["title"], data["authors"])
for c_step, info_step in enumerate(data["instructions"]):
    crear_paso_tutorial(prs,data["title"],info_step,c_step+1,len(data["instructions"]))
crear_agradecimientos(prs)


# Save and open the pptx Tutorial file
nombre_archivo = data["title"].replace(" ","") + ".pptx"
guardar_presentacion(prs, nombre_archivo)
abrir_presentacion(nombre_archivo)



