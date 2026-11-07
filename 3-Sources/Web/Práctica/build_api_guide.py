from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor


OUTPUT = Path(__file__).with_name("Guia_practica_APIs_Python.docx")


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_borders(cell, color="D9D9D9", size="6"):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = qn(f"w:{edge}")
        element = borders.find(tag)
        if element is None:
            element = OxmlElement(f"w:{edge}")
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), size)
        element.set(qn("w:color"), color)


def set_cell_margins(cell, top=90, start=100, bottom=90, end=100):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def set_run_font(run, name, size=None, bold=None, color=None):
    run.font.name = name
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), name)
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), name)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if color is not None:
        run.font.color.rgb = RGBColor(*color)


def add_paragraph(doc, text="", *, bold_lead=None, style=None, keep_next=False):
    p = doc.add_paragraph(style=style)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.12
    p.paragraph_format.keep_with_next = keep_next
    if bold_lead and text.startswith(bold_lead):
        lead = p.add_run(bold_lead)
        lead.bold = True
        p.add_run(text[len(bold_lead):])
    else:
        p.add_run(text)
    return p


def add_bullets(doc, items, level=0):
    for item in items:
        p = doc.add_paragraph(style="List Bullet" if level == 0 else "List Bullet 2")
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.08
        p.add_run(item)


def add_numbered(doc, items):
    for number, item in enumerate(items, start=1):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.08
        p.add_run(f"{number}.  {item}")


def add_code(doc, code, label=None):
    if label:
        p_label = doc.add_paragraph()
        p_label.paragraph_format.space_before = Pt(4)
        p_label.paragraph_format.space_after = Pt(2)
        p_label.paragraph_format.keep_with_next = True
        run = p_label.add_run(label)
        set_run_font(run, "Aptos", 9, True, (64, 64, 64))
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.35)
    p.paragraph_format.right_indent = Cm(0.2)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.line_spacing = 1.0
    # Los bloques cortos se mantienen enteros; los largos pueden continuar
    # en la página siguiente para no dejar grandes huecos en blanco.
    p.paragraph_format.keep_together = len(code.rstrip().splitlines()) <= 18
    p_pr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), "F3F5F7")
    p_pr.append(shd)
    spacing = p_pr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        p_pr.append(spacing)
    spacing.set(qn("w:before"), "100")
    spacing.set(qn("w:after"), "100")
    run = p.add_run(code.rstrip())
    set_run_font(run, "Consolas", 8.4, False, (28, 31, 34))
    return p


def add_table(doc, headers, rows, widths=None):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    header = table.rows[0]
    set_repeat_table_header(header)
    for i, text in enumerate(headers):
        cell = header.cells[i]
        cell.text = str(text)
        set_cell_shading(cell, "1F4E78")
        set_cell_borders(cell)
        set_cell_margins(cell)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        if widths:
            cell.width = Cm(widths[i])
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(0)
            for run in p.runs:
                set_run_font(run, "Aptos", 9, True, (255, 255, 255))
    for r_idx, row in enumerate(rows):
        cells = table.add_row().cells
        for i, value in enumerate(row):
            cell = cells[i]
            cell.text = str(value)
            set_cell_borders(cell)
            set_cell_margins(cell)
            if r_idx % 2:
                set_cell_shading(cell, "F2F6FA")
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if widths:
                cell.width = Cm(widths[i])
            for p in cell.paragraphs:
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.0
                if i == 0 and len(headers) > 2:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in p.runs:
                    set_run_font(run, "Aptos", 8.7)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    return table


def add_heading(doc, text, level=1):
    p = doc.add_heading(text, level=level)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.space_before = Pt(12 if level == 1 else 8)
    p.paragraph_format.space_after = Pt(5)
    return p


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run("Página ")
    set_run_font(run, "Aptos", 8.5, False, (90, 90, 90))
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = " PAGE "
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char1)
    run._r.append(instr_text)
    run._r.append(fld_char2)


doc = Document()
section = doc.sections[0]
section.page_width = Cm(21)
section.page_height = Cm(29.7)
section.top_margin = Cm(1.8)
section.bottom_margin = Cm(1.7)
section.left_margin = Cm(1.9)
section.right_margin = Cm(1.9)

styles = doc.styles
normal = styles["Normal"]
normal.font.name = "Aptos"
normal._element.rPr.rFonts.set(qn("w:ascii"), "Aptos")
normal._element.rPr.rFonts.set(qn("w:hAnsi"), "Aptos")
normal.font.size = Pt(10.5)
normal.font.color.rgb = RGBColor(25, 25, 25)

title_style = styles["Title"]
title_style.font.name = "Aptos Display"
title_style._element.rPr.rFonts.set(qn("w:ascii"), "Aptos Display")
title_style._element.rPr.rFonts.set(qn("w:hAnsi"), "Aptos Display")
title_style.font.size = Pt(28)
title_style.font.bold = True
title_style.font.color.rgb = RGBColor(0, 0, 0)

for style_name, size in (("Heading 1", 17), ("Heading 2", 13), ("Heading 3", 11)):
    style = styles[style_name]
    style.font.name = "Aptos Display"
    style._element.rPr.rFonts.set(qn("w:ascii"), "Aptos Display")
    style._element.rPr.rFonts.set(qn("w:hAnsi"), "Aptos Display")
    style.font.size = Pt(size)
    style.font.bold = True
    style.font.color.rgb = RGBColor(0, 0, 0)

footer = section.footer
footer_p = footer.paragraphs[0]
footer_p.text = "Guía práctica de APIs con Python"
footer_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
for run in footer_p.runs:
    set_run_font(run, "Aptos", 8.5, False, (90, 90, 90))
add_page_number(footer.add_paragraph())

# Portada
p = doc.add_paragraph()
p.style = styles["Title"]
p.paragraph_format.space_before = Pt(42)
p.paragraph_format.space_after = Pt(10)
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
run = p.add_run("Guía práctica de APIs con Python")

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(24)
run = p.add_run("De una respuesta JSON a un análisis con Pandas")
set_run_font(run, "Aptos Display", 16, False, (55, 55, 55))

add_paragraph(
    doc,
    "Este manual explica cómo pensar y trabajar con una API desde Python. El objetivo no es memorizar una receta, sino aprender a inspeccionar cada respuesta, reconocer dónde están los registros y decidir cómo convertirlos en una estructura útil para el análisis.",
)
add_paragraph(
    doc,
    "La idea central es sencilla: una API entrega una respuesta; JSON transporta la información; Python la convierte en diccionarios y listas; y Pandas resulta útil cuando queremos filtrar, limpiar, calcular métricas o guardar una tabla.",
)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(20)
p.paragraph_format.space_after = Pt(5)
run = p.add_run("Ruta mental")
set_run_font(run, "Aptos", 10, True, (31, 78, 120))

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(8)
p.paragraph_format.space_after = Pt(16)
run = p.add_run("URL  →  GET  →  respuesta HTTP  →  JSON  →  Python  →  DataFrame  →  análisis  →  CSV")
set_run_font(run, "Aptos", 11, True, (31, 78, 120))

add_paragraph(doc, "Nivel: iniciación a análisis de datos con Python")
add_paragraph(doc, "Ejemplos: NASA Open APIs, NASA JPL Sentry y FRED S&P 500")
doc.add_page_break()

# Índice
add_heading(doc, "Contenido", 1)
contents = [
    "1 Qué es una API y qué ocurre en una llamada",
    "2 El patrón mínimo con requests",
    "3 Cómo inspeccionar una respuesta JSON",
    "4 Cómo decidir qué son filas y columnas",
    "5 Seis estructuras JSON frecuentes",
    "6 Limpieza y análisis con Pandas",
    "7 Ejemplos completos con NASA",
    "8 Ejemplo financiero con el S&P 500",
    "9 API y scraping web no son lo mismo",
    "10 Paginación límites reintentos y seguridad",
    "11 Diagnóstico de errores",
    "12 Plantilla reutilizable y lista de comprobación",
]
add_numbered(doc, [item.split(" ", 1)[1] for item in contents])
add_paragraph(
    doc,
    "Puedes leer la guía en orden o utilizarla como referencia. Las secciones 3, 4 y 5 son las más importantes para aprender a moverte por respuestas que todavía no conoces.",
)

# 1
add_heading(doc, "1 Qué es una API y qué ocurre en una llamada", 1)
add_paragraph(
    doc,
    "Una API es una interfaz diseñada para que un programa solicite datos o ejecute operaciones en otro sistema. En una API web, Python envía una petición HTTP a una URL y el servidor devuelve una respuesta. Esa respuesta puede contener JSON, XML, CSV, una imagen, un archivo o un mensaje de error.",
)
add_heading(doc, "Los componentes de una petición", 2)
add_table(
    doc,
    ["Componente", "Función", "Ejemplo"],
    [
        ["URL", "Identifica el servicio o recurso", "https://api.nasa.gov/neo/rest/v1/feed"],
        ["Método", "Indica la operación", "GET para consultar"],
        ["Parámetros", "Acotan o configuran la consulta", "start_date, end_date"],
        ["Autenticación", "Identifica al cliente cuando es necesaria", "api_key"],
        ["Cabeceras", "Añaden metadatos", "Authorization, User-Agent"],
        ["Cuerpo", "Envía datos en POST o PUT", "json={...}"],
    ],
    widths=[3.1, 6.2, 7.5],
)
add_paragraph(
    doc,
    "GET se utiliza normalmente para consultar. POST suele crear o enviar información; PUT y PATCH actualizan; DELETE elimina. En análisis de datos trabajaremos sobre todo con GET.",
)

add_heading(doc, "API frente a página web", 2)
add_paragraph(
    doc,
    "Una página web está pensada para una persona y devuelve HTML para que un navegador lo represente. Una API está pensada para programas y suele devolver datos estructurados. Si existe una API oficial, normalmente es preferible a extraer información del HTML porque la estructura es más estable, explícita y fácil de validar.",
)

# 2
add_heading(doc, "2 El patrón mínimo con requests", 1)
add_paragraph(doc, "La mayoría de las consultas de este curso parten del mismo patrón:")
add_code(
    doc,
    '''import requests

URL = "https://servicio.example/api/recurso"

params = {
    "parametro": "valor",
    "api_key": "TU_CLAVE"
}

response = requests.get(URL, params=params, timeout=30)
response.raise_for_status()
data = response.json()''',
    "Patrón mínimo",
)
add_table(
    doc,
    ["Línea", "Qué hace", "Por qué importa"],
    [
        ["requests.get", "Envía una petición GET", "La consulta aún no es el JSON"],
        ["raise_for_status", "Detecta códigos HTTP de error", "Evita analizar una respuesta fallida"],
        ["response.json", "Convierte JSON a objetos Python", "Obtienes dict y list utilizables"],
    ],
    widths=[4.0, 5.5, 7.3],
)

add_heading(doc, "Parámetros y URL final", 2)
add_paragraph(
    doc,
    "El diccionario params no se envía como una tabla. Requests lo codifica en la URL. Esto evita concatenar texto manualmente y gestiona caracteres especiales.",
)
add_code(
    doc,
    '''params = {"date": "2025-10-28", "api_key": "DEMO_KEY"}
response = requests.get(URL, params=params)

print(response.url)
# https://.../recurso?date=2025-10-28&api_key=DEMO_KEY''',
)
add_paragraph(
    doc,
    "El timeout limita cuánto tiempo esperará el programa. Sin él, una conexión problemática puede mantener el proceso bloqueado durante demasiado tiempo.",
)

add_heading(doc, "Autenticación", 2)
add_paragraph(
    doc,
    "No todas las API utilizan el mismo sistema. Algunas no exigen autenticación; otras usan una API key como parámetro; y otras esperan un token en una cabecera. La documentación del endpoint indica cuál corresponde.",
)
add_code(
    doc,
    '''# API key como parámetro
params = {"api_key": API_KEY}

# Token en cabecera
headers = {"Authorization": f"Bearer {TOKEN}"}
response = requests.get(URL, headers=headers, timeout=30)''',
)
add_paragraph(
    doc,
    "No publiques claves reales en notebooks o repositorios. Para proyectos personales, una opción sencilla es leerlas desde una variable de entorno.",
)
add_code(
    doc,
    '''import os

API_KEY = os.getenv("NOMBRE_API_KEY")
if not API_KEY:
    raise ValueError("Falta configurar NOMBRE_API_KEY")''',
)

# 3
add_heading(doc, "3 Cómo inspeccionar una respuesta JSON", 1)
add_paragraph(
    doc,
    "No conviertas a DataFrame a ciegas. Primero debes comprender la forma de la respuesta. No existe una clave principal universal y la primera clave de .keys() no tiene ningún significado especial.",
)
add_code(
    doc,
    '''data = response.json()

print(type(data))

if isinstance(data, dict):
    print(data.keys())
    for clave, valor in data.items():
        print(clave, type(valor))

elif isinstance(data, list):
    print("Número de elementos:", len(data))
    print("Primer elemento:", data[0])''',
    "Inspección básica",
)
add_paragraph(
    doc,
    "Para respuestas grandes, pprint ayuda a leer un ejemplo sin convertirlo todavía en tabla.",
)
add_code(
    doc,
    '''from pprint import pprint

pprint(data)
pprint(data["data"][0])  # Un registro de ejemplo''',
)
add_heading(doc, "Correspondencia entre JSON y Python", 2)
add_table(
    doc,
    ["JSON", "Python", "Uso habitual"],
    [
        ["object { }", "dict", "Conjunto de claves y valores"],
        ["array [ ]", "list", "Colección ordenada de elementos"],
        ["string", "str", "Texto, fechas o números sin convertir"],
        ["number", "int o float", "Cantidades numéricas"],
        ["true false", "True False", "Valores booleanos"],
        ["null", "None", "Dato ausente"],
    ],
    widths=[4.0, 4.0, 8.8],
)

add_heading(doc, "Inspección recursiva opcional", 2)
add_paragraph(
    doc,
    "Cuando el JSON tiene muchos niveles, esta función muestra la estructura sin imprimir todos los valores. Es una herramienta de exploración, no una transformación obligatoria.",
)
add_code(
    doc,
    '''def mostrar_estructura(objeto, nivel=0, max_elementos=3):
    sangria = "  " * nivel

    if isinstance(objeto, dict):
        print(f"{sangria}dict con {len(objeto)} claves")
        for clave, valor in list(objeto.items())[:max_elementos]:
            print(f"{sangria}- {clave}: {type(valor).__name__}")
            mostrar_estructura(valor, nivel + 1, max_elementos)

    elif isinstance(objeto, list):
        print(f"{sangria}list con {len(objeto)} elementos")
        if objeto:
            mostrar_estructura(objeto[0], nivel + 1, max_elementos)

mostrar_estructura(data)''',
)

# 4
add_heading(doc, "4 Cómo decidir qué son filas y columnas", 1)
add_paragraph(
    doc,
    "La decisión depende de la pregunta. Una fila suele representar una observación repetida: un asteroide, una fecha, una película, una ciudad o una operación. Las columnas describen atributos comparables de esas observaciones.",
)
add_paragraph(doc, "Ejemplo: si la pregunta es qué asteroides pasaron cerca de la Tierra, una fila debe representar un asteroide o una aproximación. Si la pregunta es cuántos asteroides hubo cada día, una fila puede representar una fecha.")
add_table(
    doc,
    ["Pregunta", "Una fila representa", "Columnas posibles"],
    [
        ["Qué objetos se acercaron", "Un asteroide", "id, nombre, diámetro"],
        ["Cómo fue cada aproximación", "Un acercamiento", "fecha, velocidad, distancia"],
        ["Cuántos hubo cada día", "Una fecha", "fecha, cantidad, peligrosos"],
        ["Cuál tiene más riesgo", "Un objeto Sentry", "nombre, ip, diámetro"],
    ],
    widths=[5.1, 4.8, 7.0],
)
add_paragraph(
    doc,
    "No elijas una clave porque aparezca primero. Identifica dónde están los registros repetidos y comprueba que cada elemento tenga una estructura parecida. La documentación y la inspección del primer elemento suelen bastar.",
)

# 5
add_heading(doc, "5 Seis estructuras JSON frecuentes", 1)
add_heading(doc, "Caso 1 lista de diccionarios", 2)
add_code(doc, '''data = [
    {"nombre": "A", "diametro": 20},
    {"nombre": "B", "diametro": 35}
]

df = pd.DataFrame(data)''')
add_paragraph(doc, "Cada elemento de la lista es una fila y las claves de cada diccionario son columnas.")

add_heading(doc, "Caso 2 diccionario que contiene la lista", 2)
add_code(doc, '''data = {
    "count": 2,
    "resultados": [
        {"nombre": "A", "diametro": 20},
        {"nombre": "B", "diametro": 35}
    ]
}

df = pd.DataFrame(data["resultados"])''')
add_paragraph(doc, "El contador es metadato. La lista resultados contiene los registros que queremos tabular.")

add_heading(doc, "Caso 3 diccionario de listas", 2)
add_code(doc, '''data = {
    "nombre": ["A", "B"],
    "diametro": [20, 35]
}

df = pd.DataFrame(data)''')
add_paragraph(doc, "Las claves ya son columnas y las listas tienen el mismo número de posiciones.")

add_heading(doc, "Caso 4 un solo registro", 2)
add_code(doc, '''data = {"nombre": "A", "diametro": 20}

df = pd.DataFrame([data])''')
add_paragraph(doc, "La lista exterior indica que el diccionario es una única fila.")

add_heading(doc, "Caso 5 diccionario de diccionarios", 2)
add_code(doc, '''data = {
    "asteroide_1": {"nombre": "A", "diametro": 20},
    "asteroide_2": {"nombre": "B", "diametro": 35}
}

df = pd.DataFrame.from_dict(data, orient="index")
df = df.reset_index(names="identificador")''')
add_paragraph(doc, "orient='index' hace que cada diccionario interior sea una fila.")

add_heading(doc, "Caso 6 listas agrupadas por otra clave", 2)
add_code(doc, '''data = {
    "2026-11-01": [{"nombre": "A"}, {"nombre": "B"}],
    "2026-11-02": [{"nombre": "C"}]
}

filas = []

for fecha, registros in data.items():
    for registro in registros:
        filas.append({
            "fecha": fecha,
            "nombre": registro["nombre"]
        })

df = pd.DataFrame(filas)''')
add_paragraph(doc, "Aquí es necesario aplanar las listas. Además conservamos la fecha, porque al extraer los elementos perderíamos el grupo al que pertenecían.")

add_heading(doc, "Cuándo usar json_normalize", 2)
add_paragraph(doc, "pd.json_normalize intenta expandir diccionarios anidados. Resulta práctico para una primera exploración, aunque puede producir nombres de columna largos y no siempre aplana listas internas.")
add_code(doc, '''registros = data["resultados"]
df = pd.json_normalize(registros)

print(df.columns.tolist())''')

# 6
add_heading(doc, "6 Limpieza y análisis con Pandas", 1)
add_paragraph(doc, "Crear el DataFrame no garantiza que los tipos sean correctos. Muchas API envían números y fechas como texto. Antes de calcular métricas, revisa la estructura.")
add_code(doc, '''df.head()
df.info()
df.shape
df.columns.tolist()
df.isna().sum()''', "Diagnóstico inicial")
add_heading(doc, "Conversión de tipos", 2)
add_code(doc, '''df["ip"] = pd.to_numeric(df["ip"], errors="coerce")
df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
df["peligroso"] = df["peligroso"].astype("boolean")''')
add_paragraph(doc, "errors='coerce' convierte valores no válidos en NaN o NaT. Después puedes decidir si eliminarlos, rellenarlos o investigarlos.")
add_heading(doc, "Operaciones habituales", 2)
add_code(doc, '''df["ip"].max()
df["diametro"].mean()
df.sort_values("ip", ascending=False).head(10)
df[df["peligroso"] == True]
df.groupby("fecha").size()
df.groupby("fecha")["diametro"].mean()
df.describe(include="all")''')
add_heading(doc, "Guardado y reutilización", 2)
add_code(doc, '''df.to_csv("resultado.csv", index=False, encoding="utf-8-sig")

# Recuperar los datos sin repetir la llamada
df = pd.read_csv("resultado.csv")''')
add_paragraph(doc, "Guardar evita consumir de nuevo el límite de peticiones y conserva una instantánea reproducible. La codificación utf-8-sig facilita abrir textos con tildes en Excel para Windows.")

# 7 NASA
add_heading(doc, "7 Ejemplos completos con NASA", 1)
add_heading(doc, "APOD foto astronómica del día", 2)
add_code(doc, '''import requests

URL_APOD = "https://api.nasa.gov/planetary/apod"

params = {
    "date": "2025-10-28",
    "api_key": "DEMO_KEY"
}

response = requests.get(URL_APOD, params=params, timeout=30)
response.raise_for_status()
data = response.json()

print(data["title"])
print(data["explanation"])

imagen = requests.get(data["url"], timeout=30)
imagen.raise_for_status()

with open("nasa_cumple.jpg", "wb") as archivo:
    archivo.write(imagen.content)''')
add_paragraph(doc, "La primera respuesta es JSON y contiene metadatos. La segunda respuesta es binaria porque descarga una imagen. No todas las respuestas de una API se convierten a DataFrame.")

add_heading(doc, "NeoWs asteroides cercanos", 2)
add_code(doc, '''import pandas as pd
from datetime import date, timedelta

URL_NEO = "https://api.nasa.gov/neo/rest/v1/feed"

params = {
    "start_date": str(date.today() - timedelta(days=6)),
    "end_date": str(date.today()),
    "api_key": "DEMO_KEY"
}

response = requests.get(URL_NEO, params=params, timeout=30)
response.raise_for_status()
data = response.json()

filas = []

for fecha, lista_asteroides in data["near_earth_objects"].items():
    for asteroide in lista_asteroides:
        aproximacion = asteroide["close_approach_data"][0]
        diametro = asteroide["estimated_diameter"]["kilometers"]

        filas.append({
            "fecha": fecha,
            "id": asteroide["id"],
            "nombre": asteroide["name"],
            "diametro_min_km": diametro["estimated_diameter_min"],
            "diametro_max_km": diametro["estimated_diameter_max"],
            "peligroso": asteroide["is_potentially_hazardous_asteroid"],
            "velocidad_km_s": float(
                aproximacion["relative_velocity"]["kilometers_per_second"]
            ),
            "distancia_km": float(
                aproximacion["miss_distance"]["kilometers"]
            )
        })

asteroides = pd.DataFrame(filas)
asteroides.to_csv("asteroids.csv", index=False)''')
add_paragraph(doc, "La parte decisiva es reconocer que near_earth_objects es un diccionario de fechas y que cada fecha contiene una lista. El bucle aplana esa jerarquía y conserva la fecha como columna.")

add_heading(doc, "Sentry objetos con riesgo de impacto", 2)
add_code(doc, '''URL_SENTRY = "https://ssd-api.jpl.nasa.gov/sentry.api"

response = requests.get(URL_SENTRY, timeout=30)
response.raise_for_status()
data = response.json()

print(data.keys())
# dict_keys(['signature', 'count', 'data'])

sentry = pd.DataFrame(data["data"])
sentry["ip"] = pd.to_numeric(sentry["ip"], errors="coerce")

mayor_probabilidad = sentry.loc[[sentry["ip"].idxmax()]]
sentry.to_csv("sentry.csv", index=False)

mayor_probabilidad''')
add_paragraph(doc, "En Sentry, la lista de registros está directamente en data. ip representa la probabilidad acumulada de impacto y llega como texto, por lo que se convierte antes de aplicar idxmax.")

# 8 FRED
add_heading(doc, "8 Ejemplo financiero con el S&P 500", 1)
add_paragraph(
    doc,
    "Este ejemplo utiliza la API oficial de FRED, mantenida por el Banco de la Reserva Federal de St. Louis. La serie SP500 contiene valores diarios de cierre del índice. FRED exige una API key y puede devolver JSON cuando file_type='json'.",
)
add_heading(doc, "Objetivo del análisis", 2)
add_bullets(doc, [
    "Descargar observaciones del S&P 500 desde una fecha determinada.",
    "Convertir la lista observations en un DataFrame.",
    "Limpiar los días sin valor, representados con un punto.",
    "Calcular rentabilidad diaria, media móvil y métricas descriptivas.",
    "Guardar el resultado en CSV.",
])
add_code(doc, '''import os
import pandas as pd
import requests

URL_FRED = "https://api.stlouisfed.org/fred/series/observations"
FRED_API_KEY = os.getenv("FRED_API_KEY")

if not FRED_API_KEY:
    raise ValueError("Configura la variable de entorno FRED_API_KEY")

params = {
    "series_id": "SP500",
    "api_key": FRED_API_KEY,
    "file_type": "json",
    "observation_start": "2025-01-01",
    "sort_order": "asc"
}

response = requests.get(URL_FRED, params=params, timeout=30)
response.raise_for_status()
data = response.json()

print(data.keys())
print(type(data["observations"]))
print(data["observations"][0])''', "1 Llamada e inspección")
add_paragraph(doc, "La respuesta contiene metadatos y una lista observations. Esa lista es la unidad tabular: cada elemento representa una fecha y un valor del índice.")
add_code(doc, '''sp500 = pd.DataFrame(data["observations"])

sp500 = sp500[["date", "value"]].copy()
sp500["date"] = pd.to_datetime(sp500["date"])
sp500["value"] = pd.to_numeric(sp500["value"], errors="coerce")

# El punto usado por FRED para un día sin observación se convierte en NaN.
sp500 = sp500.dropna(subset=["value"])
sp500 = sp500.sort_values("date").reset_index(drop=True)''', "2 Conversión y limpieza")
add_code(doc, '''sp500["rentabilidad_diaria_pct"] = sp500["value"].pct_change() * 100
sp500["media_movil_20"] = sp500["value"].rolling(20).mean()

print("Último cierre:", sp500.iloc[-1]["value"])
print("Máximo del periodo:", sp500["value"].max())
print("Mínimo del periodo:", sp500["value"].min())
print("Rentabilidad media diaria:", sp500["rentabilidad_diaria_pct"].mean())

fila_maximo = sp500.loc[[sp500["value"].idxmax()]]
print(fila_maximo)

sp500.to_csv("sp500_fred.csv", index=False, encoding="utf-8-sig")''', "3 Métricas y guardado")
add_paragraph(
    doc,
    "pct_change compara cada valor con el anterior. rolling(20).mean calcula una media móvil de veinte observaciones. Estas operaciones solo funcionan de forma fiable después de convertir value a número y ordenar por fecha.",
)
add_paragraph(
    doc,
    "La serie SP500 de FRED es un índice de precios, no un índice de rentabilidad total: no incorpora dividendos. FRED indica además que la disponibilidad histórica de esta serie está sujeta al acuerdo con S&P Dow Jones Indices. Revisa las notas y condiciones antes de redistribuir los datos.",
)

# 9 scraping
add_heading(doc, "9 API y scraping web no son lo mismo", 1)
add_table(
    doc,
    ["Aspecto", "API", "Scraping web"],
    [
        ["Fuente", "Endpoint diseñado para programas", "HTML diseñado para personas"],
        ["Estructura", "JSON XML o CSV", "Etiquetas HTML"],
        ["Estabilidad", "Contrato documentado", "Puede cambiar con el diseño"],
        ["Acceso", "Clave límites y parámetros", "Robots términos y carga del sitio"],
        ["Herramientas", "requests", "BeautifulSoup o read_html"],
    ],
    widths=[3.4, 6.7, 6.7],
)
add_paragraph(
    doc,
    "El ejemplo de FRED es extracción mediante API, no scraping. El scraping se utiliza cuando la información solo aparece en una página y no existe un endpoint adecuado. Antes de hacerlo, revisa las condiciones de uso, robots.txt, derechos sobre los datos y frecuencia de acceso.",
)
add_heading(doc, "Ejemplo educativo de una tabla HTML", 2)
add_paragraph(doc, "La página de Wikipedia con las empresas del S&P 500 contiene una tabla HTML. Pandas puede leer tablas visibles con read_html. Este método depende de que la página mantenga su estructura.")
add_code(doc, '''import pandas as pd

URL_COMPONENTES = (
    "https://en.wikipedia.org/wiki/List_of_S%26P_500_companies"
)

tablas = pd.read_html(URL_COMPONENTES)
componentes = tablas[0]

print(componentes.columns)
print(componentes.head())

componentes.to_csv(
    "componentes_sp500.csv",
    index=False,
    encoding="utf-8-sig"
)''')
add_paragraph(doc, "Aquí también hay una decisión: tablas[0] significa que seleccionamos la primera tabla encontrada. En un proyecto real conviene comprobar sus columnas y no asumir eternamente que seguirá en esa posición.")
add_code(doc, '''for numero, tabla in enumerate(tablas):
    print(numero, tabla.shape, tabla.columns.tolist())''', "Inspección de tablas encontradas")

# 10 robustez
add_heading(doc, "10 Paginación límites reintentos y seguridad", 1)
add_heading(doc, "Paginación", 2)
add_paragraph(doc, "Muchas API no devuelven todos los registros en una sola respuesta. Utilizan page, offset, limit o un enlace next. Debes repetir peticiones hasta terminar y acumular los registros.")
add_code(doc, '''todos = []
pagina = 1

while True:
    params = {"page": pagina, "limit": 100}
    response = requests.get(URL, params=params, timeout=30)
    response.raise_for_status()
    data = response.json()

    registros = data["results"]
    todos.extend(registros)

    if not data.get("next"):
        break

    pagina += 1

df = pd.DataFrame(todos)''')
add_heading(doc, "Límites de peticiones", 2)
add_paragraph(doc, "Las API pueden limitar peticiones por minuto, hora o día. Un 429 significa normalmente demasiadas solicitudes. Guarda respuestas, evita repetir consultas innecesarias y respeta Retry-After cuando exista.")
add_heading(doc, "Sesiones y reintentos", 2)
add_code(doc, '''from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

session = requests.Session()

retry = Retry(
    total=3,
    backoff_factor=1,
    status_forcelist=[429, 500, 502, 503, 504],
    allowed_methods=["GET"]
)

session.mount("https://", HTTPAdapter(max_retries=retry))

response = session.get(URL, params=params, timeout=30)
response.raise_for_status()''')
add_paragraph(doc, "Los reintentos son útiles para fallos temporales, no para ocultar una URL o una clave incorrectas. El backoff espera progresivamente entre intentos.")

add_heading(doc, "Validación antes de analizar", 2)
add_code(doc, '''data = response.json()

if "data" not in data:
    raise ValueError(f"Estructura inesperada: {data.keys()}")

if not isinstance(data["data"], list):
    raise TypeError("Se esperaba una lista en data['data']")

if not data["data"]:
    print("La consulta no devolvió registros")''')

add_heading(doc, "Seguridad básica", 2)
add_bullets(doc, [
    "No incluyas API keys reales en Git, capturas o documentos públicos.",
    "Usa HTTPS para evitar transmitir datos en texto plano.",
    "Concede a cada token los permisos mínimos necesarios.",
    "No imprimas cabeceras de autorización en logs compartidos.",
    "Revoca y sustituye una clave si se ha publicado accidentalmente.",
    "Lee las condiciones de uso y las restricciones de redistribución.",
])

# 11 errores
add_heading(doc, "11 Diagnóstico de errores", 1)
add_table(
    doc,
    ["Código o error", "Significado habitual", "Qué revisar"],
    [
        ["400", "Petición incorrecta", "Nombres, formatos y combinaciones de parámetros"],
        ["401", "No autorizado", "Clave ausente, incorrecta o endpoint equivocado"],
        ["403", "Acceso prohibido", "Permisos, cuota o restricciones"],
        ["404", "Recurso no encontrado", "Ruta y versión del endpoint"],
        ["429", "Demasiadas peticiones", "Límites, Retry-After y caché"],
        ["500 503", "Error temporal del servidor", "Reintentar con espera"],
        ["JSONDecodeError", "La respuesta no es JSON válido", "Content-Type y response.text"],
        ["KeyError", "La clave no existe", "data.keys y estructura real"],
        ["TypeError", "El tipo no es el esperado", "type de cada nivel"],
    ],
    widths=[3.2, 5.4, 8.2],
)
add_heading(doc, "Bloque de diagnóstico", 2)
add_code(doc, '''response = requests.get(URL, params=params, timeout=30)

print("Estado:", response.status_code)
print("URL final:", response.url)
print("Content-Type:", response.headers.get("Content-Type"))
print("Primeros caracteres:", response.text[:300])

response.raise_for_status()''')
add_paragraph(doc, "No compartas response.url si contiene una API key real. En ese caso, oculta el valor antes de copiar el mensaje de error.")

add_heading(doc, "Secuencia recomendada cuando algo falla", 2)
add_numbered(doc, [
    "Lee el código HTTP y el mensaje devuelto.",
    "Imprime la URL final sin revelar credenciales.",
    "Compara endpoint y parámetros con la documentación actual.",
    "Comprueba el formato de fechas y valores booleanos.",
    "Inspecciona type, keys y un registro real.",
    "Solo después modifica la transformación a DataFrame.",
])

# 12 plantilla
add_heading(doc, "12 Plantilla reutilizable y lista de comprobación", 1)
add_heading(doc, "Plantilla para una nueva API", 2)
add_code(doc, '''import os
import pandas as pd
import requests

# 1. Configuración
URL = "https://servicio.example/api/recurso"
API_KEY = os.getenv("SERVICIO_API_KEY")

params = {
    "parametro": "valor",
    "api_key": API_KEY
}

# 2. Petición
response = requests.get(URL, params=params, timeout=30)
response.raise_for_status()
data = response.json()

# 3. Inspección
print(type(data))
if isinstance(data, dict):
    print(data.keys())

# 4. Selección de registros
registros = data["CAMBIAR_POR_CLAVE_REAL"]
print(type(registros), len(registros))
print(registros[0])

# 5. DataFrame y tipos
df = pd.DataFrame(registros)
# df["numero"] = pd.to_numeric(df["numero"], errors="coerce")
# df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")

# 6. Validación y análisis
print(df.shape)
print(df.info())
print(df.head())

# 7. Persistencia
df.to_csv("resultado.csv", index=False, encoding="utf-8-sig")''')

add_heading(doc, "Lista de comprobación", 2)
checklist = [
    "Sé qué pregunta quiero responder.",
    "He localizado el endpoint correcto y su documentación.",
    "Sé si requiere clave, token o ninguna autenticación.",
    "He definido timeout y comprobado el código HTTP.",
    "He inspeccionado type, keys y al menos un registro.",
    "He decidido qué representa una fila.",
    "He conservado las claves de agrupación importantes, como fecha o categoría.",
    "He convertido números y fechas antes de calcular métricas.",
    "He revisado valores ausentes y duplicados.",
    "He guardado una copia local si la consulta es costosa o limitada.",
    "No he expuesto credenciales.",
    "He respetado las condiciones de uso y límites del servicio.",
]
for item in checklist:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(3)
    p.add_run("☐ ").bold = True
    p.add_run(item)

add_heading(doc, "Qué debes recordar", 2)
add_paragraph(
    doc,
    "No necesitas memorizar todas las estructuras. Necesitas saber inspeccionar una respuesta, reconocer el patrón de registros y transformar solo lo necesario para tu pregunta. A veces el DataFrame se crea en una línea; otras veces hay que aplanar varios niveles. Esa decisión forma parte del análisis, no es un error.",
)
add_paragraph(
    doc,
    "JSON no es un obstáculo: después de response.json() ya tienes diccionarios y listas de Python. Puedes trabajar directamente con ellos. Pandas entra cuando la información tiene sentido tabular y quieres comparar observaciones, limpiar tipos, calcular métricas o exportar resultados.",
)

# Fuentes
add_heading(doc, "Fuentes y documentación", 1)
sources = [
    "NASA Open APIs: https://api.nasa.gov/",
    "NASA JPL Sentry API: https://ssd-api.jpl.nasa.gov/doc/sentry.html",
    "FRED API series observations: https://fred.stlouisfed.org/docs/api/fred/series_observations.html",
    "FRED S&P 500 series SP500: https://fred.stlouisfed.org/series/SP500",
    "Requests documentation: https://requests.readthedocs.io/",
    "Pandas DataFrame: https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.html",
    "Pandas json_normalize: https://pandas.pydata.org/docs/reference/api/pandas.json_normalize.html",
]
add_bullets(doc, sources)
add_paragraph(doc, "La documentación del endpoint es la referencia definitiva para parámetros, límites, estructura de la respuesta y condiciones de uso. Las API pueden cambiar; revisa la versión vigente cuando un ejemplo deje de funcionar.")

# Propiedades y guardado
doc.core_properties.title = "Guía práctica de APIs con Python"
doc.core_properties.subject = "Conexiones API JSON Pandas y ejemplos NASA y FRED"
doc.core_properties.author = "Aitor"
doc.core_properties.keywords = "API, Python, requests, JSON, Pandas, NASA, FRED, S&P 500"

doc.save(OUTPUT)
print(OUTPUT.resolve())
