import math
import os
import re
import unicodedata
from io import BytesIO

import pandas as pd
import streamlit as st

from PIL import Image, ImageChops

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth


st.set_page_config(
    page_title="Blue Line - Sistema PDF",
    layout="wide"
)

st.title("Blue Line - Sistema de Capacitación")
st.write("Sube el Excel de Tally, filtra por cliente y genera el PDF correspondiente.")


# ==========================================================
# CONFIGURACIÓN GENERAL
# ==========================================================

CLIENTES = [
    "SOLGAS",
    "CÁLIDDA",
    "TRANSPORTE 77",
    "GENERAL"
]

COLUMNA_CLIENTE = "CLIENTE PARA EL QUE TRABAJA"

TITULO_FORMULARIO = "Blue Line - CAPACITACIONES DE SEGURIDAD IX"
SUBTITULO_FORMULARIO = "Form: Blue Line - CAPACITACIONES DE SEGURIDAD IX"

RAZON_SOCIAL_FIJA = "BLUE LINE DISTRIBUIDORA Y COMERCIALIZADORA S.A.C."
RUC_FIJO = "20605319808"
DIRECCION_FIJA = "Otr. S/n Mza. B Lote. 07 Asc. Agropecuaria Lomas el Dorado.\nDistrito: Puente Piedra"
ACTIVIDAD_FIJA = "Transporte de carga por carretera"

CODIGO_FORMATO = "BL-SST-FO-007"
VERSION_FORMATO = "02"
FECHA_APROBACION = "Enero 2025"

# IMPORTANTE: siempre fijo
NUM_TRABAJADORES_FIJO = "38"

# El logo debe estar dentro de la misma carpeta del app.py
LOGO_PATH = "logo_blueline.png"

GRIS_HEADER = colors.Color(0.78, 0.78, 0.78)
GRIS_CLARO = colors.Color(0.90, 0.90, 0.90)
ROJO_MARCA = colors.Color(0.72, 0.25, 0.25)

ALIAS = {
    "tema": [
        "ASISTENCIA A LA CAPACITACIÓN VIRTUAL DE",
        "ASISTENCIA A LA CAPACITACION VIRTUAL DE",
    ],
    "expositor": [
        "EXPOSITOR",
        "NOMBRE DEL PONENTE",
    ],
    "nombre": [
        "NOMBRES Y APELLIDOS",
        "NOMBRE Y APELLIDOS",
        "NOMBRE Y APELLIDO",
        "APELLIDOS Y NOMBRES",
    ],
    "dni": [
        "DNI",
        "N° DNI",
        "NRO DNI",
    ],
    "cargo": [
        "CARGO",
        "ÁREA",
        "AREA",
    ],
    "fecha_capacitacion": [
        "FECHA DE LA CAPACITACIÓN",
        "FECHA DE LA CAPACITACION",
        "FECHA",
    ],
    "responsable": [
        "NOMBRE Y CARGO DEL RESPONSABLE DEL REGISTRO",
        "NOMBRE / CARGO DEL RESPONSABLE DEL REGISTRO:",
        "NOMBRE / CARGO DEL RESPONSABLE DEL REGISTRO",
    ],
    "fecha_registro": [
        "FECHA DEL REGISTRO",
        "UNTITLED DATE FIELD",
    ],
    "cliente": [
        "CLIENTE PARA EL QUE TRABAJA",
    ],
}


# ==========================================================
# FUNCIONES DE LIMPIEZA Y LECTURA DE EXCEL
# ==========================================================

def limpiar_texto(valor):
    if pd.isna(valor):
        return ""
    valor = str(valor)
    valor = valor.replace("\t", " ")
    valor = valor.replace("\n", " ")
    valor = " ".join(valor.split())
    valor = valor.replace(" *", "")
    valor = valor.replace("*", "")

    if valor.endswith(".0") and valor.replace(".0", "").isdigit():
        valor = valor.replace(".0", "")

    return valor.strip()


def normalizar_para_comparar(valor):
    valor = limpiar_texto(valor).upper()
    valor = unicodedata.normalize("NFD", valor)
    valor = "".join(c for c in valor if unicodedata.category(c) != "Mn")
    return valor


def limpiar_columna(columna):
    return limpiar_texto(columna)


def detectar_fila_encabezado(df_raw):
    objetivo = normalizar_para_comparar(COLUMNA_CLIENTE)

    for i, fila in df_raw.iterrows():
        valores = [normalizar_para_comparar(v) for v in fila.tolist()]
        if objetivo in valores:
            return i

    objetivos_alternativos = [
        normalizar_para_comparar("NOMBRES Y APELLIDOS"),
        normalizar_para_comparar("NOMBRE Y APELLIDOS"),
        normalizar_para_comparar("DNI"),
    ]

    for i, fila in df_raw.iterrows():
        valores = [normalizar_para_comparar(v) for v in fila.tolist()]
        coincidencias = sum(1 for obj in objetivos_alternativos if obj in valores)
        if coincidencias >= 2:
            return i

    return None


def leer_excel_inteligente(archivo):
    if archivo.name.lower().endswith(".csv"):
        df_raw = pd.read_csv(archivo, header=None, dtype=object)
    else:
        df_raw = pd.read_excel(archivo, header=None, dtype=object)

    fila_encabezado = detectar_fila_encabezado(df_raw)

    if fila_encabezado is None:
        raise ValueError(
            "No se encontró la fila de encabezados. "
            "El Excel debe tener columnas como CLIENTE PARA EL QUE TRABAJA, NOMBRES Y APELLIDOS y DNI."
        )

    columnas = df_raw.iloc[fila_encabezado].tolist()
    columnas = [limpiar_columna(c) for c in columnas]

    df = df_raw.iloc[fila_encabezado + 1:].copy()
    df.columns = columnas
    df = df.dropna(how="all")

    columnas_validas = []
    for c in df.columns:
        columnas_validas.append(str(c).strip() != "" and str(c).strip().lower() != "nan")

    df = df.loc[:, columnas_validas]

    return df


def normalizar_columnas(df):
    df = df.copy()
    df.columns = [limpiar_columna(c) for c in df.columns]

    renombres = {
        "Untitled short answer field": "DNI",
        "Untitled short answer field (2)": "CARGO",
        "Untitled date field": "FECHA DEL REGISTRO",
        "NOMBRE Y APELLIDOS": "NOMBRES Y APELLIDOS",
        "Nombre y Apellido": "NOMBRES Y APELLIDOS",
        "DOMICILIO (Dirección, distrito, departamento, provincia) (2)": "TIPO DE ACTIVIDAD ECONÓMICA",
        "NOMBRE Y CARGO DEL RESPONSABLE DEL REGISTRO *": "NOMBRE Y CARGO DEL RESPONSABLE DEL REGISTRO"
    }

    df = df.rename(columns=renombres)
    df = df.loc[:, ~df.columns.duplicated()]
    return df


def preparar_para_mostrar(df):
    df_mostrar = df.copy()
    for col in df_mostrar.columns:
        df_mostrar[col] = df_mostrar[col].apply(limpiar_texto)
    return df_mostrar


def convertir_excel_descarga(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Filtrado")
    output.seek(0)
    return output


def buscar_columna(df, alias_key):
    posibles = ALIAS.get(alias_key, [])
    mapa = {normalizar_para_comparar(c): c for c in df.columns}

    for posible in posibles:
        clave = normalizar_para_comparar(posible)
        if clave in mapa:
            return mapa[clave]
    return None


def valor_fila(fila, df, alias_key, default=""):
    col = buscar_columna(df, alias_key)
    if col is None:
        return default
    return limpiar_texto(fila.get(col, default))


def primer_valor(df, alias_key, default=""):
    col = buscar_columna(df, alias_key)
    if col is None:
        return default

    for v in df[col].tolist():
        txt = limpiar_texto(v)
        if txt:
            return txt

    return default


def filtrar_participantes_validos(df):
    col_nombre = buscar_columna(df, "nombre")
    col_dni = buscar_columna(df, "dni")

    if col_nombre is None or col_dni is None:
        raise ValueError("Faltan columnas obligatorias para el PDF: NOMBRES Y APELLIDOS y DNI.")

    df = df.copy()
    df = df[df[col_nombre].apply(limpiar_texto) != ""]
    df = df[df[col_dni].apply(limpiar_texto) != ""]
    return df


def dividir_responsable(texto_resp):
    texto_resp = limpiar_texto(texto_resp)

    if "/" in texto_resp:
        partes = texto_resp.split("/", 1)
        return limpiar_texto(partes[0]), limpiar_texto(partes[1])

    if " - " in texto_resp:
        partes = texto_resp.split(" - ", 1)
        return limpiar_texto(partes[0]), limpiar_texto(partes[1])

    return texto_resp, ""


# ==========================================================
# FUNCIONES DE FECHA
# ==========================================================

def parse_fecha(valor):
    txt = limpiar_texto(valor)
    if not txt:
        return None

    meses = {
        "ENE": 1, "ENERO": 1,
        "FEB": 2, "FEBRERO": 2,
        "MAR": 3, "MARZO": 3,
        "ABR": 4, "ABRIL": 4,
        "MAY": 5, "MAYO": 5,
        "JUN": 6, "JUNIO": 6,
        "JUL": 7, "JULIO": 7,
        "AGO": 8, "AGOSTO": 8,
        "SEP": 9, "SET": 9, "SEPTIEMBRE": 9, "SETIEMBRE": 9,
        "OCT": 10, "OCTUBRE": 10,
        "NOV": 11, "NOVIEMBRE": 11,
        "DIC": 12, "DICIEMBRE": 12,
    }

    m = re.match(r"^(\d{1,2})[-/ ]([A-Za-zÁÉÍÓÚáéíóúÑñ]+)[-/ ](\d{4})$", txt)
    if m:
        dia = int(m.group(1))
        mes_txt = normalizar_para_comparar(m.group(2))
        anio = int(m.group(3))
        mes = meses.get(mes_txt)
        if mes:
            return pd.Timestamp(year=anio, month=mes, day=dia)

    fecha = pd.to_datetime(txt, errors="coerce", dayfirst=True)
    if pd.isna(fecha):
        fecha = pd.to_datetime(txt, errors="coerce", dayfirst=False)

    if pd.isna(fecha):
        return None

    return fecha


def fecha_ddmmyyyy(valor):
    fecha = parse_fecha(valor)
    if fecha is None:
        return limpiar_texto(valor)
    return fecha.strftime("%d/%m/%Y")


def fecha_dd_mmm_yyyy(valor):
    fecha = parse_fecha(valor)
    if fecha is None:
        return limpiar_texto(valor)

    meses = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
    return f"{fecha.day:02d}-{meses[fecha.month - 1]}-{fecha.year}"


# ==========================================================
# FUNCIONES PARA DIBUJAR PDF
# ==========================================================

def wrap_text(text, max_width, font_name, font_size):
    text = limpiar_texto(text)
    if not text:
        return [""]

    raw_lines = text.split("\n")
    final_lines = []

    for raw in raw_lines:
        words = raw.split()

        if not words:
            final_lines.append("")
            continue

        line = ""
        for word in words:
            test = word if not line else line + " " + word

            if stringWidth(test, font_name, font_size) <= max_width:
                line = test
            else:
                if line:
                    final_lines.append(line)
                line = word

        if line:
            final_lines.append(line)

    return final_lines


def draw_text_in_cell(
    c,
    text,
    x,
    y_top,
    w,
    h,
    font_name="Helvetica",
    font_size=7,
    align="left",
    valign="middle",
    bold=False,
    padding=3,
    max_lines=None
):
    if bold:
        font_name = "Helvetica-Bold"

    text = limpiar_texto(text)
    lines = wrap_text(text, w - 2 * padding, font_name, font_size)
    line_height = font_size + 2

    if max_lines is None:
        max_lines = max(1, int((h - 2 * padding) // line_height))

    if len(lines) > max_lines:
        lines = lines[:max_lines]
        if lines:
            last = lines[-1]
            while stringWidth(last + "...", font_name, font_size) > w - 2 * padding and len(last) > 1:
                last = last[:-1]
            lines[-1] = last + "..."

    total_text_h = len(lines) * line_height

    if valign == "top":
        y = y_top - padding - font_size
    elif valign == "bottom":
        y = y_top - h + padding + total_text_h - font_size
    else:
        y = y_top - (h - total_text_h) / 2 - font_size

    c.setFont(font_name, font_size)
    c.setFillColor(colors.black)

    for line in lines:
        if align == "center":
            tx = x + (w - stringWidth(line, font_name, font_size)) / 2
        elif align == "right":
            tx = x + w - padding - stringWidth(line, font_name, font_size)
        else:
            tx = x + padding

        c.drawString(tx, y, line)
        y -= line_height


def draw_cell(c, x, y_top, w, h, fill=None, stroke_color=colors.black, line_width=0.7):
    if fill is not None:
        c.setFillColor(fill)
        c.rect(x, y_top - h, w, h, fill=1, stroke=0)

    c.setStrokeColor(stroke_color)
    c.setLineWidth(line_width)
    c.rect(x, y_top - h, w, h, fill=0, stroke=1)


def draw_cell_text(
    c,
    text,
    x,
    y_top,
    w,
    h,
    fill=None,
    font_size=7,
    bold=False,
    align="left",
    valign="middle",
    line_width=0.7
):
    draw_cell(c, x, y_top, w, h, fill=fill, line_width=line_width)
    draw_text_in_cell(
        c,
        text,
        x,
        y_top,
        w,
        h,
        font_size=font_size,
        bold=bold,
        align=align,
        valign=valign
    )


def draw_checkbox(c, x, y, checked=False):
    c.setStrokeColor(colors.black)
    c.setLineWidth(0.7)
    c.rect(x, y, 11, 11, fill=0, stroke=1)

    if checked:
        c.setFillColor(ROJO_MARCA)
        c.rect(x + 1.5, y + 1.5, 8, 8, fill=1, stroke=0)
        c.setFillColor(colors.black)


def obtener_logo_reader():
    """
    Recorta espacios blancos del logo para que quede bien dentro del recuadro.
    """
    if not os.path.exists(LOGO_PATH):
        return None

    try:
        img = Image.open(LOGO_PATH).convert("RGBA")

        fondo = Image.new("RGBA", img.size, (255, 255, 255, 255))
        diff = ImageChops.difference(img, fondo)

        bbox = diff.getbbox()
        if bbox:
            img = img.crop(bbox)

        buffer = BytesIO()
        img.save(buffer, format="PNG")
        buffer.seek(0)

        return ImageReader(buffer)

    except Exception:
        return None


def draw_logo_blueline(c, x, y_top, w, h):
    draw_cell(c, x, y_top, w, h)

    logo = obtener_logo_reader()

    if logo is None:
        c.setFont("Helvetica-Bold", 14)
        c.setFillColor(colors.Color(0.1, 0.12, 0.32))
        c.drawString(x + 18, y_top - 30, "BLUE LINE")
        c.setFillColor(colors.black)
        return

    iw, ih = logo.getSize()

    max_w = w - 2
    max_h = h - 2

    scale = min(max_w / iw, max_h / ih)

    draw_w = iw * scale
    draw_h = ih * scale

    img_x = x + (w - draw_w) / 2
    img_y = (y_top - h) + (h - draw_h) / 2

    c.drawImage(
        logo,
        img_x,
        img_y,
        width=draw_w,
        height=draw_h,
        preserveAspectRatio=True,
        mask="auto"
    )


def obtener_datos_comunes(df, tema_manual=None, fecha_manual=None, expositor_manual=None, responsable_manual=None):
    tema = tema_manual if tema_manual is not None else primer_valor(df, "tema", "")
    fecha = fecha_manual if fecha_manual is not None else primer_valor(df, "fecha_capacitacion", "")
    expositor = expositor_manual if expositor_manual is not None else primer_valor(df, "expositor", "")
    responsable = responsable_manual if responsable_manual is not None else primer_valor(df, "responsable", "")

    return {
        "tema": limpiar_texto(tema),
        "fecha_general": fecha_ddmmyyyy(fecha),
        "fecha_individual": fecha_dd_mmm_yyyy(fecha),
        "expositor": limpiar_texto(expositor),
        "responsable": limpiar_texto(responsable),
    }


# ==========================================================
# PÁGINA CONSOLIDADA
# ==========================================================

def draw_consolidado(c, df_chunk, df_total, datos, page_num, total_pages, hora_inicio, hora_termino):
    W, H = A4
    x0 = 10
    y = H - 10
    full_w = W - 20

    draw_cell_text(
        c,
        "REGISTRO DE INDUCCIÓN, CAPACITACIÓN, ENTRENAMIENTO Y SIMULACROS DE EMERGENCIA",
        x0,
        y,
        full_w,
        18,
        fill=GRIS_CLARO,
        font_size=7.6,
        bold=True,
        align="center"
    )
    y -= 18

    logo_w = 120
    info_w = full_w - logo_w - 60
    version_w = 60
    header_h = 52

    draw_logo_blueline(c, x0, y, logo_w, header_h)

    x_info = x0 + logo_w
    draw_cell(c, x_info, y, info_w, header_h)

    draw_cell(c, x_info, y, info_w / 2, 26)
    draw_cell(c, x_info + info_w / 2, y, info_w / 2, 26)
    draw_cell(c, x_info, y - 26, info_w / 2, 26)
    draw_cell(c, x_info + info_w / 2, y - 26, info_w / 2, 26)

    draw_text_in_cell(c, "Tipo de Documento:", x_info, y, info_w / 2, 26, font_size=7, bold=True)
    draw_text_in_cell(c, "FORMATO", x_info + 100, y, info_w / 2 - 100, 26, font_size=7, bold=True)

    draw_text_in_cell(c, "Código:", x_info + info_w / 2, y, info_w / 2, 26, font_size=7, bold=True)
    draw_text_in_cell(c, CODIGO_FORMATO, x_info + info_w / 2 + 65, y, info_w / 2 - 65, 26, font_size=7, bold=True)

    draw_text_in_cell(c, "Fecha de aprobación:", x_info, y - 26, info_w / 2, 26, font_size=7, bold=True)
    draw_text_in_cell(c, FECHA_APROBACION, x_info + 105, y - 26, info_w / 2 - 105, 26, font_size=7, bold=True)

    draw_text_in_cell(c, "Página:", x_info + info_w / 2, y - 26, info_w / 2, 26, font_size=7, bold=True)
    draw_text_in_cell(c, f"{page_num} de {total_pages}", x_info + info_w / 2 + 65, y - 26, info_w / 2 - 65, 26, font_size=7, bold=True)

    draw_cell_text(
        c,
        f"Versión {VERSION_FORMATO}",
        x0 + logo_w + info_w,
        y,
        version_w,
        header_h,
        fill=GRIS_CLARO,
        font_size=7,
        bold=True,
        align="center"
    )

    y -= header_h + 8

    draw_cell_text(
        c,
        "Datos de la Empresa",
        x0,
        y,
        full_w,
        14,
        fill=GRIS_HEADER,
        font_size=7,
        bold=True,
        align="center"
    )
    y -= 14

    col_ws = [115, 85, 120, 205, full_w - 115 - 85 - 120 - 205]
    headers = ["Razón Social:", "RUC:", "Dirección:", "Actividad Económica:", "N° Trabajadores:"]
    row_h1 = 18
    row_h2 = 42

    x = x0
    for w, htext in zip(col_ws, headers):
        draw_cell_text(c, htext, x, y, w, row_h1, fill=GRIS_CLARO, font_size=6.4, align="center")
        x += w

    y -= row_h1

    valores = [
        RAZON_SOCIAL_FIJA,
        RUC_FIJO,
        DIRECCION_FIJA,
        ACTIVIDAD_FIJA,
        NUM_TRABAJADORES_FIJO
    ]

    x = x0
    for w, val in zip(col_ws, valores):
        draw_cell_text(c, val, x, y, w, row_h2, font_size=6.4, align="center")
        x += w

    y -= row_h2 + 6

    box_h = 42
    draw_cell(c, x0, y, full_w, box_h)

    c.setFont("Helvetica", 7)
    c.setFillColor(colors.black)

    c.drawString(x0 + 2, y - 14, "Inducción")
    draw_checkbox(c, x0 + 85, y - 19, checked=False)

    c.drawString(x0 + 130, y - 14, "Capacitación")
    draw_checkbox(c, x0 + 205, y - 19, checked=True)

    c.drawString(x0 + 300, y - 14, "Entrenamiento")
    draw_checkbox(c, x0 + 380, y - 19, checked=False)

    c.drawString(x0 + 430, y - 14, "Simulacro de Emergencia")
    draw_checkbox(c, x0 + 535, y - 19, checked=False)

    c.drawString(x0 + 2, y - 33, "Información")
    draw_checkbox(c, x0 + 85, y - 38, checked=False)

    c.drawString(x0 + 130, y - 33, "Charla de 5min.")
    draw_checkbox(c, x0 + 205, y - 38, checked=False)

    c.drawString(x0 + 330, y - 33, "Otros : __________________")

    y -= box_h + 8

    draw_cell_text(
        c,
        "Datos Generales",
        x0,
        y,
        full_w,
        14,
        fill=GRIS_HEADER,
        font_size=7,
        bold=True,
        align="center"
    )
    y -= 14

    draw_cell(c, x0, y, full_w * 0.74, 22)
    draw_text_in_cell(c, f"Tema: {datos['tema']}", x0, y, full_w * 0.74, 22, font_size=6.8)

    draw_cell(c, x0 + full_w * 0.74, y, full_w * 0.26, 22)
    draw_text_in_cell(c, f"Fecha: {datos['fecha_general']}", x0 + full_w * 0.74, y, full_w * 0.26, 22, font_size=6.8)

    y -= 22

    draw_cell(c, x0, y, full_w, 22)
    draw_text_in_cell(c, f"Nombre del Ponente: {datos['expositor']}", x0, y, full_w, 22, font_size=6.8)

    y -= 22

    thirds = [full_w * 0.35, full_w * 0.39, full_w * 0.26]

    draw_cell(c, x0, y, thirds[0], 28)
    draw_text_in_cell(c, f"Hora de inicio:{hora_inicio}", x0, y, thirds[0], 28, font_size=6.8)

    draw_cell(c, x0 + thirds[0], y, thirds[1], 28)
    draw_text_in_cell(c, f"Hora de término: {hora_termino}", x0 + thirds[0], y, thirds[1], 28, font_size=6.8)

    draw_cell(c, x0 + thirds[0] + thirds[1], y, thirds[2], 28)
    draw_text_in_cell(c, "Firma:", x0 + thirds[0] + thirds[1], y, thirds[2], 28, font_size=6.8)

    y -= 28 + 5

    draw_cell_text(
        c,
        "Para ser llenado por el participante",
        x0,
        y,
        full_w,
        14,
        fill=GRIS_HEADER,
        font_size=7,
        bold=True,
        align="center"
    )
    y -= 14

    headers = ["N°", "Apellidos y Nombres", "DNI", "Área", "Firma", "Observaciones"]
    col_ws = [22, 185, 120, 110, 80, full_w - 22 - 185 - 120 - 110 - 80]
    header_h = 16
    row_h = 18

    x = x0
    for w, htext in zip(col_ws, headers):
        draw_cell_text(c, htext, x, y, w, header_h, fill=GRIS_CLARO, font_size=6.8, bold=True, align="center")
        x += w

    y -= header_h

    start_index = (page_num - 1) * 20
    rows = df_chunk.to_dict("records")

    for i in range(20):
        x = x0

        if i < len(rows):
            fila = pd.Series(rows[i])
            datos_fila = [
                str(start_index + i + 1),
                valor_fila(fila, df_total, "nombre").upper(),
                valor_fila(fila, df_total, "dni"),
                valor_fila(fila, df_total, "cargo").upper(),
                "",
                "",
            ]
        else:
            datos_fila = [str(start_index + i + 1), "", "", "", "", ""]

        for w, val in zip(col_ws, datos_fila):
            draw_cell_text(c, val, x, y, w, row_h, font_size=6.4, align="center")
            x += w

        y -= row_h

    y -= 6

    nombre_resp, cargo_resp = dividir_responsable(datos["responsable"])

    draw_cell_text(
        c,
        "RESPONSABLE DEL REGISTRO",
        x0,
        y,
        full_w,
        14,
        fill=GRIS_HEADER,
        font_size=6.6,
        bold=True,
        align="center"
    )
    y -= 14

    left_w = 55
    mid_w = full_w - 55 - 80 - 80
    date_label_w = 80
    date_val_w = 80
    rh = 16

    draw_cell_text(c, "NOMBRE", x0, y, left_w, rh, fill=GRIS_CLARO, font_size=6.2, bold=True)
    draw_cell_text(c, nombre_resp, x0 + left_w, y, mid_w, rh, font_size=6.2)

    draw_cell_text(c, "FECHA", x0 + left_w + mid_w, y, date_label_w, rh, fill=GRIS_CLARO, font_size=6.2, bold=True)
    draw_cell_text(c, datos["fecha_general"], x0 + left_w + mid_w + date_label_w, y, date_val_w, rh, font_size=6.2, bold=True, align="center")

    y -= rh

    draw_cell_text(c, "CARGO", x0, y, left_w, rh, fill=GRIS_CLARO, font_size=6.2, bold=True)
    draw_cell_text(c, cargo_resp, x0 + left_w, y, mid_w, rh, font_size=6.2)

    draw_cell_text(c, "FIRMA", x0 + left_w + mid_w, y, date_label_w, rh, fill=GRIS_CLARO, font_size=6.2, bold=True)
    draw_cell_text(c, "", x0 + left_w + mid_w + date_label_w, y, date_val_w, rh, font_size=6.2)


# ==========================================================
# PÁGINAS INDIVIDUALES POR CONDUCTOR
# ==========================================================

def draw_individual(c, fila, df_total, datos):
    W, H = A4
    x0 = 42
    y = H - 42
    table_w = W - 84

    c.setFont("Helvetica-Bold", 13)
    c.setFillColor(colors.black)
    c.drawString(x0, y, TITULO_FORMULARIO)

    y -= 16

    c.setFont("Helvetica", 9)
    c.setFillColor(colors.grey)
    c.drawString(x0, y, SUBTITULO_FORMULARIO)
    c.setFillColor(colors.black)

    y -= 18

    left_w = 230
    right_w = table_w - left_w

    rows = [
        ("RAZÓN SOCIAL O DENOMINACIÓN SOCIAL", RAZON_SOCIAL_FIJA, 28),
        ("RUC", RUC_FIJO, 28),
        ("DOMICILIO ( Dirección, distrito, departamento, provincia)", DIRECCION_FIJA, 44),
        ("TIPO DE ACTIVIDAD ECONÓMICA", ACTIVIDAD_FIJA, 28),
        ("N° TRABAJADORES", NUM_TRABAJADORES_FIJO, 28),
        ("ASISTENCIA A LA CAPACITACIÓN VIRTUAL DE", datos["tema"], 28),
        ("EXPOSITOR", datos["expositor"], 28),
        ("Nombre y Apellido", valor_fila(fila, df_total, "nombre"), 28),
        ("DNI", valor_fila(fila, df_total, "dni"), 28),
        ("Cargo", valor_fila(fila, df_total, "cargo"), 28),
        ("Fecha", datos["fecha_individual"], 28),
        ("Firma", "", 48),
        ("Nombre / Cargo del Responsable del Registro:", datos["responsable"], 28),
        ("Fecha", datos["fecha_individual"], 28),
    ]

    for etiqueta, valor, h in rows:
        draw_cell_text(c, etiqueta, x0, y, left_w, h, font_size=8, bold=True)
        draw_cell_text(c, valor, x0 + left_w, y, right_w, h, font_size=8)
        y -= h


# ==========================================================
# GENERACIÓN PDF COMPLETO
# ==========================================================

def generar_pdf_completo(df_pdf, cliente, tema, fecha, expositor, hora_inicio, hora_termino, responsable):
    df_pdf = filtrar_participantes_validos(df_pdf)

    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)

    datos = obtener_datos_comunes(
        df_pdf,
        tema_manual=tema,
        fecha_manual=fecha,
        expositor_manual=expositor,
        responsable_manual=responsable,
    )

    total = len(df_pdf)
    total_paginas_consolidadas = max(1, math.ceil(total / 20))

    for page_num in range(1, total_paginas_consolidadas + 1):
        inicio = (page_num - 1) * 20
        fin = inicio + 20
        df_chunk = df_pdf.iloc[inicio:fin]

        draw_consolidado(
            c,
            df_chunk,
            df_pdf,
            datos,
            page_num,
            total_paginas_consolidadas,
            hora_inicio,
            hora_termino
        )

        c.showPage()

    for _, fila in df_pdf.iterrows():
        draw_individual(c, fila, df_pdf, datos)
        c.showPage()

    c.save()
    buffer.seek(0)

    return buffer


# ==========================================================
# UTILIDADES DE APP
# ==========================================================

def cargar_y_normalizar(archivo_excel):
    df = leer_excel_inteligente(archivo_excel)
    df = normalizar_columnas(df)
    return df


def obtener_clientes_disponibles(df):
    col_cliente = buscar_columna(df, "cliente")
    if col_cliente is None:
        return []

    clientes = []
    for v in df[col_cliente].tolist():
        txt = limpiar_texto(v).upper()
        if txt and txt not in clientes:
            clientes.append(txt)

    return clientes


def filtrar_por_cliente(df, cliente):
    col_cliente = buscar_columna(df, "cliente")

    if col_cliente is None or cliente == "USAR TODO EL EXCEL":
        return df.copy()

    return df[
        df[col_cliente].apply(normalizar_para_comparar) == normalizar_para_comparar(cliente)
    ].copy()


# ==========================================================
# INTERFAZ STREAMLIT
# ==========================================================

tab1, tab2 = st.tabs(["1. Filtrar Excel", "2. Generar PDF"])

with tab1:
    st.subheader("Filtrar Excel por cliente")

    archivo_excel = st.file_uploader(
        "Sube tu Excel completo de Tally",
        type=["xlsx", "xls", "csv"],
        key="excel_filtro"
    )

    if archivo_excel is not None:
        try:
            df = cargar_y_normalizar(archivo_excel)
            st.success("Excel cargado correctamente.")

            st.write("Columnas detectadas:")
            st.write(list(df.columns))

            st.dataframe(preparar_para_mostrar(df), width="stretch")

            if buscar_columna(df, "cliente") is None:
                st.error(f"No se encontró la columna obligatoria: {COLUMNA_CLIENTE}")
                st.stop()

            cliente_seleccionado = st.selectbox(
                "Selecciona el cliente",
                CLIENTES,
                key="cliente_filtro"
            )

            df_filtrado = filtrar_por_cliente(df, cliente_seleccionado)

            st.write(f"Cliente seleccionado: **{cliente_seleccionado}**")
            st.write(f"Registros encontrados: **{len(df_filtrado)}**")

            if len(df_filtrado) > 0:
                st.dataframe(preparar_para_mostrar(df_filtrado), width="stretch")

                excel_filtrado = convertir_excel_descarga(df_filtrado)

                st.download_button(
                    label="Descargar Excel filtrado",
                    data=excel_filtrado,
                    file_name=f"Excel_filtrado_{cliente_seleccionado.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No hay registros para ese cliente.")

        except Exception as e:
            st.error("Ocurrió un error al leer el Excel.")
            st.exception(e)


with tab2:
    st.subheader("Generar PDF desde Excel filtrado")
    st.write("Sube aquí el Excel filtrado que descargaste en el paso anterior. También puedes subir el Excel completo y elegir cliente.")

    archivo_pdf_excel = st.file_uploader(
        "Sube el Excel para generar el PDF",
        type=["xlsx", "xls", "csv"],
        key="excel_pdf"
    )

    if archivo_pdf_excel is not None:
        try:
            df_pdf = cargar_y_normalizar(archivo_pdf_excel)
            df_pdf = filtrar_participantes_validos(df_pdf)

            clientes_disponibles = obtener_clientes_disponibles(df_pdf)
            opciones_cliente = ["USAR TODO EL EXCEL"] + [c for c in CLIENTES if c in clientes_disponibles]

            if len(opciones_cliente) == 1 and clientes_disponibles:
                opciones_cliente += clientes_disponibles

            cliente_pdf = st.selectbox(
                "Cliente para el PDF",
                opciones_cliente,
                key="cliente_pdf"
            )

            df_pdf_filtrado = filtrar_por_cliente(df_pdf, cliente_pdf)
            df_pdf_filtrado = filtrar_participantes_validos(df_pdf_filtrado)

            st.write(f"Registros que entrarán al PDF: **{len(df_pdf_filtrado)}**")
            st.dataframe(preparar_para_mostrar(df_pdf_filtrado), width="stretch")

            if len(df_pdf_filtrado) == 0:
                st.warning("No hay registros para generar PDF.")
                st.stop()

            tema_detectado = primer_valor(df_pdf_filtrado, "tema", "")
            fecha_detectada = primer_valor(df_pdf_filtrado, "fecha_capacitacion", "")
            expositor_detectado = primer_valor(df_pdf_filtrado, "expositor", "")
            responsable_detectado = primer_valor(df_pdf_filtrado, "responsable", "")

            st.markdown("### Datos del formato")
            st.write("Estos datos se toman del Excel. Revisa antes de generar el PDF.")

            col1, col2 = st.columns(2)

            with col1:
                tema_pdf = st.text_input("Tema", value=tema_detectado)
                expositor_pdf = st.text_input("Expositor", value=expositor_detectado)
                hora_inicio = st.text_input("Hora de inicio", value="8:30 pm")

            with col2:
                fecha_pdf = st.text_input("Fecha de capacitación", value=fecha_ddmmyyyy(fecha_detectada))
                responsable_pdf = st.text_input("Responsable del registro", value=responsable_detectado)
                hora_termino = st.text_input("Hora de término", value="09:00 pm")

            st.info("Las firmas todavía no se insertan. El PDF dejará los espacios de firma en blanco.")

            if st.button("Generar PDF correspondiente"):
                pdf_generado = generar_pdf_completo(
                    df_pdf_filtrado,
                    cliente_pdf,
                    tema_pdf,
                    fecha_pdf,
                    expositor_pdf,
                    hora_inicio,
                    hora_termino,
                    responsable_pdf,
                )

                nombre_cliente_archivo = (
                    cliente_pdf
                    .replace("USAR TODO EL EXCEL", "GENERAL")
                    .replace(" ", "_")
                    .replace("Á", "A")
                )

                st.success("PDF generado correctamente.")

                st.download_button(
                    label="Descargar PDF generado",
                    data=pdf_generado,
                    file_name=f"Formato_BlueLine_{nombre_cliente_archivo}.pdf",
                    mime="application/pdf"
                )

        except Exception as e:
            st.error("Ocurrió un error al generar el PDF.")
            st.exception(e)