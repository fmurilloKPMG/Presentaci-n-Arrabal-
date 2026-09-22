# -*- coding: utf-8 -*-
"""Comité de seguimiento Arrabal — 3 diapositivas: estado, calendario, comentarios.
   Regenera Comite_Arrabal_2026-09-22.pptx a partir de estos datos.
   Fuente: Report_semanal_Arrabal_2026-09-11_avance.xlsx (hojas "Plan tareas" y "Gantt avance").
"""
import datetime
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

SRC = "_plantilla_KPMG_base.pptx"
DST = "Comite_Arrabal_2026-09-22.pptx"

# ── Paleta KPMG (fijada en el prompt, no negociable) ──
KPMG_BLUE   = RGBColor(0x00, 0x33, 0x8D)
COBALT      = RGBColor(0x1E, 0x49, 0xE2)
LIGHT_BLUE  = RGBColor(0xBF, 0xD0, 0xF8)
RULE        = RGBColor(0xD3, 0xDA, 0xE8)
INK         = RGBColor(0x0C, 0x23, 0x3C)
INK2        = RGBColor(0x3F, 0x51, 0x6B)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
VERDE       = RGBColor(0x0E, 0x7A, 0x4F)   # En plan
AMBAR       = RGBColor(0xB8, 0x86, 0x0B)   # Con desviación
ROJO        = RGBColor(0xA6, 0x19, 0x2E)   # En riesgo
GRIS        = RGBColor(0x9A, 0xA8, 0xC0)   # No iniciada / pendiente

FONT = "Arial"

# ── Datos reales, calculados desde Report_semanal_Arrabal_2026-09-11 ──
INICIATIVAS = [
    {
        "codigo": "RPR03", "nombre": "Recopilación automática de nuevos proyectos",
        "que_va": "Automatiza la captación de convocatorias y ayudas desde fuentes públicas y privadas",
        "estado": "En riesgo", "color": ROJO, "avance": 0.64,
        "fecha_ini": datetime.date(2026,7,14), "fecha_fin": datetime.date(2026,10,9),
        "hitos_riesgo": [datetime.date(2026,7,22), datetime.date(2026,9,4)],
        "comentario": (
            "43 tareas planificadas: 21 completadas, 6 en curso, 14 sin iniciar. "
            "Dos tareas bloqueadas y con vencimiento ya pasado a fecha del informe (11 sept.): "
            "la validación del entorno del cliente (venció 22 jul.) y la integración con "
            "Salesforce y el modelo CRM (venció 4 sept.). Ambas bloquean el arranque de las "
            "fuentes de las Fases 3 y 4 (14 fuentes, todas sin iniciar). "
            "Necesita decisión del comité: desbloqueo de acceso al entorno cliente y a Salesforce."
        ),
    },
    {
        "codigo": "RPR04", "nombre": "Overview económico y administrativo",
        "que_va": "Cuadro de mando económico-administrativo accesible y trazable para los proyectos",
        "estado": "Con desviación", "color": AMBAR, "avance": 0.93,
        "fecha_ini": datetime.date(2026,7,27), "fecha_fin": datetime.date(2026,8,28),
        "hitos_riesgo": [],
        "comentario": (
            "3 de 4 tareas completadas (modelo económico, automatizaciones y reporting). "
            "Queda la validación final con los equipos implicados, al 70 % y por encima de su "
            "fecha prevista (28 ago.). Sin bloqueos técnicos: pendiente de agenda de validación."
        ),
    },
    {
        "codigo": "RG10", "nombre": "Optimización y documentación del CRM",
        "que_va": "Mejora el modelo CRM y documenta su uso para el equipo",
        "estado": "Con desviación", "color": AMBAR, "avance": 0.94,
        "fecha_ini": datetime.date(2026,7,15), "fecha_fin": datetime.date(2026,8,28),
        "hitos_riesgo": [],
        "comentario": (
            "4 de 5 tareas completadas (revisión inicial, modelo, automatizaciones y "
            "documentación). Igual que RPR04: la validación final y puesta en uso está al "
            "55 % y superó su fecha prevista (28 ago.). Pendiente de cierre con cliente."
        ),
    },
    {
        "codigo": "RC08", "nombre": "Automatización para el seguimiento de usuarios",
        "que_va": "[por confirmar]",
        "estado": "Por confirmar", "color": GRIS, "avance": None,
        "fecha_ini": None, "fecha_fin": None,
        "hitos_riesgo": [],
        "comentario": (
            "[por confirmar] — no incluida en el informe de origen (Report_semanal_Arrabal "
            "2026-09-11). Sin plan de tareas ni fecha de inicio registrados."
        ),
    },
]

HOY = datetime.date(2026, 9, 22)

# ══════════════════════════════════════════════════════════════
def set_font(run, size=12, bold=False, color=INK, italic=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color

def add_text(slide, x, y, w, h, text, size=12, bold=False, color=INK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, wrap=True, italic=False,
             line_spacing=1.15):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = line_spacing
    r = p.add_run()
    r.text = text
    set_font(r, size, bold, color, italic)
    return box

def add_pill(slide, x, y, w, h, text, fill_color, text_color=WHITE, size=11):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.adjustments[0] = 0.5
    shp.fill.solid(); shp.fill.fore_color.rgb = fill_color
    shp.line.fill.background()
    shp.shadow.inherit = False
    tf = shp.text_frame
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = text
    set_font(r, size, True, text_color)
    return shp

def add_pill_outline(slide, x, y, w, h, text, edge_color, text_color):
    """Pastilla 'por confirmar': contorno discontinuo, sin relleno."""
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.adjustments[0] = 0.5
    shp.fill.background()
    shp.line.color.rgb = edge_color
    shp.line.width = Pt(1.25)
    shp.line.dash_style = 3  # dashed (MSO line dash)
    shp.shadow.inherit = False
    tf = shp.text_frame
    tf.margin_left = tf.margin_right = Emu(0)
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = text
    set_font(r, 11, True, text_color)
    return shp

def add_rect(slide, x, y, w, h, color, line=False):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid(); shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp

def add_rounded(slide, x, y, w, h, color, radius=0.06):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.adjustments[0] = radius
    shp.fill.solid(); shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp

def add_line(slide, x1, y1, x2, y2, color, weight=0.75, dash=None):
    ln = slide.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    ln.line.color.rgb = color
    ln.line.width = Pt(weight)
    if dash:
        ln.line.dash_style = dash
    ln.shadow.inherit = False
    return ln

def add_diamond(slide, cx, cy, size, color, outline=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.DIAMOND, Inches(cx - size/2), Inches(cy - size/2), Inches(size), Inches(size))
    shp.fill.solid(); shp.fill.fore_color.rgb = color
    if outline:
        shp.line.color.rgb = outline; shp.line.width = Pt(1.5)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp

def set_slide_title(slide, text):
    slide.shapes.title.text_frame.text = text
    r = slide.shapes.title.text_frame.paragraphs[0].runs[0]
    set_font(r, 28, True, KPMG_BLUE)

def footer(slide, texto_izq="Comité de seguimiento · Arrabal", num=""):
    add_text(slide, 0.6, 7.16, 8, 0.28, texto_izq, size=9, color=INK2)
    if num:
        add_text(slide, 12.3, 7.16, 0.43, 0.28, num, size=9, color=INK2, align=PP_ALIGN.RIGHT)

# ══════════════════════════════════════════════════════════════
prs = Presentation(SRC)

# Eliminar todas las diapositivas de contenido existentes (conservando master/layouts).
# drop_rel() quita también la parte del paquete; si solo se quita del sldIdLst,
# la parte huérfana se queda dentro del .pptx y python-pptx puede reasignar su
# mismo nombre a una diapositiva nueva -> zip con entradas duplicadas.
xml_slides = prs.slides._sldIdLst
for sld in list(xml_slides):
    rId = sld.get(qn('r:id'))
    prs.part.drop_rel(rId)
    xml_slides.remove(sld)

layout_content = prs.slide_layouts[7]   # "Title Only"

# ─────────────────────────────────────────────────────────────
# DIAPOSITIVA 1 — Estado de las iniciativas
# ─────────────────────────────────────────────────────────────
s1 = prs.slides.add_slide(layout_content)
set_slide_title(s1, "Estado de las iniciativas")

col_x = {"cod": 0.6, "nom": 1.9, "que": 4.5, "est": 9.5, "av": 11.35}
col_w = {"cod": 1.3, "nom": 2.6, "que": 5.0, "est": 1.85, "av": 1.38}
header_y = 1.35
row_h = 1.02
row_gap = 0.14
row0_y = header_y + 0.42

headers = [("cod","CÓDIGO"), ("nom","INICIATIVA"), ("que","DE QUÉ VA"), ("est","ESTADO"), ("av","AVANCE")]
for key, label in headers:
    add_text(s1, col_x[key], header_y, col_w[key], 0.32, label, size=10.5, bold=True, color=INK2)
add_line(s1, 0.6, header_y+0.38, 12.73, header_y+0.38, RULE, weight=1.25)

for i, ini in enumerate(INICIATIVAS):
    y = row0_y + i*(row_h+row_gap)
    add_line(s1, 0.6, y+row_h, 12.73, y+row_h, RULE, weight=0.75)

    add_text(s1, col_x["cod"], y, col_w["cod"], row_h, ini["codigo"], size=14, bold=True, color=KPMG_BLUE, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s1, col_x["nom"], y, col_w["nom"], row_h, ini["nombre"], size=12, bold=True, color=INK, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s1, col_x["que"], y, col_w["que"], row_h, ini["que_va"], size=11.5, color=INK2, anchor=MSO_ANCHOR.MIDDLE)

    pill_w, pill_h = 1.75, 0.36
    py = y + (row_h - pill_h)/2
    if ini["estado"] == "Por confirmar":
        add_pill_outline(s1, col_x["est"], py, pill_w, pill_h, "POR CONFIRMAR", GRIS, INK2)
    else:
        add_pill(s1, col_x["est"], py, pill_w, pill_h, ini["estado"].upper(), ini["color"])

    if ini["avance"] is not None:
        bar_w, bar_h = 1.0, 0.14
        bx, by = col_x["av"], y + row_h*0.35
        add_rounded(s1, bx, by, bar_w, bar_h, RULE, radius=0.5)
        add_rounded(s1, bx, by, bar_w*ini["avance"], bar_h, KPMG_BLUE, radius=0.5)
        add_text(s1, bx, by+bar_h+0.04, bar_w+0.3, 0.24, f'{ini["avance"]*100:.0f} %', size=11, bold=True, color=INK)
    else:
        add_text(s1, col_x["av"], y, col_w["av"], row_h, "—", size=13, bold=True, color=GRIS, anchor=MSO_ANCHOR.MIDDLE)

legend_y = row0_y + 4*(row_h+row_gap) + 0.06
add_text(s1, 0.6, legend_y, 12, 0.3,
         "Estado calculado sobre tareas bloqueadas o con fecha vencida a 11 sept. 2026, no solo sobre el campo “Estado” del reporte de origen.",
         size=9.5, italic=True, color=INK2)
footer(s1, num="1")

# ─────────────────────────────────────────────────────────────
# DIAPOSITIVA 2 — Calendario
# ─────────────────────────────────────────────────────────────
s2 = prs.slides.add_slide(layout_content)
set_slide_title(s2, "Calendario")

axis_start = datetime.date(2026, 7, 1)
axis_end   = datetime.date(2026, 10, 15)
chart_x0, chart_x1 = 2.3, 12.73
chart_w = chart_x1 - chart_x0
total_days = (axis_end - axis_start).days

def x_for(d):
    return chart_x0 + chart_w * (d - axis_start).days / total_days

band_top = 1.55
band_h = 0.62
band_gap = 0.30
months = [(datetime.date(2026,7,1),"JUL"), (datetime.date(2026,8,1),"AGO"),
          (datetime.date(2026,9,1),"SEP"), (datetime.date(2026,10,1),"OCT")]

grid_bottom = band_top + 4*(band_h+band_gap) - band_gap + 0.15
for md, label in months:
    mx = x_for(md)
    add_line(s2, mx, band_top-0.15, mx, grid_bottom, RULE, weight=0.75, dash=2)
    add_text(s2, mx, band_top-0.42, 1.0, 0.24, label, size=10, bold=True, color=INK2)

for i, ini in enumerate(INICIATIVAS):
    y = band_top + i*(band_h+band_gap)
    add_text(s2, 0.6, y, 1.55, band_h, ini["codigo"], size=13, bold=True, color=KPMG_BLUE, anchor=MSO_ANCHOR.MIDDLE)

    if ini["fecha_ini"] is None:
        add_text(s2, chart_x0, y, chart_w, band_h,
                  "Sin fechas registradas — [por confirmar]", size=10.5, italic=True,
                  color=GRIS, anchor=MSO_ANCHOR.MIDDLE)
        continue

    x0 = x_for(ini["fecha_ini"]); x1 = x_for(ini["fecha_fin"])
    band_y = y + (band_h - 0.30)/2
    total_w = x1 - x0
    add_rounded(s2, x0, band_y, total_w, 0.30, LIGHT_BLUE, radius=0.5)
    add_rounded(s2, x0, band_y, total_w*ini["avance"], 0.30, KPMG_BLUE, radius=0.5)

    add_diamond(s2, x0, band_y+0.15, 0.14, KPMG_BLUE)
    add_diamond(s2, x1, band_y+0.15, 0.14, KPMG_BLUE)

    for hr in ini["hitos_riesgo"]:
        hx = x_for(hr)
        add_diamond(s2, hx, band_y+0.15, 0.20, ROJO, outline=WHITE)

    add_text(s2, x0, band_y+0.32, total_w+0.5, 0.2, f'{ini["avance"]*100:.0f} %', size=9.5, bold=True, color=INK2)

hoy_x = x_for(HOY)
add_line(s2, hoy_x, band_top-0.15, hoy_x, grid_bottom, AMBAR, weight=1.75)
add_text(s2, hoy_x-0.55, grid_bottom+0.03, 1.4, 0.22, "HOY · 22 sept.", size=9.5, bold=True, color=AMBAR, align=PP_ALIGN.CENTER)

leg_y = grid_bottom + 0.45
add_rounded(s2, 0.6, leg_y+0.03, 0.35, 0.14, KPMG_BLUE, radius=0.5)
add_text(s2, 1.05, leg_y-0.04, 1.6, 0.26, "Ejecutado", size=10, color=INK2)
add_rounded(s2, 2.6, leg_y+0.03, 0.35, 0.14, LIGHT_BLUE, radius=0.5)
add_text(s2, 3.05, leg_y-0.04, 1.6, 0.26, "Previsto", size=10, color=INK2)
add_diamond(s2, 4.85, leg_y+0.10, 0.16, ROJO, outline=WHITE)
add_text(s2, 5.05, leg_y-0.04, 3.5, 0.26, "Tarea bloqueada / fecha vencida", size=10, color=INK2)
footer(s2, num="2")

# ─────────────────────────────────────────────────────────────
# DIAPOSITIVA 3 — Comentarios sobre las iniciativas
# ─────────────────────────────────────────────────────────────
s3 = prs.slides.add_slide(layout_content)
set_slide_title(s3, "Comentarios sobre las iniciativas")

block_y = 1.35
block_h = 1.32
block_gap = 0.10

for i, ini in enumerate(INICIATIVAS):
    y = block_y + i*(block_h+block_gap)
    add_rect(s3, 0.6, y, 0.05, block_h-0.12, ini["color"] if ini["estado"]!="Por confirmar" else GRIS)
    add_text(s3, 0.85, y, 1.55, 0.3, ini["codigo"], size=14, bold=True, color=KPMG_BLUE)
    add_text(s3, 0.85, y+0.30, 1.55, 0.55, ini["nombre"], size=9.5, color=INK2, wrap=True, line_spacing=1.05)
    add_text(s3, 2.55, y, 10.2, block_h-0.12, ini["comentario"], size=11, color=INK, wrap=True, line_spacing=1.12,
              italic=(ini["estado"]=="Por confirmar"))

footer(s3, num="3")

# ══════════════════════════════════════════════════════════════
prs.save(DST)
print("Guardado:", DST)
print("Diapositivas:", len(prs.slides._sldIdLst))
