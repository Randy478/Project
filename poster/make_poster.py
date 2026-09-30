"""Erzeugt das A3-Poster "Die geneigte Ebene" als .pptx (fuer den Import in Canva)."""
import math

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Mm, Pt

TEXT = RGBColor(0x22, 0x22, 0x22)
BLUE = RGBColor(0x1F, 0x6F, 0xB2)
LIGHT = RGBColor(0xEA, 0xF2, 0xFA)
GREY = RGBColor(0x88, 0x88, 0x88)
RED = RGBColor(0xC0, 0x39, 0x2B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT = "Arial"

prs = Presentation()
prs.slide_width, prs.slide_height = Mm(297), Mm(420)
slide = prs.slides.add_slide(prs.slide_layouts[6])


def text(x, y, w, h, paras, size=16, color=TEXT, bold=False, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP):
    """paras: Liste von Strings oder (String, dict) mit Abweichungen."""
    box = slide.shapes.add_textbox(Mm(x), Mm(y), Mm(w), Mm(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Mm(1)
    for i, p in enumerate(paras):
        s, opt = (p, {}) if isinstance(p, str) else p
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = opt.get("align", align)
        para.space_after = Pt(opt.get("after", 4))
        run = para.add_run()
        run.text = s
        run.font.name = FONT
        run.font.size = Pt(opt.get("size", size))
        run.font.bold = opt.get("bold", bold)
        run.font.color.rgb = opt.get("color", color)
    return box


def rect(x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, dash=False):
    s = slide.shapes.add_shape(shape, Mm(x), Mm(y), Mm(w), Mm(h))
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(1.5)
        if dash:
            s.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    style = s._element.find(qn("p:style"))
    if style is not None:  # Theme-Schatten entfernen
        s._element.remove(style)
    return s


def arrow(x1, y1, x2, y2, color, width=3):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Mm(x1), Mm(y1), Mm(x2), Mm(y2))
    c.line.color.rgb = color
    c.line.width = Pt(width)
    ln = c.line._get_or_add_ln()
    tail = etree.SubElement(ln, qn("a:tailEnd"))
    tail.set("type", "triangle")
    tail.set("w", "med")
    tail.set("len", "med")
    return c


def heading(x, y, w, s):
    text(x, y, w, 12, [s], size=26, bold=True, color=BLUE)


M = 15          # Rand
W = 297 - 2 * M  # nutzbare Breite
COL = (W - 8) / 2

# ---------- Kopf ----------
text(M, 12, W, 25, ["Die geneigte Ebene"], size=60, bold=True)
text(M, 36, W, 12, ["Goldene Regel der Mechanik"], size=30, color=BLUE)
text(M, 49, W, 8, ["Can, Antonio, Florian, Philly  ·  MCG"], size=14, color=GREY)
rect(M, 59, W, 1.2, fill=BLUE)

# ---------- Was ist eine geneigte Ebene? ----------
y = 66
heading(M, y, COL, "Was ist das?")
text(M, y + 14, COL, 50, [
    "• eine schräge Rampe",
    "• ein Körper wird hochgezogen statt senkrecht gehoben",
    "• kraftumformende Einrichtung: man braucht weniger Kraft",
    "• Beispiel: Passstraße mit vielen Kurven",
])

# ---------- Goldene Regel ----------
x2 = M + COL + 8
rect(x2, y, COL, 72, fill=LIGHT)
heading(x2 + 4, y + 3, COL - 8, "Die Goldene Regel")
text(x2 + 4, y + 16, COL - 8, 18,
     ["„Was man an Kraft spart, muss man mit Strecke bezahlen.“"],
     size=19, bold=True)
text(x2 + 4, y + 36, COL - 8, 36, [
    "• lange Strecke → kleine Kraft",
    "• kurze Strecke → große Kraft",
    "• die Arbeit W = F · s bleibt gleich",
    "• Energie spart man nicht!",
])

# ---------- Skizze Rampe ----------
y = 146
heading(M, y, W, "So funktioniert es")
tx, ty, tw, th = M + 20, y + 22, 120, 45        # Dreieck (Rampe)
ramp = rect(tx, ty, tw, th, fill=RGBColor(0xD9, 0xB9, 0x8C), shape=MSO_SHAPE.RIGHT_TRIANGLE)
# Dreieck spiegeln, damit die Rampe nach rechts ansteigt
ramp._element.spPr.find(qn("a:xfrm")).set("flipH", "1")
rect(tx - 5, ty + th, tw + 15, 2, fill=GREY)  # Boden

ang = math.degrees(math.atan(th / tw))
# Wagen auf der Rampe (bei 45 % der Strecke)
t = 0.45
cx, cy = tx + t * tw, ty + th - t * th
cw, ch = 18, 8
off = ch / 2 + 0.5
px = cx - math.sin(math.radians(ang)) * off
py = cy - math.cos(math.radians(ang)) * off
car = rect(px - cw / 2, py - ch / 2, cw, ch, fill=BLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
car.rotation = -ang
# Kraftpfeil entlang der Rampe
d = 32
arrow(px + math.cos(math.radians(ang)) * 10, py - math.sin(math.radians(ang)) * 10,
      px + math.cos(math.radians(ang)) * d, py - math.sin(math.radians(ang)) * d, RED)
text(px + 14, py - 22, 30, 8, ["Kraft F"], size=16, bold=True, color=RED)
# Strecke s
s_lbl = text(tx + tw * 0.55, ty + th * 0.55, 40, 8, ["Strecke s"], size=16, bold=True)
s_lbl.rotation = -ang
# Hoehe h
arrow(tx + tw + 6, ty + th, tx + tw + 6, ty, TEXT, width=1.5)
text(tx + tw + 8, ty + th / 2 - 4, 30, 8, ["Höhe h"], size=16, bold=True)

# Rechenbeispiel
text(M + 180, y + 16, W - 180, 55, [
    ("Rechenbeispiel:", {"bold": True}),
    "Wagen (1,0 kg) auf 1,0 m Höhe",
    ("direkt hoch:", {"bold": True, "after": 0}),
    "9,8 N · 1,0 m = 9,8 J",
    ("über die Rampe:", {"bold": True, "after": 0}),
    "4,1 N · 2,4 m ≈ 9,8 J",
    ("→ gleiche Arbeit!", {"bold": True, "color": BLUE}),
], size=16)

# ---------- Experiment ----------
y = 225
rect(M, y - 3, W, 1.2, fill=BLUE)
heading(M, y + 2, W, "Experiment: Rampen im Vergleich")
text(M, y + 16, 120, 80, [
    ("Aufbau", {"bold": True}),
    "• Wagen (1,0 kg), Rampe, Kraftmesser",
    "• Rampe endet immer in 1,0 m Höhe",
    ("Durchführung", {"bold": True}),
    "• Wagen langsam die Rampe hochziehen",
    "• Kraft F ablesen, Strecke s messen",
    "• mit verschieden steilen Rampen wiederholen",
    "• jedes Mal F · s ausrechnen",
])

# Messwerte-Tabelle
rows = [("s in m", "F in N", "F · s in N·m"),
        ("2,4", "4,1", "9,8"), ("2,0", "4,9", "9,8"), ("1,6", "6,2", "9,9"),
        ("1,2", "8,1", "9,7"), ("1,0", "9,9", "9,9")]
tbl = slide.shapes.add_table(len(rows), 3, Mm(M + 130), Mm(y + 17), Mm(W - 130), Mm(54)).table
for r, row in enumerate(rows):
    for c, val in enumerate(row):
        cell = tbl.cell(r, c)
        cell.text = val
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        f = p.runs[0].font
        f.name, f.size, f.bold = FONT, Pt(15), r == 0
        f.color.rgb = WHITE if r == 0 else TEXT
        cell.fill.solid()
        cell.fill.fore_color.rgb = BLUE if r == 0 else (LIGHT if r % 2 else WHITE)
text(M + 130, y + 72, W - 130, 6, ["Messwerte aus [1], S. 91"], size=11, color=GREY)

text(M + 130, y + 80, W - 130, 20, [
    ("Ergebnis: F · s ist immer ca. 9,8 N·m", {"bold": True}),
    "→ Goldene Regel bestätigt!",
], size=16)

# ---------- Alltag ----------
y = 330
rect(M, y - 3, W, 1.2, fill=BLUE)
heading(M, y + 2, W, "Im Alltag")
for i, cap in enumerate([
        "Bild 1: Passstraße – lange Strecke, weniger Kraft",
        "Bild 2: Rampe – lange Strecke, weniger Kraft"]):
    bx = M + i * (COL + 8)
    ph = rect(bx, y + 15, COL, 50, line=GREY, dash=True)
    tf = ph.text_frame
    tf.text = "eigenes Foto hier einfügen"
    run = tf.paragraphs[0].runs[0]
    run.font.name, run.font.size, run.font.color.rgb = FONT, Pt(14), GREY
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    text(bx, y + 66, COL, 8, [cap], size=14)

# ---------- Quellen ----------
text(M, 405, W, 10, [
    "Quellen: [1] Universum Physik (Schulbuch), S. 90–93  ·  Fotos: eigene Aufnahmen"],
    size=11, color=GREY)

prs.save("/home/user/Project/poster/Poster_Geneigte_Ebene.pptx")
print("ok")
