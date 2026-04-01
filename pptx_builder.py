"""
IBM Z Newsletter – PPTX Builder (vollständig programmatisch)

Alle Folien werden von Grund auf gebaut.
Aus dem Template werden nur die IBM-Bilder (Banner, Logos) entnommen.

Struktur:
  Folie 1  : Cover  – linke Seitenleiste (TOC) + rechts Artikel 1
  Folie 2+ : Inhaltsfolien – 2 Artikel pro Folie
  Letzte   : Abschlussfolie aus Template (unverändert)
"""

import copy, io, os, re, shutil
from typing import List, Optional

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt, Emu

from config import OUTPUT_DIR, TEMPLATE_FILE

# ── Farben ────────────────────────────────────────────────────────────────────
IBM_BLUE   = RGBColor(0x00, 0x43, 0xCE)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
NEAR_BLACK = RGBColor(0x16, 0x16, 0x16)
MID_GRAY   = RGBColor(0x6F, 0x6F, 0x6F)
DIV_GRAY   = RGBColor(0xD8, 0xD8, 0xD8)
BAR_GRAY   = RGBColor(0xE7, 0xE6, 0xE6)   # Header/Footer/Sidebar-Farbe

FONT = "IBM Plex Sans"

GERMAN_MONTHS = {
    1:"Januar", 2:"Februar", 3:"März",    4:"April",
    5:"Mai",    6:"Juni",    7:"Juli",     8:"August",
    9:"September", 10:"Oktober", 11:"November", 12:"Dezember",
}

# ── Geometrie (A4 Hochformat) ─────────────────────────────────────────────────
W, H          = Inches(8.27), Inches(11.69)
HEADER_H      = Inches(0.66)
FOOTER_Y      = Inches(11.13)
FOOTER_H      = Inches(0.56)

# Seitenleiste (Cover)
SIDE_W        = Inches(2.80)
SIDE_COL_L    = Inches(0.14)
SIDE_COL_W    = SIDE_W - SIDE_COL_L - Inches(0.10)

# Inhaltsbereich
CONTENT_TOP   = Inches(0.82)
CONTENT_BOT   = Inches(10.95)
MARGIN        = Inches(0.28)
NUM_W         = Inches(0.38)   # Zahlen-Spalte
TEXT_L_FULL   = MARGIN + NUM_W + Inches(0.06)   # Text ab hier (Inhaltsfolien)
TEXT_W_FULL   = W - TEXT_L_FULL - Inches(0.22)

# Text-Spalte auf Cover (rechte Seite)
COVER_R_L     = SIDE_W + Inches(0.20)
COVER_R_W     = W - COVER_R_L - Inches(0.20)


# ═══════════════════════════════════════════════════════════════════════════════
# Hilfsfunktionen
# ═══════════════════════════════════════════════════════════════════════════════

def _rect(slide, l, t, w, h, color):
    s = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, l, t, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = color; s.line.fill.background()
    return s


def _tb(slide, l, t, w, h, text, size,
        bold=False, color=NEAR_BLACK, align=PP_ALIGN.LEFT,
        wrap=True, italic=False):
    """Einfaches Textfeld."""
    box = slide.shapes.add_textbox(l, t, w, h)
    tf  = box.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = color; r.font.name = FONT
    return box


def _multi_para_tb(slide, l, t, w, h, text, size, color=NEAR_BLACK):
    """Textfeld mit Absätzen – getrennt durch Leerzeilen im Text."""
    box = slide.shapes.add_textbox(l, t, w, h)
    tf  = box.text_frame
    tf.word_wrap = True
    paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]
    if not paragraphs:
        paragraphs = [text.strip()]
    for i, para_text in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if i > 0:
            p.space_before = Pt(5)
        r = p.add_run()
        r.text = para_text
        r.font.size = Pt(size)
        r.font.color.rgb = color
        r.font.name = FONT
    return box


def _link_tb(slide, l, t, w, h, text, url):
    """Textfeld mit Hyperlink."""
    box = slide.shapes.add_textbox(l, t, w, h)
    tf  = box.text_frame; tf.word_wrap = False
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = text
    r.font.size = Pt(9); r.font.color.rgb = IBM_BLUE
    r.font.underline = True; r.font.name = FONT
    try:
        rId = slide.part.relate_to(
            url,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            is_external=True)
        rPr = r._r.get_or_add_rPr()
        hl = etree.SubElement(rPr, qn("a:hlinkClick"))
        hl.set(qn("r:id"), rId)
    except Exception:
        pass
    return box


def _divider(slide, y):
    _rect(slide, MARGIN, y, W - MARGIN * 2, Inches(0.015), DIV_GRAY)


def _picture(slide, img_bytes, l, t, w, h):
    if img_bytes:
        slide.shapes.add_picture(io.BytesIO(img_bytes), l, t, w, h)


def _get_image_bytes(article: dict):
    """Gibt vorgeladene Bild-Bytes zurück (oder None wenn kein Bild)."""
    b = article.get("image_bytes")
    return b if b else None


# ── Bilder aus Template extrahieren ──────────────────────────────────────────

def _extract_images(prs):
    """
    Gibt zurück:
      banner  – breites IBM-Foto oben auf der Titelfolie
      logo_hd – IBM-Logo (klein, oben rechts, aus Inhaltsfolie)
    """
    banner = logo_hd = None

    # Banner: größtes Bild auf Slide 0 (Grafik 4)
    for shape in prs.slides[0].shapes:
        if shape.shape_type == 13 and shape.width > Inches(7.0):
            banner = shape.image.blob
            break

    # Logo: kleines Bild oben rechts auf einer Inhaltsfolie
    for slide in list(prs.slides)[1:]:
        for shape in slide.shapes:
            if (shape.shape_type == 13
                    and shape.top < Inches(1.0)
                    and shape.left > Inches(6.0)):
                logo_hd = shape.image.blob
                break
        if logo_hd:
            break

    return banner, logo_hd


# ══════════════════════════════════════════════════════════════════════════════
# Header & Footer
# ══════════════════════════════════════════════════════════════════════════════

def _header(slide, logo_bytes):
    _rect(slide, Inches(-0.02), Inches(0), W + Inches(0.04), HEADER_H, BAR_GRAY)
    _picture(slide, logo_bytes, Inches(7.16), Inches(0.17), Inches(0.82), Inches(0.33))


def _footer(slide, logo_bytes, page_num):
    _rect(slide, Inches(-0.02), FOOTER_Y, W + Inches(0.04), FOOTER_H, BAR_GRAY)
    _tb(slide, Inches(0.10), Inches(11.44), Inches(2.20), Inches(0.22),
        "© 2026 IBM Corporation", 6, color=WHITE)
    _tb(slide, Inches(3.90), Inches(11.36), Inches(0.50), Inches(0.28),
        str(page_num), 8, color=WHITE, align=PP_ALIGN.CENTER)
    _picture(slide, logo_bytes, Inches(7.16), Inches(11.31), Inches(0.82), Inches(0.33))


# ══════════════════════════════════════════════════════════════════════════════
# Artikel-Block (wird auf Inhaltsfolien und Cover verwendet)
# ══════════════════════════════════════════════════════════════════════════════

def _article_block(slide, article, num, top, block_h,
                   l_num, l_text, w_text, num_color=IBM_BLUE, text_color=NEAR_BLACK):
    """
    Zeichnet einen Artikel-Block:
      – Zahl + Titel oben
      – Zusammenfassung (Absätze) darunter
      – Artikelbild (falls vorhanden) danach
      – Link + Autor/Datum am unteren Rand des Blocks verankert
    """
    pad = Inches(0.10)
    y   = top + pad

    # ── Zahl ──────────────────────────────────────────────────────────────────
    _tb(slide, l_num, y, NUM_W, Inches(0.45),
        str(num), 16, bold=True, color=num_color, align=PP_ALIGN.CENTER)

    # ── Titel ─────────────────────────────────────────────────────────────────
    title_h = Inches(0.55)
    _tb(slide, l_text, y, w_text, title_h,
        article["title"], 11, bold=True, color=text_color)

    # ── Link + Autor am unteren Rand verankert ────────────────────────────────
    author_top = top + block_h - Inches(0.28)
    link_top   = author_top - Inches(0.30)

    pub = article.get("published")
    date_str = (f"{pub.day}. {GERMAN_MONTHS.get(pub.month,'')} {pub.year}"
                if pub else "")
    _link_tb(slide, l_text, link_top, w_text, Inches(0.26),
             "→ Mehr Informationen erhalten Sie hier",
             article.get("url", ""))
    _tb(slide, l_text, author_top, w_text, Inches(0.26),
        f"{article.get('author','')}  ·  {date_str}",
        8, color=MID_GRAY, italic=True)

    # ── Verfügbarer Bereich für Text + Bild ───────────────────────────────────
    sum_top    = y + title_h + Inches(0.08)
    content_h  = link_top - sum_top - Inches(0.10)   # Platz zw. Titel und Link

    # ── Artikelbild (bereits beim Scrapen geladen) ────────────────────────────
    img_bytes = _get_image_bytes(article)

    MAX_IMG_H = min(Inches(1.60), content_h * 0.38)
    MAX_IMG_W = w_text

    if img_bytes:
        try:
            # Natürliche Größe ermitteln (keine Dimension vorgeben)
            tmp = slide.shapes.add_picture(io.BytesIO(img_bytes), l_text, Inches(0))
            nat_w, nat_h = tmp.width, tmp.height
            tmp._element.getparent().remove(tmp._element)

            # Proportional verkleinern wenn nötig – nie vergrößern, nie strecken
            scale = min(1.0, MAX_IMG_W / nat_w, MAX_IMG_H / nat_h)
            img_w = int(nat_w * scale)
            img_h = int(nat_h * scale)

            sum_h   = max(content_h - img_h - Inches(0.12), Inches(0.80))
            img_top = sum_top + sum_h + Inches(0.10)
            # Textboxen haben ~0.05" internen Innenabstand links → Bild angleichen
            img_left = l_text + Inches(0.05)
            slide.shapes.add_picture(io.BytesIO(img_bytes),
                                     img_left, img_top, img_w, img_h)
        except Exception:
            img_bytes = None

    if not img_bytes:
        sum_h = max(content_h, Inches(0.80))

    # ── Zusammenfassung (Absätze) ─────────────────────────────────────────────
    _multi_para_tb(slide, l_text, sum_top, w_text, sum_h,
                   article["summary"], 9, color=text_color)


# ══════════════════════════════════════════════════════════════════════════════
# Inhaltsfolien
# ══════════════════════════════════════════════════════════════════════════════

def _new_slide(prs):
    """Leere Folie (Blank-Layout), alle automatisch hinzugefügten Platzhalter entfernt."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    for ph in list(slide.placeholders):
        ph.element.getparent().remove(ph.element)
    return slide


def _move_before_last(prs):
    """Verschiebt die zuletzt hinzugefügte Folie vor die letzte (Abschlussfolie)."""
    lst = prs.slides._sldIdLst
    items = list(lst)
    last = items[-1]
    lst.remove(last)
    lst.insert(len(items) - 2, last)


def _content_slide(prs, articles, page_num, first_num, logo):
    slide = _new_slide(prs)
    _header(slide, logo)
    _footer(slide, logo, page_num)

    n = len(articles)
    if n == 0:
        _move_before_last(prs)
        return

    total_h = CONTENT_BOT - CONTENT_TOP
    block_h = total_h / n

    for i, article in enumerate(articles):
        top = CONTENT_TOP + i * block_h
        _article_block(slide, article, first_num + i,
                       top, block_h,
                       MARGIN, TEXT_L_FULL, TEXT_W_FULL)
        if i < n - 1:
            _divider(slide, top + block_h - Inches(0.16))

    _move_before_last(prs)


# ══════════════════════════════════════════════════════════════════════════════
# Titelfolie (Cover)
# ══════════════════════════════════════════════════════════════════════════════

def _cover_slide(prs_slide, banner, logo,
                 month_name, year, issue_str, articles):
    """
    Baut die Titelfolie komplett neu auf:
      – Behält nur das IBM-Bannerfoto (Grafik 4) und aktualisiert sonstige Texte
      – Baut Seitenleiste (dunkelblau) + TOC (weiße Schrift) neu
      – Artikel 1 auf der rechten Seite
    """
    slide = prs_slide

    # ── Alle alten dynamischen Shapes löschen ─────────────────────────────────
    # Behalten: Grafik 4 (Banner), Grafik 82, Picture 14 (Logos/Badges im Header)
    # Löschen:  alle Shapes mit Text (außer 'IBM Z Newsletter', 'Issue No.') +
    #           Ovale + alte Textfelder im Inhaltsbereich
    KEEP = {"Grafik 4", "Grafik 82", "Picture 14", "Picture 46",
            "Rechteck 53", "Rechteck 80", "Textfeld 5"}  # Banner + Querbalken + IBM Z Newsletter

    to_del = []
    for shape in slide.shapes:
        if shape.name in KEEP:
            continue
        if shape.shape_type == 13:   # Alle Bilder behalten
            continue
        to_del.append(shape._element)

    for el in to_del:
        try:
            slide.shapes._spTree.remove(el)
        except Exception:
            pass

    # ── Seitenleiste ──────────────────────────────────────────────────────────
    sidebar_top = Inches(2.00)
    sidebar_h   = H - sidebar_top
    _rect(slide, Inches(0), sidebar_top, SIDE_W, sidebar_h, BAR_GRAY)

    # ── Issue-Nummer auf dem bestehenden Querbalken (Rechteck 80) ─────────────
    _tb(slide, Inches(0.20), Inches(1.55), Inches(2.0), Inches(0.34),
        f"Issue No. {issue_str}", 9, color=NEAR_BLACK)

    # Monat / Jahr-Box (oben links, überlappt mit Grafik 4)
    _rect(slide, Inches(0.07), Inches(0.0), Inches(1.22), Inches(1.40),
          RGBColor(0x00, 0x2D, 0x9C))  # etwas dunkleres IBM-Blau
    _tb(slide, Inches(0.10), Inches(0.12), Inches(1.10), Inches(0.60),
        month_name, 11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    _tb(slide, Inches(0.10), Inches(0.68), Inches(1.10), Inches(0.46),
        str(year), 10, color=WHITE, align=PP_ALIGN.CENTER)

    # ── TOC: THEMEN-Überschrift ────────────────────────────────────────────────
    _tb(slide, SIDE_COL_L, Inches(2.10), SIDE_W - SIDE_COL_L - Inches(0.06),
        Inches(0.44), "THEMEN", 12, bold=True, color=NEAR_BLACK)

    # ── TOC: Artikel-Liste ────────────────────────────────────────────────────
    toc_start = Inches(2.60)
    row_h     = Inches(0.65)   # Höhe pro Eintrag (Platz für 2-zeilige Titel)

    for i, art in enumerate(articles):
        y = toc_start + i * row_h

        # Zahl in IBM-Blau
        _tb(slide, SIDE_COL_L, y, Inches(0.26), Inches(0.50),
            str(i + 1), 9, bold=True, color=IBM_BLUE, align=PP_ALIGN.CENTER)

        # Vollständiger Titel, schwarze Schrift, Zeilenumbruch aktiv
        _tb(slide, SIDE_COL_L + Inches(0.30), y,
            SIDE_COL_W - Inches(0.30), Inches(0.55),
            art["title"], 7, color=NEAR_BLACK, wrap=True)

    # ── Seitenleiste: Copyright unten ─────────────────────────────────────────
    _tb(slide, SIDE_COL_L, Inches(10.95), SIDE_COL_W, Inches(0.22),
        "© 2026 IBM Corporation", 6, color=MID_GRAY)

    # Seitenzahl (Mitte unten)
    _tb(slide, Inches(3.90), Inches(11.36), Inches(0.50), Inches(0.28),
        "1", 8, color=NEAR_BLACK, align=PP_ALIGN.CENTER)

    # ── Artikel 1 auf der rechten Seite ───────────────────────────────────────
    if articles:
        art_top  = Inches(2.05)
        art_h    = Inches(8.70)
        # Wir nutzen SIDE_W + etwas Abstand als linke Kante
        _article_block(slide, articles[0], 1,
                       art_top, art_h,
                       COVER_R_L, COVER_R_L + NUM_W + Inches(0.06),
                       COVER_R_W - NUM_W - Inches(0.06))


# ══════════════════════════════════════════════════════════════════════════════
# Folie löschen (inkl. Part-Entfernung aus dem Package)
# ══════════════════════════════════════════════════════════════════════════════

def _delete_slide(prs, index):
    """Entfernt eine Folie inklusive ihres XML-Parts vollständig aus der Präsentation."""
    slides = prs.slides
    slide  = list(slides)[index]

    # Relationship aus der Präsentation entfernen
    rId = slides._sldIdLst[index].get("r:id") or slides._sldIdLst[index].get(
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    )

    # Über prs.part den richtigen rId finden
    prs_part = prs.slides._pPr if hasattr(prs.slides, "_pPr") else None
    # Einfachere Methode: direkt aus slides._sldIdLst und Teil-Beziehungen
    xml_slides = prs.slides
    lst = xml_slides._sldIdLst
    sldId_elem = list(lst)[index]

    # rId aus dem Element holen (Attribut kann unterschiedlich benannt sein)
    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    rid = sldId_elem.get(f"{{{r_ns}}}id")

    # Part aus dem Package löschen
    slide_part = slide.part
    prs.part.drop_rel(rid)

    # Element aus sldIdLst entfernen
    lst.remove(sldId_elem)


# ══════════════════════════════════════════════════════════════════════════════
# Hauptfunktion
# ══════════════════════════════════════════════════════════════════════════════

def build_newsletter(articles, month, year, issue_number,
                     output_filename=None):

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    month_name = GERMAN_MONTHS.get(month, str(month))
    fname = output_filename or f"IBM_Z_Newsletter_{month_name}_{year}.pptx"
    out   = os.path.join(OUTPUT_DIR, fname)
    shutil.copy2(TEMPLATE_FILE, out)

    prs = Presentation(out)
    banner, logo = _extract_images(prs)

    # ── Titelfolie aktualisieren ──────────────────────────────────────────────
    _cover_slide(prs.slides[0], banner, logo,
                 month_name, year, issue_number or "?", articles)

    # ── Alle Zwischenfolien (Inhalte + Events) löschen, Abschlussfolie behalten
    # Richtig: Slide-Parts komplett entfernen (nicht nur aus _sldIdLst)
    while len(prs.slides) > 2:
        _delete_slide(prs, 1)

    # ── Inhaltsfolien generieren (Artikel 2, 3, … – Artikel 1 ist auf Cover) ─
    remaining = articles[1:]
    groups    = [remaining[i:i + 2] for i in range(0, len(remaining), 2)] if remaining else []

    for page_idx, group in enumerate(groups):
        first_num = 2 + page_idx * 2
        _content_slide(prs, group, page_idx + 2, first_num, logo)

    # ── Abschlussfolie: Header-Balken + Logo hinzufügen ──────────────────────
    closing_slide = list(prs.slides)[-1]
    _header(closing_slide, logo)

    prs.save(out)
    return out
