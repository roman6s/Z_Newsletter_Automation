"""
IBM Z Newsletter PPTX Builder

Layout:
  Slide 0: Cover (from template) – TOC left, Article 1 right
  Slides 1..N-1: Content slides (programmatically generated, 2 articles each)
  Last slide: Closing (from template, untouched)

Content slide structure:
  - IBM Blue header bar + logo
  - Article sections (oval number, title, summary, link, author/date)
  - Thin divider between articles
  - IBM Blue footer bar + copyright + page number
"""

import copy
import io
import os
import re
import shutil
from datetime import date
from typing import List, Optional, Tuple

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt, Emu

from config import OUTPUT_DIR, TEMPLATE_FILE

# ── Constants ─────────────────────────────────────────────────────────────────

GERMAN_MONTHS = {
    1: "Januar", 2: "Februar", 3: "März", 4: "April",
    5: "Mai", 6: "Juni", 7: "Juli", 8: "August",
    9: "September", 10: "Oktober", 11: "November", 12: "Dezember",
}

IBM_BLUE    = RGBColor(0x00, 0x43, 0xCE)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
DARK_GRAY   = RGBColor(0x16, 0x16, 0x16)
MID_GRAY    = RGBColor(0x6F, 0x6F, 0x6F)
LIGHT_GRAY  = RGBColor(0xCC, 0xCC, 0xCC)
LINK_COLOR  = IBM_BLUE

# Slide dimensions (A4 portrait)
SLIDE_W = Inches(8.27)
SLIDE_H = Inches(11.69)

# Layout zones
HEADER_H      = Inches(0.66)
FOOTER_TOP    = Inches(11.13)
FOOTER_H      = Inches(0.56)
CONTENT_TOP   = Inches(0.82)
CONTENT_BOT   = Inches(10.95)
MARGIN_L      = Inches(0.28)
MARGIN_R      = Inches(8.00)
CONTENT_W     = MARGIN_R - MARGIN_L

ARTICLES_PER_SLIDE = 2
FONT_FAMILY = "IBM Plex Sans"


# ── Slide management ──────────────────────────────────────────────────────────

def _delete_slide(prs: Presentation, index: int):
    """Remove a slide by index from the presentation."""
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    if 0 <= index < len(slides):
        xml_slides.remove(slides[index])


def _move_slide(prs: Presentation, from_idx: int, to_idx: int):
    """Move a slide from one index to another."""
    xml = prs.slides._sldIdLst
    slides = list(xml)
    elem = slides[from_idx]
    xml.remove(elem)
    xml.insert(to_idx, elem)


# ── Image extraction ──────────────────────────────────────────────────────────

def _extract_logo(prs: Presentation) -> Tuple[bytes, str]:
    """
    Extract the IBM logo image bytes from the template.
    Returns (image_bytes, content_type).
    """
    # Find in first content slide – "Picture 46" in header area
    for slide in list(prs.slides)[1:]:
        for shape in slide.shapes:
            if shape.shape_type == 13 and "46" in shape.name:
                if shape.top < Inches(1.0):
                    return shape.image.blob, shape.image.content_type
    # Fallback: any small picture in header
    for slide in list(prs.slides)[1:]:
        for shape in slide.shapes:
            if shape.shape_type == 13 and shape.top < Inches(1.0):
                return shape.image.blob, shape.image.content_type
    return None, None


# ── Low-level text helpers ────────────────────────────────────────────────────

def _set_para_text(p_elem, text: str):
    """Replace all text in an <a:p> element, preserving first run formatting."""
    runs = p_elem.findall(qn("a:r"))
    if not runs:
        r = etree.SubElement(p_elem, qn("a:r"))
        t = etree.SubElement(r, qn("a:t"))
        t.text = text
        return
    t = runs[0].find(qn("a:t"))
    if t is None:
        t = etree.SubElement(runs[0], qn("a:t"))
    t.text = text
    for r in runs[1:]:
        t2 = r.find(qn("a:t"))
        if t2 is not None:
            t2.text = ""


def _replace_shape_text(shape, new_text: str):
    """
    Replace all text in a shape, keeping the first run's formatting as template.
    Newlines become separate paragraphs.
    """
    if not shape.has_text_frame:
        return
    txBody = shape.text_frame._txBody
    all_runs = txBody.findall(".//" + qn("a:r"))
    run_template = copy.deepcopy(all_runs[0]) if all_runs else None

    for p in txBody.findall(qn("a:p")):
        txBody.remove(p)

    for line in (new_text.split("\n") if new_text else [""]):
        p = etree.SubElement(txBody, qn("a:p"))
        if run_template is not None:
            r = copy.deepcopy(run_template)
            t = r.find(qn("a:t"))
            if t is None:
                t = etree.SubElement(r, qn("a:t"))
            t.text = line
            p.append(r)
        else:
            r = etree.SubElement(p, qn("a:r"))
            t = etree.SubElement(r, qn("a:t"))
            t.text = line


# ── Content slide builder ─────────────────────────────────────────────────────

def _add_rect(slide, left, top, width, height, color: RGBColor):
    """Add a solid-color rectangle with no border."""
    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE,
                                   left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def _add_textbox(slide, left, top, width, height,
                 text: str, font_size: int, bold=False,
                 color: RGBColor = DARK_GRAY,
                 align=PP_ALIGN.LEFT,
                 font_name: str = FONT_FAMILY,
                 word_wrap=True) -> object:
    """Add a text box with a single paragraph."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = word_wrap
    tf.auto_size = None

    para = tf.paragraphs[0]
    para.alignment = align
    run = para.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font_name
    return txBox


def _add_hyperlink_textbox(slide, left, top, width, height,
                            display_text: str, url: str) -> object:
    """Add a text box with a clickable hyperlink."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = False

    para = tf.paragraphs[0]
    run = para.add_run()
    run.text = display_text
    run.font.size = Pt(9)
    run.font.color.rgb = LINK_COLOR
    run.font.name = FONT_FAMILY
    run.font.underline = True

    # Add hyperlink via XML
    try:
        rId = slide.part.relate_to(
            url,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            is_external=True,
        )
        rPr = run._r.get_or_add_rPr()
        hl = etree.SubElement(rPr, qn("a:hlinkClick"))
        hl.set(qn("r:id"), rId)
    except Exception:
        pass  # Hyperlink fails gracefully

    return txBox


def _add_divider(slide, y: Emu):
    """Add a thin horizontal divider line."""
    from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
    line = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE,
                                  MARGIN_L, y, CONTENT_W, Inches(0.02))
    line.fill.solid()
    line.fill.fore_color.rgb = LIGHT_GRAY
    line.line.fill.background()


def _add_oval_bullet(slide, left, top, number: int) -> object:
    """Add a IBM-blue oval with a white number."""
    size = Inches(0.30)
    oval = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL,
                                  left, top, size, size)
    oval.fill.solid()
    oval.fill.fore_color.rgb = IBM_BLUE
    oval.line.fill.background()

    tf = oval.text_frame
    tf.margin_top = Inches(0.01)
    tf.margin_bottom = 0
    tf.margin_left = 0
    tf.margin_right = 0
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.CENTER
    run = para.add_run()
    run.text = str(number)
    run.font.color.rgb = WHITE
    run.font.size = Pt(7)
    run.font.bold = True
    run.font.name = FONT_FAMILY
    return oval


def _build_header(slide, logo_bytes, logo_ct):
    """Add IBM Blue header bar + logo to a slide."""
    _add_rect(slide, Inches(-0.02), Inches(0), Inches(8.31), HEADER_H, IBM_BLUE)
    if logo_bytes:
        slide.shapes.add_picture(
            io.BytesIO(logo_bytes),
            Inches(7.16), Inches(0.16), Inches(0.82), Inches(0.33),
        )


def _build_footer(slide, logo_bytes, logo_ct, page_number: int):
    """Add IBM Blue footer bar + copyright + page number + logo."""
    _add_rect(slide, Inches(-0.02), FOOTER_TOP, Inches(8.31), FOOTER_H, IBM_BLUE)

    # Copyright
    _add_textbox(slide,
                 Inches(0.07), Inches(11.44), Inches(2.0), Inches(0.22),
                 "© 2026 IBM Corporation", font_size=6,
                 color=WHITE, font_name=FONT_FAMILY)

    # Page number
    _add_textbox(slide,
                 Inches(3.9), Inches(11.35), Inches(0.5), Inches(0.30),
                 str(page_number), font_size=8,
                 color=WHITE, align=PP_ALIGN.CENTER, font_name=FONT_FAMILY)

    # Logo
    if logo_bytes:
        slide.shapes.add_picture(
            io.BytesIO(logo_bytes),
            Inches(7.16), Inches(11.31), Inches(0.82), Inches(0.33),
        )


def _build_article_block(slide, article: dict, article_number: int,
                          top: Emu, available_height: Emu) -> Emu:
    """
    Draw one article block (oval, title, summary, link, author/date).
    Returns the bottom Y coordinate of the block.
    """
    oval_size   = Inches(0.30)
    text_left   = MARGIN_L + Inches(0.42)
    text_width  = CONTENT_W - Inches(0.42)

    # ── Oval bullet ──
    _add_oval_bullet(slide, MARGIN_L, top + Inches(0.02), article_number)

    # ── Title ──
    title_h = Inches(0.55)
    _add_textbox(slide, text_left, top, text_width, title_h,
                 article["title"], font_size=11, bold=True,
                 color=DARK_GRAY, font_name=FONT_FAMILY)

    # ── Summary ──
    summary_top = top + title_h + Inches(0.05)
    summary_h   = Inches(2.50)
    _add_textbox(slide, text_left, summary_top, text_width, summary_h,
                 article["summary"], font_size=9,
                 color=DARK_GRAY, font_name=FONT_FAMILY, word_wrap=True)

    # ── Link ──
    link_top = summary_top + summary_h + Inches(0.06)
    _add_hyperlink_textbox(slide,
                           text_left, link_top, text_width, Inches(0.28),
                           "→ Mehr Informationen erhalten Sie hier",
                           article.get("url", ""))

    # ── Author / Date ──
    pub = article.get("published")
    date_str = ""
    if pub:
        date_str = f"{pub.day}. {GERMAN_MONTHS.get(pub.month, '')} {pub.year}"
    author_line = f"{article.get('author', '')}  ·  {date_str}"
    author_top  = link_top + Inches(0.30)
    _add_textbox(slide, text_left, author_top, text_width, Inches(0.25),
                 author_line, font_size=8,
                 color=MID_GRAY, font_name=FONT_FAMILY)

    block_bottom = author_top + Inches(0.30)
    return block_bottom


def _create_content_slide(prs: Presentation, articles: List[dict],
                           page_number: int, logo_bytes, logo_ct) -> object:
    """
    Create a fresh content slide with up to ARTICLES_PER_SLIDE articles.
    Inserts before the last slide (closing).
    """
    blank_layout = prs.slide_layouts[6]  # "Leer" / blank
    slide = prs.slides.add_slide(blank_layout)

    _build_header(slide, logo_bytes, logo_ct)
    _build_footer(slide, logo_bytes, logo_ct, page_number)

    n = len(articles)
    if n == 0:
        return slide

    # Distribute vertical space evenly
    total_h  = CONTENT_BOT - CONTENT_TOP
    block_h  = total_h / n
    divider_gap = Inches(0.25)

    for i, article in enumerate(articles):
        top = CONTENT_TOP + i * block_h + Inches(0.10)
        avail = block_h - Inches(0.20)
        _build_article_block(slide, article, page_number * ARTICLES_PER_SLIDE - (n - i - 1),
                             top, avail)

        # Divider between articles (not after the last one)
        if i < n - 1:
            divider_y = CONTENT_TOP + (i + 1) * block_h - divider_gap
            _add_divider(slide, divider_y)

    # Move newly added slide to position before closing slide
    _move_slide(prs, len(prs.slides) - 1, len(prs.slides) - 2)
    return slide


# ── Cover slide update ────────────────────────────────────────────────────────

def _update_cover_slide(slide, month_name: str, year: int,
                         issue_number: str, articles: List[dict]):
    """Update cover: month/year, issue number, TOC, Article 1 content."""

    month_re = re.compile(
        r"(Januar|Februar|März|April|Mai|Juni|Juli|August|"
        r"September|Oktober|November|Dezember)"
    )

    # ── Update month/year, issue number ──────────────────────────────────────
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text

        if month_re.search(text) and re.search(r"\d{4}", text):
            txBody = shape.text_frame._txBody
            paras = txBody.findall(qn("a:p"))
            non_empty = [p for p in paras if
                         "".join(t.text or "" for t in p.findall(".//" + qn("a:t"))).strip()]
            if len(non_empty) >= 2:
                _set_para_text(non_empty[0], month_name)
                _set_para_text(non_empty[-1], str(year))
            elif len(non_empty) == 1:
                _set_para_text(non_empty[0], f"{month_name} {year}")
            continue

        if "Issue No." in text:
            _replace_shape_text(shape, f"Issue No. {issue_number}")
            continue

    # ── Clear right content area (old article) ───────────────────────────────
    to_remove = []
    for shape in slide.shapes:
        if (hasattr(shape, "left") and shape.left > Inches(2.5)
                and hasattr(shape, "top") and shape.top > Inches(2.0)
                and shape.shape_type != 13):  # keep images
            to_remove.append(shape._element)
    for elem in to_remove:
        slide.shapes._spTree.remove(elem)

    # ── Rebuild TOC (left sidebar) ────────────────────────────────────────────
    _rebuild_toc(slide, articles)

    # ── Add Article 1 to right content area ───────────────────────────────────
    if articles:
        _add_article1_to_cover(slide, articles[0])


def _rebuild_toc(slide, articles: List[dict]):
    """Remove old TOC ovals + text boxes, add fresh ones for current articles."""
    # Remove existing TOC items (small shapes in left sidebar below header)
    to_remove = []
    for shape in slide.shapes:
        in_sidebar = (hasattr(shape, "left") and shape.left < Inches(3.0)
                      and hasattr(shape, "top") and shape.top > Inches(2.1))
        is_toc_item = (shape.has_text_frame and shape.width < Inches(2.7)) or \
                      ("Oval" in shape.name)
        if in_sidebar and is_toc_item:
            to_remove.append(shape._element)

    # Also remove the EVENTS label and sub-items
    events_labels = ["EVENTS", "und vieles mehr"]
    for shape in slide.shapes:
        if shape.has_text_frame:
            for label in events_labels:
                if label in shape.text_frame.text:
                    to_remove.append(shape._element)
                    break

    seen = set()
    for elem in to_remove:
        if id(elem) not in seen:
            seen.add(id(elem))
            try:
                slide.shapes._spTree.remove(elem)
            except Exception:
                pass

    # Add fresh TOC items
    toc_start_y = Inches(2.38)
    item_h      = Inches(0.40)
    oval_x      = Inches(0.13)
    text_x      = Inches(0.48)
    text_w      = Inches(2.30)

    for i, article in enumerate(articles):
        y = toc_start_y + i * item_h

        # Oval bullet
        _add_oval_bullet(slide, oval_x, y + Inches(0.04), i + 1)

        # Short title
        short = article["title"]
        if len(short) > 30:
            short = short[:27].rsplit(" ", 1)[0] + "..."

        txBox = slide.shapes.add_textbox(text_x, y, text_w, Inches(0.35))
        tf = txBox.text_frame
        tf.word_wrap = False
        para = tf.paragraphs[0]
        run = para.add_run()
        run.text = short
        run.font.size = Pt(8)
        run.font.color.rgb = WHITE
        run.font.name = FONT_FAMILY


def _add_article1_to_cover(slide, article: dict):
    """Add Article 1 content to the right side of the cover slide."""
    left  = Inches(3.00)
    width = Inches(5.05)
    top_start = Inches(2.25)

    # Oval + Title
    _add_oval_bullet(slide, left, top_start + Inches(0.02), 1)
    title_left = left + Inches(0.42)
    title_w    = width - Inches(0.42)
    _add_textbox(slide, title_left, top_start, title_w, Inches(0.60),
                 article["title"], font_size=11, bold=True,
                 color=DARK_GRAY, font_name=FONT_FAMILY)

    # Summary
    summary_top = top_start + Inches(0.65)
    _add_textbox(slide, left, summary_top, width, Inches(5.80),
                 article["summary"], font_size=9,
                 color=DARK_GRAY, font_name=FONT_FAMILY, word_wrap=True)

    # Link
    link_top = summary_top + Inches(5.90)
    _add_hyperlink_textbox(slide, left, link_top, width, Inches(0.28),
                           "→ Mehr Informationen erhalten Sie hier",
                           article.get("url", ""))

    # Author / Date
    pub = article.get("published")
    date_str = f"{pub.day}. {GERMAN_MONTHS.get(pub.month, '')} {pub.year}" if pub else ""
    author_top = link_top + Inches(0.30)
    _add_textbox(slide, left, author_top, width, Inches(0.25),
                 f"{article.get('author', '')}  ·  {date_str}",
                 font_size=8, color=MID_GRAY, font_name=FONT_FAMILY)


# ── Main entry point ──────────────────────────────────────────────────────────

def build_newsletter(
    articles: List[dict],
    month: int,
    year: int,
    issue_number: str,
    output_filename: Optional[str] = None,
) -> str:
    """
    Build the newsletter PPTX.
    Returns path to the generated file.
    """
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    month_name = GERMAN_MONTHS.get(month, str(month))
    if output_filename is None:
        output_filename = f"IBM_Z_Newsletter_{month_name}_{year}.pptx"
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    # Work on a copy (keep original template untouched)
    shutil.copy2(TEMPLATE_FILE, output_path)
    prs = Presentation(output_path)

    # ── Extract logo once ─────────────────────────────────────────────────────
    logo_bytes, logo_ct = _extract_logo(prs)

    # ── Update cover slide ────────────────────────────────────────────────────
    _update_cover_slide(prs.slides[0], month_name, year, issue_number, articles)

    # ── Delete all slides except cover (0) and closing (last) ─────────────────
    # Delete in reverse order to keep indices valid
    for i in range(len(prs.slides) - 2, 0, -1):
        _delete_slide(prs, i)
    # Now prs.slides = [cover, closing]

    # ── Generate content slides (articles[1:], Article 1 is on cover) ─────────
    remaining = articles[1:]
    groups = [remaining[i:i + ARTICLES_PER_SLIDE]
              for i in range(0, len(remaining), ARTICLES_PER_SLIDE)]

    # Page numbers: cover = 1, content slides = 2, 3, ...
    for page_idx, group in enumerate(groups):
        _create_content_slide(prs, group, page_idx + 2, logo_bytes, logo_ct)

    prs.save(output_path)
    return output_path
