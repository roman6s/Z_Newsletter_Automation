"""
IBM Z Newsletter PPTX Builder

Layout:
  Slide 0 : Cover (from template) – TOC left sidebar, Article 1 right
  Slides 1+: Content slides (programmatically generated, 2 articles each)
  Last     : Closing slide (from template, untouched)

Design principles:
  - No oval/circle shapes – numbers as styled text for perfect alignment
  - Two-column article layout: narrow number column | wide content column
  - Consistent IBM Blue (#0043CE) for numbers and accents
  - Cover TOC uses white text on the template's dark blue sidebar
"""

import copy
import io
import os
import re
import shutil
from typing import List, Optional, Tuple

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt, Emu

from config import OUTPUT_DIR, TEMPLATE_FILE

# ── Colours ───────────────────────────────────────────────────────────────────
IBM_BLUE   = RGBColor(0x00, 0x43, 0xCE)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
DARK_GRAY  = RGBColor(0x16, 0x16, 0x16)
MID_GRAY   = RGBColor(0x6F, 0x6F, 0x6F)
LIGHT_GRAY = RGBColor(0xD8, 0xD8, 0xD8)

GERMAN_MONTHS = {
    1: "Januar", 2: "Februar", 3: "März", 4: "April",
    5: "Mai",    6: "Juni",    7: "Juli",  8: "August",
    9: "September", 10: "Oktober", 11: "November", 12: "Dezember",
}

FONT = "IBM Plex Sans"

# ── Slide geometry ────────────────────────────────────────────────────────────
HEADER_H    = Inches(0.66)
FOOTER_TOP  = Inches(11.13)
CONTENT_TOP = Inches(0.82)
CONTENT_BOT = Inches(10.95)

# Content column positions (content slides)
NUM_L   = Inches(0.28)   # number column left edge
NUM_W   = Inches(0.38)   # number column width
TEXT_L  = Inches(0.70)   # text column left edge
TEXT_W  = Inches(7.35)   # text column width

ARTICLES_PER_SLIDE = 2


# ── Slide list helpers ────────────────────────────────────────────────────────

def _delete_slide(prs: Presentation, index: int):
    lst = prs.slides._sldIdLst
    items = list(lst)
    if 0 <= index < len(items):
        lst.remove(items[index])


def _move_slide(prs: Presentation, from_idx: int, to_idx: int):
    lst = prs.slides._sldIdLst
    items = list(lst)
    elem = items[from_idx]
    lst.remove(elem)
    lst.insert(to_idx, elem)


# ── Logo extraction ───────────────────────────────────────────────────────────

def _extract_logo(prs: Presentation) -> Tuple[Optional[bytes], Optional[str]]:
    """Return (bytes, content_type) of the IBM header logo from the template."""
    for slide in list(prs.slides)[1:]:
        for shape in slide.shapes:
            if shape.shape_type == 13 and "46" in shape.name:
                if shape.top < Inches(1.0):
                    return shape.image.blob, shape.image.content_type
    for slide in list(prs.slides)[1:]:
        for shape in slide.shapes:
            if shape.shape_type == 13 and shape.top < Inches(1.0):
                return shape.image.blob, shape.image.content_type
    return None, None


# ── XML / text helpers ────────────────────────────────────────────────────────

def _para_set_text(p_elem, text: str):
    """Overwrite text of the first run in an <a:p>, clear the rest."""
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


def _shape_set_text(shape, text: str):
    """Replace all text in a shape, keeping first-run formatting as base."""
    if not shape.has_text_frame:
        return
    txBody = shape.text_frame._txBody
    runs = txBody.findall(".//" + qn("a:r"))
    tmpl = copy.deepcopy(runs[0]) if runs else None
    for p in txBody.findall(qn("a:p")):
        txBody.remove(p)
    for line in (text.split("\n") if text else [""]):
        p = etree.SubElement(txBody, qn("a:p"))
        if tmpl is not None:
            r = copy.deepcopy(tmpl)
            t = r.find(qn("a:t"))
            if t is None:
                t = etree.SubElement(r, qn("a:t"))
            t.text = line
            p.append(r)
        else:
            r = etree.SubElement(p, qn("a:r"))
            t = etree.SubElement(r, qn("a:t"))
            t.text = line


# ── Primitive builders ────────────────────────────────────────────────────────

def _rect(slide, left, top, width, height, color: RGBColor):
    s = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE,
                               left, top, width, height)
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    return s


def _textbox(slide, left, top, width, height,
             text: str, size: int,
             bold=False, color: RGBColor = DARK_GRAY,
             align=PP_ALIGN.LEFT, wrap=True,
             italic=False, underline=False) -> object:
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.underline = underline
    r.font.color.rgb = color
    r.font.name = FONT
    return tb


def _hyperlink_textbox(slide, left, top, width, height,
                       display: str, url: str):
    """Text box with a clickable hyperlink."""
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = display
    r.font.size = Pt(9)
    r.font.color.rgb = IBM_BLUE
    r.font.underline = True
    r.font.name = FONT
    try:
        rId = slide.part.relate_to(
            url,
            "http://schemas.openxmlformats.org/officeDocument/2006/"
            "relationships/hyperlink",
            is_external=True,
        )
        rPr = r._r.get_or_add_rPr()
        hl = etree.SubElement(rPr, qn("a:hlinkClick"))
        hl.set(qn("r:id"), rId)
    except Exception:
        pass
    return tb


def _divider(slide, y: Emu):
    """Thin full-width separator line."""
    ln = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(0.28), y, Inches(7.73), Inches(0.02),
    )
    ln.fill.solid()
    ln.fill.fore_color.rgb = LIGHT_GRAY
    ln.line.fill.background()


# ── Header / footer ───────────────────────────────────────────────────────────

def _add_header(slide, logo_bytes, logo_ct):
    _rect(slide, Inches(-0.02), Inches(0), Inches(8.31), HEADER_H, IBM_BLUE)
    if logo_bytes:
        slide.shapes.add_picture(
            io.BytesIO(logo_bytes),
            Inches(7.16), Inches(0.17), Inches(0.82), Inches(0.33),
        )


def _add_footer(slide, logo_bytes, logo_ct, page_num: int):
    _rect(slide, Inches(-0.02), FOOTER_TOP, Inches(8.31), Inches(0.56), IBM_BLUE)
    _textbox(slide, Inches(0.10), Inches(11.44), Inches(2.0), Inches(0.22),
             "© 2026 IBM Corporation", size=6, color=WHITE)
    _textbox(slide, Inches(3.90), Inches(11.36), Inches(0.50), Inches(0.28),
             str(page_num), size=8, color=WHITE, align=PP_ALIGN.CENTER)
    if logo_bytes:
        slide.shapes.add_picture(
            io.BytesIO(logo_bytes),
            Inches(7.16), Inches(11.31), Inches(0.82), Inches(0.33),
        )


# ── Article block (content slides) ───────────────────────────────────────────

def _article_block(slide, article: dict, number: int,
                   top: Emu, block_height: Emu) -> Emu:
    """
    Draw one article in a two-column layout:
      Left  (0.28"–0.68"): article number (IBM Blue, bold, 16pt)
      Right (0.70"–8.05"): title · summary · link · author/date

    Returns the bottom Y of the block.
    """
    padding_top = Inches(0.12)
    y = top + padding_top

    # ── Number (left column) ──────────────────────────────────────────────────
    _textbox(slide, NUM_L, y, NUM_W, Inches(0.50),
             str(number), size=18, bold=True,
             color=IBM_BLUE, align=PP_ALIGN.CENTER)

    # ── Title ─────────────────────────────────────────────────────────────────
    title_h = Inches(0.52)
    _textbox(slide, TEXT_L, y, TEXT_W, title_h,
             article["title"], size=11, bold=True, color=DARK_GRAY)

    # ── Summary ───────────────────────────────────────────────────────────────
    summary_top = y + title_h + Inches(0.06)
    # Give summary the lion's share of the remaining block height
    summary_h = block_height - padding_top - title_h - Inches(0.06) \
                - Inches(0.30) - Inches(0.28) - Inches(0.12)
    summary_h = max(summary_h, Inches(1.20))
    _textbox(slide, TEXT_L, summary_top, TEXT_W, summary_h,
             article["summary"], size=9, color=DARK_GRAY, wrap=True)

    # ── Link ──────────────────────────────────────────────────────────────────
    link_top = summary_top + summary_h + Inches(0.06)
    _hyperlink_textbox(slide, TEXT_L, link_top, TEXT_W, Inches(0.25),
                       "→ Mehr Informationen erhalten Sie hier",
                       article.get("url", ""))

    # ── Author / Date ─────────────────────────────────────────────────────────
    pub = article.get("published")
    date_str = (f"{pub.day}. {GERMAN_MONTHS.get(pub.month, '')} {pub.year}"
                if pub else "")
    author_top = link_top + Inches(0.27)
    _textbox(slide, TEXT_L, author_top, TEXT_W, Inches(0.24),
             f"{article.get('author', '')}  ·  {date_str}",
             size=8, color=MID_GRAY, italic=True)

    return top + block_height


# ── Content slide factory ─────────────────────────────────────────────────────

def _new_content_slide(prs: Presentation, articles: List[dict],
                       page_num: int, first_article_num: int,
                       logo_bytes, logo_ct):
    """
    Generate one content slide and insert it before the closing slide.
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

    _add_header(slide, logo_bytes, logo_ct)
    _add_footer(slide, logo_bytes, logo_ct, page_num)

    if not articles:
        _move_slide(prs, len(prs.slides) - 1, len(prs.slides) - 2)
        return slide

    n = len(articles)
    total_h = CONTENT_BOT - CONTENT_TOP
    block_h = total_h / n

    for i, article in enumerate(articles):
        top = CONTENT_TOP + i * block_h
        _article_block(slide, article, first_article_num + i, top, block_h)
        if i < n - 1:
            _divider(slide, top + block_h - Inches(0.18))

    _move_slide(prs, len(prs.slides) - 1, len(prs.slides) - 2)
    return slide


# ── Cover slide ───────────────────────────────────────────────────────────────

def _update_cover(slide, month_name: str, year: int,
                  issue_str: str, articles: List[dict]):
    """Update the cover slide in-place."""
    month_re = re.compile(
        r"(Januar|Februar|März|April|Mai|Juni|Juli|August|"
        r"September|Oktober|November|Dezember)"
    )

    # Update month / year / issue number using existing shapes
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        txt = shape.text_frame.text
        if month_re.search(txt) and re.search(r"\d{4}", txt):
            tb = shape.text_frame._txBody
            paras = tb.findall(qn("a:p"))
            nonempty = [p for p in paras if
                        "".join(t.text or "" for t in
                                p.findall(".//" + qn("a:t"))).strip()]
            if len(nonempty) >= 2:
                _para_set_text(nonempty[0], month_name)
                _para_set_text(nonempty[-1], str(year))
            elif nonempty:
                _para_set_text(nonempty[0], f"{month_name} {year}")
        elif "Issue No." in txt:
            _shape_set_text(shape, f"Issue No. {issue_str}")

    # Remove old article content and old TOC items from right + sidebar
    to_del = []
    for shape in slide.shapes:
        # Right-side article content (old template text)
        if (hasattr(shape, "left") and shape.left > Inches(2.5)
                and hasattr(shape, "top") and shape.top > Inches(2.0)
                and shape.shape_type != 13):
            to_del.append(shape._element)
        # Old TOC items in sidebar (ovals + small text boxes below y=2")
        if (hasattr(shape, "left") and shape.left < Inches(3.0)
                and hasattr(shape, "top") and shape.top > Inches(2.1)
                and shape.shape_type != 13):
            to_del.append(shape._element)

    seen: set = set()
    for el in to_del:
        eid = id(el)
        if eid not in seen:
            seen.add(eid)
            try:
                slide.shapes._spTree.remove(el)
            except Exception:
                pass

    # Rebuild TOC
    _build_toc(slide, articles)

    # Article 1 on the right
    if articles:
        _cover_article(slide, articles[0])


def _build_toc(slide, articles: List[dict]):
    """
    Add a clean numbered list for the TOC on the cover sidebar.
    Text is WHITE to be visible on the dark blue sidebar background.
    Layout: "1  Short title" per line.
    """
    y      = Inches(2.35)
    row_h  = Inches(0.38)
    num_l  = Inches(0.14)
    num_w  = Inches(0.28)
    txt_l  = Inches(0.46)
    txt_w  = Inches(2.32)

    for i, article in enumerate(articles):
        top = y + i * row_h

        # Number
        _textbox(slide, num_l, top, num_w, Inches(0.34),
                 str(i + 1), size=9, bold=True,
                 color=WHITE, align=PP_ALIGN.CENTER)

        # Short title
        short = article["title"]
        if len(short) > 30:
            short = short[:27].rsplit(" ", 1)[0] + "..."
        _textbox(slide, txt_l, top, txt_w, Inches(0.34),
                 short, size=8, color=WHITE, wrap=False)


def _cover_article(slide, article: dict):
    """
    Place Article 1 in the right content area of the cover slide.
    Layout mirrors the content slides: large number left, text right.
    """
    top   = Inches(2.25)
    left  = Inches(3.00)
    width = Inches(5.05)
    num_w = Inches(0.45)
    txt_l = left + num_w + Inches(0.05)
    txt_w = width - num_w - Inches(0.05)

    # Number
    _textbox(slide, left, top, num_w, Inches(0.55),
             "1", size=20, bold=True,
             color=IBM_BLUE, align=PP_ALIGN.CENTER)

    # Title
    _textbox(slide, txt_l, top, txt_w, Inches(0.55),
             article["title"], size=11, bold=True, color=DARK_GRAY)

    # Summary
    sum_top = top + Inches(0.60)
    _textbox(slide, left, sum_top, width, Inches(5.70),
             article["summary"], size=9, color=DARK_GRAY, wrap=True)

    # Link
    link_top = sum_top + Inches(5.80)
    _hyperlink_textbox(slide, left, link_top, width, Inches(0.28),
                       "→ Mehr Informationen erhalten Sie hier",
                       article.get("url", ""))

    # Author / date
    pub = article.get("published")
    date_str = (f"{pub.day}. {GERMAN_MONTHS.get(pub.month, '')} {pub.year}"
                if pub else "")
    _textbox(slide, left, link_top + Inches(0.30), width, Inches(0.25),
             f"{article.get('author', '')}  ·  {date_str}",
             size=8, color=MID_GRAY, italic=True)


# ── Main ──────────────────────────────────────────────────────────────────────

def build_newsletter(articles: List[dict],
                     month: int, year: int, issue_number: str,
                     output_filename: Optional[str] = None) -> str:
    """Build and save the newsletter. Returns the output file path."""
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    month_name = GERMAN_MONTHS.get(month, str(month))
    fname = output_filename or f"IBM_Z_Newsletter_{month_name}_{year}.pptx"
    out   = os.path.join(OUTPUT_DIR, fname)
    shutil.copy2(TEMPLATE_FILE, out)

    prs = Presentation(out)
    logo_bytes, logo_ct = _extract_logo(prs)

    # ── Cover ─────────────────────────────────────────────────────────────────
    _update_cover(prs.slides[0], month_name, year, issue_number, articles)

    # ── Strip all slides except cover (0) and closing (last) ──────────────────
    for i in range(len(prs.slides) - 2, 0, -1):
        _delete_slide(prs, i)
    # prs.slides == [cover, closing]

    # ── Content slides (articles[1:], article[0] is on cover) ─────────────────
    remaining = articles[1:]
    groups = [remaining[i:i + ARTICLES_PER_SLIDE]
              for i in range(0, max(len(remaining), 1), ARTICLES_PER_SLIDE)]

    for page_idx, group in enumerate(groups):
        first_num = 1 + 1 + page_idx * ARTICLES_PER_SLIDE  # article 1 on cover
        _new_content_slide(prs, group, page_idx + 2, first_num,
                           logo_bytes, logo_ct)

    prs.save(out)
    return out
