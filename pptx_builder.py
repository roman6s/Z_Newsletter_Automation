"""
Builds the IBM Z Newsletter PPTX from the template.

Template structure (Oktober Newsletter.pptx):
  Slide 0: Cover - title, date, issue number, TOC sidebar
  Slides 1-4: Content slides - articles with oval numbers, titles, body text
  Slide 5: Events / Upcoming Webinars
  Slide 6: Closing / Additional Info (kept as-is)
"""

import copy
import os
import re
import shutil
from datetime import date
from typing import List, Optional

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt
from lxml import etree

from config import OUTPUT_DIR, TEMPLATE_FILE

GERMAN_MONTHS = {
    1: "Januar", 2: "Februar", 3: "März", 4: "April",
    5: "Mai", 6: "Juni", 7: "Juli", 8: "August",
    9: "September", 10: "Oktober", 11: "November", 12: "Dezember",
}

# Template content capacity: 4 content slides × 2 articles
ARTICLES_PER_SLIDE = 2
NUM_CONTENT_SLIDES = 4


# ---------------------------------------------------------------------------
# Low-level XML text helpers
# ---------------------------------------------------------------------------

def _get_all_runs(shape):
    """Return all <a:r> elements in a shape's text frame."""
    if not shape.has_text_frame:
        return []
    txBody = shape.text_frame._txBody
    return txBody.findall(".//" + qn("a:r"))


def _set_shape_text(shape, new_text: str):
    """
    Replace ALL text in a shape with new_text, preserving the formatting
    of the very first run found (font, size, bold, color etc.).
    Newlines in new_text become separate paragraphs.
    """
    if not shape.has_text_frame:
        return

    txBody = shape.text_frame._txBody

    # Grab a copy of the first run as formatting template
    all_runs = _get_all_runs(shape)
    run_template = copy.deepcopy(all_runs[0]) if all_runs else None

    # Remove all <a:p> elements
    for p in txBody.findall(qn("a:p")):
        txBody.remove(p)

    lines = new_text.split("\n") if new_text else [""]

    for line in lines:
        p_elem = etree.SubElement(txBody, qn("a:p"))
        if run_template is not None:
            r_elem = copy.deepcopy(run_template)
            t_elem = r_elem.find(qn("a:t"))
            if t_elem is None:
                t_elem = etree.SubElement(r_elem, qn("a:t"))
            t_elem.text = line
            p_elem.append(r_elem)
        else:
            r_elem = etree.SubElement(p_elem, qn("a:r"))
            t_elem = etree.SubElement(r_elem, qn("a:t"))
            t_elem.text = line


def _set_shape_text_in_first_run(shape, new_text: str):
    """Replace only the first run text in the first paragraph (light touch)."""
    if not shape.has_text_frame:
        return
    runs = _get_all_runs(shape)
    if runs:
        t = runs[0].find(qn("a:t"))
        if t is not None:
            t.text = new_text
        # Clear remaining runs
        for r in runs[1:]:
            t = r.find(qn("a:t"))
            if t is not None:
                t.text = ""


# ---------------------------------------------------------------------------
# Slide helpers
# ---------------------------------------------------------------------------

def _find_shapes_with_text(slide, substring: str):
    return [s for s in slide.shapes
            if s.has_text_frame and substring in s.text_frame.text]


def _content_ovals(slide):
    """
    Return oval shapes that mark article sections.
    These are small round shapes in the left margin below the header.
    """
    return sorted(
        [s for s in slide.shapes
         if "Oval" in s.name
         and hasattr(s, "top") and s.top > Inches(0.65)
         and hasattr(s, "left") and s.left < Inches(0.6)],
        key=lambda s: s.top,
    )


def _find_title_for_oval(slide, oval):
    """
    Find the title text box for a given oval anchor.
    Title is: text box at roughly same vertical level as oval, to its right.
    """
    candidates = []
    for shape in slide.shapes:
        if not shape.has_text_frame or "Oval" in shape.name:
            continue
        vert_dist = abs(shape.top - oval.top)
        horiz_ok = shape.left > oval.left + oval.width - Inches(0.1)
        if vert_dist < Inches(0.25) and horiz_ok:
            candidates.append((vert_dist, id(shape), shape))
    candidates.sort()
    return candidates[0][2] if candidates else None


def _find_body_for_oval(slide, oval, title_shape, next_oval_top=None):
    """
    Find the body text box for a given oval.
    Body is: tall text box below the oval/title, within the article's vertical zone.
    """
    ref_top = oval.top + oval.height
    if title_shape:
        ref_top = max(ref_top, title_shape.top + title_shape.height - Inches(0.05))

    zone_bottom = next_oval_top if next_oval_top else Inches(11.0)

    candidates = []
    for shape in slide.shapes:
        if not shape.has_text_frame or "Oval" in shape.name:
            continue
        if title_shape and shape is title_shape:
            continue
        in_zone = shape.top >= ref_top - Inches(0.1) and shape.top < zone_bottom
        is_tall = shape.height > Inches(0.35)
        if in_zone and is_tall:
            candidates.append((shape.top, id(shape), shape))
    candidates.sort()
    return candidates[0][2] if candidates else None


# ---------------------------------------------------------------------------
# Cover slide update
# ---------------------------------------------------------------------------

def _update_cover_slide(slide, month_name: str, year: int,
                         issue_number: str, articles: List[dict]):
    """Update the cover slide: month/year, issue number, TOC, clear old content."""

    month_pattern = re.compile(
        r"(Januar|Februar|März|April|Mai|Juni|Juli|August|"
        r"September|Oktober|November|Dezember)"
    )

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        full_text = shape.text_frame.text

        # Month + Year box (shape "Rechteck 66" – tall narrow box with month\n\nyear)
        if month_pattern.search(full_text) and re.search(r"\d{4}", full_text):
            txBody = shape.text_frame._txBody
            paras = txBody.findall(qn("a:p"))
            # Set first non-empty para to month, last non-empty to year
            non_empty = [p for p in paras if
                         "".join(t.text or "" for t in p.findall(".//" + qn("a:t"))).strip()]
            if len(non_empty) >= 2:
                _para_set_text(non_empty[0], month_name)
                _para_set_text(non_empty[-1], str(year))
            elif len(non_empty) == 1:
                _para_set_text(non_empty[0], f"{month_name} {year}")
            continue

        # Issue number box
        if "Issue No." in full_text:
            _set_shape_text(shape, f"Issue No. {issue_number}")
            continue

        # Clear old main article content (large boxes in the right content area)
        # These are wide boxes (> 4") in the content area (top > 2", left > 2.5")
        if (hasattr(shape, "top") and shape.top > Inches(2.0)
                and hasattr(shape, "left") and shape.left > Inches(2.5)
                and hasattr(shape, "width") and shape.width > Inches(4.0)):
            _set_shape_text(shape, "")
            continue

    # TOC: update short article titles
    _update_toc(slide, articles)


def _para_set_text(p_elem, text: str):
    """Replace the text of all runs in a single <a:p> element."""
    runs = p_elem.findall(qn("a:r"))
    if not runs:
        r = etree.SubElement(p_elem, qn("a:r"))
        t = etree.SubElement(r, qn("a:t"))
        t.text = text
        return
    # Set first run, clear others
    t = runs[0].find(qn("a:t"))
    if t is None:
        t = etree.SubElement(runs[0], qn("a:t"))
    t.text = text
    for r in runs[1:]:
        t2 = r.find(qn("a:t"))
        if t2 is not None:
            t2.text = ""


def _update_toc(slide, articles: List[dict]):
    """
    Update the table of contents entries in the left sidebar.
    These are small text boxes (width < 2.5") in the left column.
    """
    toc_shapes = sorted(
        [s for s in slide.shapes
         if s.has_text_frame
         and s.width < Inches(2.6)
         and hasattr(s, "left") and s.left < Inches(3.0)
         and hasattr(s, "top") and s.top > Inches(2.0)
         and s.name.startswith("TextBox")],
        key=lambda s: s.top,
    )

    for i, shape in enumerate(toc_shapes):
        if i < len(articles):
            title = articles[i]["title"]
            short = title[:26].rsplit(" ", 1)[0] + "..." if len(title) > 28 else title
            _set_shape_text(shape, short)
        else:
            _set_shape_text(shape, "")

    # Update oval bullet numbers (TOC numbering)
    toc_ovals = sorted(
        [s for s in slide.shapes
         if "Oval" in s.name
         and hasattr(s, "top") and s.top > Inches(2.0)
         and hasattr(s, "left") and s.left < Inches(0.6)],
        key=lambda s: s.top,
    )
    for i, oval in enumerate(toc_ovals):
        if i < len(articles):
            _set_shape_text(oval, str(i + 1))
        else:
            _set_shape_text(oval, "")


# ---------------------------------------------------------------------------
# Content slide filling
# ---------------------------------------------------------------------------

def _fill_content_slide(slide, article_group: List[dict],
                         page_number: int, global_offset: int):
    """
    Fill a content slide with 0-2 articles.

    article_group: list of article dicts for this slide
    page_number: displayed page number
    global_offset: number of articles already shown on previous slides
    """
    # Update page number in footer area
    for shape in slide.shapes:
        if (shape.has_text_frame
                and shape.text_frame.text.strip().lstrip("-").isdigit()
                and hasattr(shape, "top") and shape.top > Inches(11.0)
                and hasattr(shape, "width") and shape.width < Inches(0.5)):
            _set_shape_text(shape, str(page_number))

    # Identify article section anchors (ovals) and their zones
    ovals = _content_ovals(slide)

    for slot_idx, oval in enumerate(ovals):
        next_oval_top = ovals[slot_idx + 1].top if slot_idx + 1 < len(ovals) else None
        title_shape = _find_title_for_oval(slide, oval)
        body_shape = _find_body_for_oval(slide, oval, title_shape, next_oval_top)

        if slot_idx < len(article_group):
            article = article_group[slot_idx]
            article_num = global_offset + slot_idx + 1

            # Oval: article number
            _set_shape_text(oval, str(article_num))

            # Title
            if title_shape:
                _set_shape_text(title_shape, article["title"])

            # Body: summary + author + date
            if body_shape:
                summary = article["summary"]
                pub = article.get("published")
                if pub:
                    date_str = (f"{pub.day}. {GERMAN_MONTHS.get(pub.month, '')} "
                                f"{pub.year}")
                    summary += f"\n\nAutor: {article['author']} | {date_str}"
                _set_shape_text(body_shape, summary)
        else:
            # Clear unused slot
            _set_shape_text(oval, "")
            if title_shape:
                _set_shape_text(title_shape, "")
            if body_shape:
                _set_shape_text(body_shape, "")


# ---------------------------------------------------------------------------
# Main builder
# ---------------------------------------------------------------------------

def build_newsletter(
    articles: List[dict],
    month: int,
    year: int,
    issue_number: str,
    output_filename: Optional[str] = None,
) -> str:
    """
    Build the newsletter PPTX from the template.

    articles: list of dicts with keys: title, author, url, published, summary
    month: newsletter month (int)
    year: newsletter year (int)
    issue_number: e.g. "14"
    output_filename: optional output file name

    Returns: path to the generated PPTX file.
    """
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    month_name = GERMAN_MONTHS.get(month, str(month))
    if output_filename is None:
        output_filename = f"IBM_Z_Newsletter_{month_name}_{year}.pptx"
    output_path = os.path.join(OUTPUT_DIR, output_filename)
    shutil.copy2(TEMPLATE_FILE, output_path)

    prs = Presentation(output_path)

    # Cap articles to template capacity
    max_articles = NUM_CONTENT_SLIDES * ARTICLES_PER_SLIDE
    if len(articles) > max_articles:
        print(f"  Hinweis: {len(articles)} Artikel gefunden, zeige die ersten {max_articles}.")
        articles = articles[:max_articles]

    # --- Slide 0: Cover ---
    _update_cover_slide(prs.slides[0], month_name, year, issue_number, articles)

    # --- Slides 1-4: Content ---
    for slide_idx in range(NUM_CONTENT_SLIDES):
        pptx_slide_idx = slide_idx + 1  # 0=cover, so content starts at 1
        if pptx_slide_idx >= len(prs.slides):
            break
        start = slide_idx * ARTICLES_PER_SLIDE
        group = articles[start:start + ARTICLES_PER_SLIDE]
        _fill_content_slide(
            prs.slides[pptx_slide_idx],
            group,
            page_number=pptx_slide_idx + 1,
            global_offset=start,
        )

    # --- Last slides (events, closing): leave untouched ---

    prs.save(output_path)
    return output_path
