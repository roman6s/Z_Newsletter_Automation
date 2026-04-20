"""
Scraper for IBM Z DACH Community Blog.
Fetches all blog articles within a given date range.
"""

import re
import time
from datetime import date, datetime
from dataclasses import dataclass
from typing import List, Optional, Tuple

import requests
from bs4 import BeautifulSoup

from config import BLOG_URL

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "de-DE,de;q=0.9,en;q=0.8",
}


@dataclass
class Article:
    title: str
    author: str
    url: str
    published: date
    full_text: str = ""
    image_url: str = ""
    image_bytes: bytes = b""


def _extract_entries(soup: BeautifulSoup) -> List[dict]:
    """Extract blog entry metadata from a listing page."""
    entries = []
    pattern = re.compile(r"lvSearchResults_hypTitle_(\d+)")
    for a_tag in soup.find_all("a", id=pattern):
        idx = pattern.search(a_tag["id"]).group(1)
        title = a_tag.get("title", a_tag.get_text(strip=True))
        url = a_tag.get("href", "")

        # Extract date from URL: /blogs/{author}/{yyyy}/{mm}/{dd}/...
        date_match = re.search(r"/(\d{4})/(\d{2})/(\d{2})/", url)
        pub_date = None
        if date_match:
            pub_date = date(
                int(date_match.group(1)),
                int(date_match.group(2)),
                int(date_match.group(3)),
            )

        # Extract author from ByLine div
        byline_id = f"MainCopy_ctl19_lvSearchResults_pnlPostedByCreateDate_{idx}"
        byline_div = soup.find(id=byline_id)
        author = ""
        if byline_div:
            author_a = byline_div.find("a")
            if author_a:
                author = author_a.get_text(strip=True)

        if title and url and pub_date:
            entries.append({"title": title, "author": author, "url": url, "date": pub_date})

    return entries


def _get_hidden_fields(soup: BeautifulSoup) -> dict:
    """Collect ALL hidden input fields – required for ASP.NET postback (incl. CSRF token)."""
    fields = {}
    for inp in soup.find_all("input", type="hidden"):
        name = inp.get("name", "")
        if name:
            fields[name] = inp.get("value", "")
    return fields


def _find_next_pager_target(
    soup: BeautifulSoup,
    current_page: int,
    visited: set,
) -> Tuple[Optional[str], Optional[str]]:
    """
    Parse the rendered pager HTML and return the postback (target, arg) for
    the next page to fetch.

    Strategy:
      1. Look for a link whose visible text equals str(current_page + 1).
      2. Fallback: look for a "next group" arrow link (>, >>, », …).
      3. Return (None, None) when no further navigation is available.

    Note: visited stores (target, arg) tuples – NOT just target strings –
    because all pager links share the same ASP.NET target control name but
    use different arguments to identify the page.
    """
    next_page_text = str(current_page + 1)
    arrow_texts = {">", ">>", "»", "›", "...", "Next", "Nächste"}

    arrow_candidate: Tuple[Optional[str], Optional[str]] = (None, None)

    for a in soup.find_all("a", href=True):
        href = a.get("href", "")
        m = re.search(r"__doPostBack\('([^']+)','([^']*)'\)", href)
        if not m:
            continue
        target = m.group(1)
        arg = m.group(2)
        if "Pager" not in target:
            continue
        if (target, arg) in visited:
            continue
        text = a.get_text(strip=True)
        if text == next_page_text:
            return target, arg
        if text in arrow_texts and arrow_candidate == (None, None):
            arrow_candidate = (target, arg)

    return arrow_candidate


def fetch_articles_in_range(
    start_date: date,
    end_date: date,
    verbose: bool = True,
) -> List[Article]:
    """
    Fetch all blog articles published between start_date and end_date (inclusive).
    Returns a list of Article objects with full text loaded.
    """
    session = requests.Session()
    session.headers.update(HEADERS)

    if verbose:
        print("Fetching blog listing page...")

    resp = session.get(BLOG_URL, timeout=30)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    all_entries: List[dict] = []
    visited_targets: set = set()
    stop_early = False
    current_page = 1
    max_pages = 100  # safety ceiling

    def _process_entries(entries: List[dict]) -> None:
        nonlocal stop_early
        for e in entries:
            if e["date"] > end_date:
                continue  # too new – skip, but keep paginating
            if e["date"] < start_date:
                stop_early = True
                return  # gone past start_date, no need to look further
            all_entries.append(e)

    # Process page 1
    _process_entries(_extract_entries(soup))

    # Paginate through remaining pages
    while not stop_early and current_page < max_pages:
        target, arg = _find_next_pager_target(soup, current_page, visited_targets)
        if target is None:
            break

        visited_targets.add((target, arg))
        hidden = _get_hidden_fields(soup)
        hidden["__EVENTTARGET"] = target
        hidden["__EVENTARGUMENT"] = arg or ""

        if verbose:
            print(f"  Seite {current_page + 1} laden...")

        try:
            resp = session.post(BLOG_URL, data=hidden, timeout=30)
            resp.raise_for_status()
        except Exception as e:
            print(f"  Warnung: Seite {current_page + 1} konnte nicht geladen werden: {e}")
            break

        soup = BeautifulSoup(resp.text, "html.parser")
        entries = _extract_entries(soup)

        if not entries:
            break

        _process_entries(entries)
        current_page += 1
        time.sleep(0.5)

    if verbose:
        print(f"Found {len(all_entries)} articles in date range.")

    # Deduplicate by URL
    seen: set = set()
    unique_entries: List[dict] = []
    for e in all_entries:
        if e["url"] not in seen:
            seen.add(e["url"])
            unique_entries.append(e)

    # Fetch full article text
    articles: List[Article] = []
    for i, e in enumerate(unique_entries):
        if verbose:
            print(f"  Artikel {i+1}/{len(unique_entries)}: {e['title'][:60]}...")
        full_text, image_url = _fetch_article_content(session, e["url"])
        image_bytes = b""
        if image_url:
            try:
                r = session.get(image_url, timeout=5)
                if r.ok and "image" in r.headers.get("content-type", ""):
                    image_bytes = r.content
            except Exception:
                image_bytes = b""
        articles.append(
            Article(
                title=e["title"],
                author=e["author"],
                url=e["url"],
                published=e["date"],
                full_text=full_text,
                image_url=image_url,
                image_bytes=image_bytes,
            )
        )
        time.sleep(0.3)

    return articles


def _fetch_article_content(session: requests.Session, url: str):
    """Fetch article text and first image URL. Returns (text, image_url)."""
    try:
        resp = session.get(url, timeout=30)
        resp.raise_for_status()
        html = resp.text
        soup = BeautifulSoup(html, "html.parser")

        # ── Artikeltext ───────────────────────────────────────────────────────
        body = None
        for selector in [
            {"id": re.compile(r"BlogPostBody|PostBody|ArticleBody", re.I)},
            {"class": re.compile(r"blog-post-body|post-content|entry-content", re.I)},
        ]:
            body = soup.find(attrs=selector)
            if body:
                break

        if not body:
            candidates = soup.find_all("div", class_=True)
            best = None
            best_len = 0
            for div in candidates:
                t = div.get_text(separator=" ", strip=True)
                if len(t) > best_len and len(t) < 50000:
                    best_len = len(t)
                    best = div
            body = best

        text = ""
        if body:
            for tag in body.find_all(["nav", "aside", "script", "style", "footer"]):
                tag.decompose()
            text = body.get_text(separator="\n", strip=True)
            text = re.sub(r"\n{3,}", "\n\n", text)
            text = text[:8000]

        # ── Artikelbild (IBM Community / Higher Logic) ────────────────────────
        SKIP_GUIDS = {
            "8b2c700c-5b4c-4e59-a864-e9ba84f18b1d",
            "dfd5be75-7434-44d5-beed-41626462dd64",
        }
        SKIP_NAMES = ("icon-", "loading", "avatar", "logo", "badge", "button")

        image_url = ""
        IMG_EXT = r'[^\s"\'<>]+\.(?:jpg|jpeg|png|webp|bmp)'

        uploaded = re.findall(
            r'https?://[^\s"\'<>]+/UploadedImages/' + IMG_EXT, html, re.I)
        for u in uploaded:
            if any(g in u for g in SKIP_GUIDS):
                continue
            if any(k in u.lower() for k in SKIP_NAMES):
                continue
            image_url = u
            break

        if not image_url:
            featured = re.findall(
                r'https?://[^\s"\'<>]+/FeaturedImages/' + IMG_EXT, html, re.I)
            if featured:
                image_url = featured[0]

        return text, image_url
    except Exception as e:
        return f"[Fehler beim Laden: {e}]", ""
