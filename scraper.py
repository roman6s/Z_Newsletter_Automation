"""
Scraper for IBM Z DACH Community Blog.
Fetches all blog articles within a given date range.
"""

import re
import time
from datetime import date, datetime
from dataclasses import dataclass
from typing import List, Optional

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
    """Extract ASP.NET hidden form fields needed for postback pagination."""
    fields = {}
    for field_id in ("__VIEWSTATE", "__VIEWSTATEGENERATOR", "__EVENTVALIDATION"):
        tag = soup.find("input", {"id": field_id})
        if tag:
            fields[field_id] = tag.get("value", "")
    fields["ctl00$DefaultMasterHdnCommunityKey"] = "9a8b7fc3-b167-447a-8e14-adf93406eccc"
    fields["ScriptManager1_TSM"] = ""
    fields["StyleSheetManager1_TSSM"] = ""
    return fields


def _get_page_count(soup: BeautifulSoup) -> int:
    """Determine total number of pages from pager HTML."""
    # Pager contains numbered links via __doPostBack
    pager = soup.find("div", id=re.compile(r"SearchResultDataPager"))
    if not pager:
        return 1
    # Count ctl01$ctl0X links (page numbers)
    links = pager.find_all("a", href=re.compile(r"__doPostBack"))
    # Last link is usually "next >" - page numbers are all but last
    # Alternatively just look for the last numbered link
    page_links = [
        a for a in soup.find_all("a", href=True)
        if "SearchResultDataPager$ctl01$ctl" in str(a.get("href", ""))
    ]
    return max(len(page_links), 1)


def _postback_target_for_page(page: int) -> str:
    """Return the __EVENTTARGET for the given page number (1-based)."""
    # Page 1 = ctl01, page 2 = ctl02, ... page 5 = ctl05
    # After 5 pages there may be a "next" group (ctl02$ctl00)
    # Simple mapping for first set of pages
    return f"ctl00$MainCopy$ctl19$SearchResultDataPager$ctl01$ctl{page:02d}"


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
        print(f"Fetching blog listing page...")

    resp = session.get(BLOG_URL, timeout=30)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    all_entries = []
    stop_early = False  # Stop when we go past start_date

    # Collect from page 1
    page_entries = _extract_entries(soup)
    for e in page_entries:
        if e["date"] > end_date:
            continue
        if e["date"] < start_date:
            stop_early = True
            break
        all_entries.append(e)

    if not stop_early:
        # Paginate through remaining pages
        page_num = 2
        while True:
            hidden = _get_hidden_fields(soup)
            hidden["__EVENTTARGET"] = _postback_target_for_page(page_num)
            hidden["__EVENTARGUMENT"] = ""

            if verbose:
                print(f"  Fetching page {page_num}...")

            try:
                resp = session.post(BLOG_URL, data=hidden, timeout=30)
                resp.raise_for_status()
            except Exception as e:
                print(f"  Warning: Could not fetch page {page_num}: {e}")
                break

            soup = BeautifulSoup(resp.text, "html.parser")
            page_entries = _extract_entries(soup)

            if not page_entries:
                # Try "next" group button
                if verbose:
                    print(f"  No entries on page {page_num}, trying next group...")
                hidden["__EVENTTARGET"] = (
                    "ctl00$MainCopy$ctl19$SearchResultDataPager$ctl02$ctl00"
                )
                try:
                    resp = session.post(BLOG_URL, data=hidden, timeout=30)
                    resp.raise_for_status()
                    soup = BeautifulSoup(resp.text, "html.parser")
                    page_entries = _extract_entries(soup)
                    page_num = 2  # Reset within new group
                except Exception:
                    break

            if not page_entries:
                break

            added_any = False
            for e in page_entries:
                if e["date"] > end_date:
                    continue
                if e["date"] < start_date:
                    stop_early = True
                    break
                all_entries.append(e)
                added_any = True

            if stop_early or not added_any:
                break

            page_num += 1
            time.sleep(0.5)  # Be polite

    if verbose:
        print(f"Found {len(all_entries)} articles in date range.")

    # Deduplicate by URL
    seen = set()
    unique_entries = []
    for e in all_entries:
        if e["url"] not in seen:
            seen.add(e["url"])
            unique_entries.append(e)

    # Fetch full article text
    articles = []
    for i, e in enumerate(unique_entries):
        if verbose:
            print(f"  Loading article {i+1}/{len(unique_entries)}: {e['title'][:60]}...")
        full_text, image_url = _fetch_article_content(session, e["url"])
        articles.append(
            Article(
                title=e["title"],
                author=e["author"],
                url=e["url"],
                published=e["date"],
                full_text=full_text,
                image_url=image_url,
            )
        )
        time.sleep(0.3)

    return articles


def _fetch_article_content(session: requests.Session, url: str):
    """Fetch article text and first image URL. Returns (text, image_url)."""
    try:
        resp = session.get(url, timeout=30)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")

        # IBM Community (Higher Logic) article body
        body = None
        for selector in [
            {"id": re.compile(r"BlogPostBody|PostBody|ArticleBody", re.I)},
            {"class": re.compile(r"blog-post-body|post-content|entry-content", re.I)},
        ]:
            body = soup.find(attrs=selector)
            if body:
                break

        # Fallback: largest content div
        if not body:
            candidates = soup.find_all("div", class_=True)
            best = None
            best_len = 0
            for div in candidates:
                text = div.get_text(separator=" ", strip=True)
                if len(text) > best_len and len(text) < 50000:
                    best_len = len(text)
                    best = div
            body = best

        image_url = ""
        text = ""

        if body:
            # Extract first meaningful image before removing tags
            skip_keywords = ("avatar", "logo", "icon", "badge", "banner", "button")
            for img in body.find_all("img"):
                src = img.get("src", "") or img.get("data-src", "")
                if not src or not src.startswith("http"):
                    continue
                if any(k in src.lower() for k in skip_keywords):
                    continue
                # Prefer larger images (skip tiny ones via width/height attrs)
                w = img.get("width", "")
                h = img.get("height", "")
                try:
                    if int(w) < 100 or int(h) < 80:
                        continue
                except (ValueError, TypeError):
                    pass
                image_url = src
                break

            for tag in body.find_all(["nav", "aside", "script", "style", "footer"]):
                tag.decompose()
            text = body.get_text(separator="\n", strip=True)
            text = re.sub(r"\n{3,}", "\n\n", text)
            text = text[:8000]

        return text, image_url
    except Exception as e:
        return f"[Fehler beim Laden: {e}]", ""
