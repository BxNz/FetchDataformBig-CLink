"""
Big C product scraper -> Excel (name + image URL + product link + optional embedded image)

Run examples:
    python bigc_scraper.py                              (uses urls.txt if it exists, else the default link)
    python bigc_scraper.py URL1 URL2 URL3 --embed       (several category links)
    python bigc_scraper.py --pages 5 --embed            (up to 5 pages per link)

urls.txt: one link per line (lines starting with # are ignored)
"""
import argparse
import io
import os
import re
import time
from urllib.parse import parse_qs, urlencode, urlparse, urlunparse

import requests
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from playwright.sync_api import sync_playwright

DEFAULT_URL = "https://www.bigc.co.th/en/category/battery?sort=nameasc&limit=90"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Referer": "https://www.bigc.co.th/",
    "Accept": "image/avif,image/webp,image/png,image/*,*/*;q=0.8",
}

# Runs inside the page: finds product links that contain an <img>
JS_EXTRACT = """
() => {
  const out = [], seen = new Set();
  document.querySelectorAll('a[href*="/product/"]').forEach(a => {
    const href = a.href;
    if (seen.has(href)) return;
    const img = a.querySelector('img');
    if (!img) return;
    let src = img.currentSrc || img.src || img.getAttribute('data-src') || '';
    try {
      const u = new URL(src, location.href);
      if (u.searchParams.get('url')) src = decodeURIComponent(u.searchParams.get('url'));
    } catch (e) {}
    let name = (img.alt || '').trim();
    if (!name) {
      const t = a.innerText.split('\\n').map(s => s.trim()).filter(Boolean);
      name = t.sort((x, y) => y.length - x.length)[0] || '';
    }
    seen.add(href);
    out.push({name, image: src, link: href});
  });
  return out;
}
"""


def set_page(url: str, n: int) -> str:
    """Return url with &page=n (page 1 keeps the original url)."""
    if n <= 1:
        return url
    parts = urlparse(url)
    q = parse_qs(parts.query)
    q["page"] = [str(n)]
    return urlunparse(parts._replace(query=urlencode(q, doseq=True)))


def scroll_all(page):
    last = 0
    for _ in range(30):
        page.mouse.wheel(0, 2500)
        page.wait_for_timeout(700)
        count = page.evaluate("document.querySelectorAll('a[href*=\"/product/\"]').length")
        if count == last and count > 0:
            break
        last = count


def scrape(urls, max_pages=1):
    results, seen = [], set()
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  # visible window avoids most bot blocks
        ctx = browser.new_context(user_agent=HEADERS["User-Agent"], viewport={"width": 1400, "height": 900})
        page = ctx.new_page()
        for url in urls:
            category = url.split("/category/")[-1].split("?")[0] if "/category/" in url else url
            for n in range(1, max_pages + 1):
                target = set_page(url, n)
                print(f"Loading {target}")
                try:
                    page.goto(target, wait_until="domcontentloaded", timeout=60000)
                    page.wait_for_timeout(4000)
                    scroll_all(page)
                    items = page.evaluate(JS_EXTRACT)
                except Exception as e:
                    print(f"  failed: {e}")
                    break
                new = [it for it in items if it["link"] not in seen]
                for it in new:
                    seen.add(it["link"])
                    it["category"] = category
                    results.append(it)
                print(f"  {len(new)} new products")
                if not new:  # empty or repeated page -> no more pages
                    break
        browser.close()
    return results


SESSION = requests.Session()
SESSION.headers.update(HEADERS)


def image_candidates(url: str):
    """Try the untransformed image before Cloudflare's potentially resized version."""
    if url.startswith("//"):
        url = "https:" + url
    plain = re.sub(r"/cdn-cgi/image/[^/]+/", "/", url)
    return [plain, url] if plain != url else [url]


def fetch_image(url: str) -> io.BytesIO:
    """Download the original image bytes, falling back to transformed URLs if needed."""
    last_err = None
    for u in image_candidates(url):
        for attempt in range(3):
            try:
                r = SESSION.get(u, timeout=30)
                if r.status_code in (403, 429, 500, 502, 503):
                    last_err = f"{r.status_code} for {u}"
                    time.sleep(1.5 * (attempt + 1))
                    continue
                r.raise_for_status()
                return io.BytesIO(r.content)
            except requests.RequestException as e:
                last_err = str(e)
                time.sleep(1.5 * (attempt + 1))
    raise RuntimeError(last_err or "download failed")


def save_excel(items, path="bigc_products.xlsx", embed=False):
    wb = Workbook()
    ws = wb.active
    ws.title = "Products"
    ws.append(["#", "Category", "Name", "Image URL", "Product Link"] + (["Image"] if embed else []))
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 60
    ws.column_dimensions["D"].width = 50
    ws.column_dimensions["E"].width = 50
    if embed:
        ws.column_dimensions["F"].width = 14

    for i, it in enumerate(items, 1):
        ws.append([i, it.get("category", ""), it["name"], it["image"], it["link"]])
        if embed and it["image"]:
            try:
                img = XLImage(fetch_image(it["image"]))
                col, row = 5, i  # zero-based: column F, row i+1
                m = 9525 * 3                  # 3px margin inside the cell
                cell_w = (14 * 7 + 5) * 9525  # column F width 14 chars, in EMU
                cell_h = 75 * 12700           # row height 75pt, in EMU
                img.anchor = TwoCellAnchor(
                    editAs="twoCell",
                    _from=AnchorMarker(col=col, colOff=m, row=row, rowOff=m),
                    to=AnchorMarker(col=col, colOff=cell_w - m, row=row, rowOff=cell_h - m),
                )
                ws.add_image(img)
                ws.row_dimensions[i + 1].height = 75
            except Exception as e:
                print(f"  image failed for row {i}: {e}")
            time.sleep(0.4)

    wb.save(path)
    print(f"Saved {len(items)} products -> {path}")


def read_urls(args_urls):
    if args_urls:
        return args_urls
    if os.path.exists("urls.txt"):
        with open("urls.txt", encoding="utf-8") as f:
            urls = [l.strip() for l in f if l.strip() and not l.strip().startswith("#")]
        if urls:
            return urls
    return [DEFAULT_URL]


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("urls", nargs="*", help="one or more category links")
    ap.add_argument("--out", default="bigc_products.xlsx")
    ap.add_argument("--pages", type=int, default=1, help="max pages per link (default 1)")
    ap.add_argument("--embed", action="store_true", help="embed images inside Excel cells")
    args = ap.parse_args()

    data = scrape(read_urls(args.urls), args.pages)
    if not data:
        print("No products found. The page may have blocked the browser or its layout changed.")
    else:
        save_excel(data, args.out, args.embed)