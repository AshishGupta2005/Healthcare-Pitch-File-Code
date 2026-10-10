"""
Healthcare pitch deck collector.

For every deck on the PitchDecky healthcare category it:
  1. reads the detail page (name, industry, amount, year, stage, website, deck link),
  2. downloads the deck assets it can reach (PDF/PPT/PPTX, slide images, cover image),
  3. extracts text (PDF text layer; OCR for images and scanned PDFs),
  4. writes decks_healthcare.xlsx with the columns:
     Name | Industry | Amount Raised | Year | Stage | Website  (+ source/notes columns)

Nothing is guessed. A field that cannot be read stays blank and the Notes column says why.

Setup:
  pip install requests beautifulsoup4 openpyxl pillow pytesseract pypdf pdf2image
  plus system packages: tesseract-ocr, poppler-utils

Usage:
  python scraper.py --out ./healthcare_decks            # everything
  python scraper.py --out ./healthcare_decks --limit 5  # quick test
Please respect each site's terms and the decks' copyright; use for internal research.
"""
import argparse, io, os, re, time, json
from urllib.parse import urljoin, urlparse
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

BASE = "https://www.pitchdecky.com"
LISTING = BASE + "/categories/healthcare"
UA = {"User-Agent": "Mozilla/5.0 (research script; contact: you@example.com)"}
STAGES = ["Pre-Seed", "Seed", "Series-A", "Series-B", "Series-C", "Series-D",
          "Series-E", "Later Stage", "Private", "Other"]
STAGE_RE = re.compile(r"^(%s)$" % "|".join(re.escape(s) for s in STAGES), re.I)
SKIP_HOSTS = ("pitchdecky.com", "leadloft.com", "facebook.com", "twitter.com",
              "linkedin.com", "webflow", "x.com")
DECK_HOSTS = ("slideshare.net", "docsend.com", "speakerdeck.com", "drive.google.com",
              "docs.google.com", "scribd.com", "pitch.com", "canva.com", "figma.com")

s = requests.Session(); s.headers.update(UA)

def get(url, **kw):
    r = s.get(url, timeout=30, **kw); r.raise_for_status(); return r

def slugify(t):
    return re.sub(r"[^a-z0-9]+", "-", t.lower()).strip("-") or "unknown"

def listing_urls():
    soup = BeautifulSoup(get(LISTING).text, "html.parser")
    seen, out = set(), []
    for a in soup.find_all("a", href=True):
        h = urljoin(BASE, a["href"])
        if "/pitch-deck/" in h and h not in seen:
            seen.add(h); out.append(h)
    return out

def label_value(lines, label):
    """Value on the line after 'Label:' (Webflow renders label and value separately)."""
    for i, ln in enumerate(lines):
        if ln.strip().rstrip(":").lower() == label.lower():
            for nxt in lines[i + 1:i + 4]:
                if nxt.strip():
                    return nxt.strip()
    return ""

def parse_detail(url):
    html = get(url).text
    soup = BeautifulSoup(html, "html.parser")
    notes = []
    h1 = soup.find("h1")
    name = re.sub(r"\s*Pitch Deck\s*$", "", h1.get_text(strip=True)) if h1 else ""
    meta = soup.find("meta", attrs={"name": "description"})
    desc = meta["content"] if meta and meta.get("content") else ""

    lines = soup.get_text("\n").splitlines()
    industry = label_value(lines, "Industry")
    amount = label_value(lines, "Amount Raised")
    year = label_value(lines, "Founded")
    stage = label_value(lines, "Stage")

    # Stage field is sometimes bound to the website link on this site; only accept real stage names.
    if not STAGE_RE.match(stage or ""):
        m = re.search(r"fundraise \(([^)]+)\)", desc)
        stage = m.group(1).strip() if m and STAGE_RE.match(m.group(1).strip()) else ""
        if stage: notes.append("Stage taken from page meta description.")
        else: notes.append("Stage not stated on page.")
    if not re.match(r"^\$?[\d.,]+\s*[KMBkmb]?\+?$", amount or ""):
        if amount: notes.append(f"Amount field unusual: {amount!r}")
        else: notes.append("Amount not stated on page.")
    if year and not re.match(r"^(19|20)\d\d$", year):
        notes.append(f"Year field unusual: {year!r}"); year = ""
    if year: notes.append("Year is the page's 'Founded' value; may not be the funding year.")

    website = ""
    m = re.search(r"\(((?:https?://)?[\w.-]+\.[a-z]{2,}[^)\s]*)\)\s+pitch deck", desc, re.I)
    if m: website = m.group(1)
    if not website:
        for a in soup.find_all("a", href=True):
            h = a["href"]
            if h.startswith("http") and not any(x in urlparse(h).netloc for x in SKIP_HOSTS) \
               and not any(x in urlparse(h).netloc for x in DECK_HOSTS):
                website = h; break
    if website and not website.startswith("http"): website = "https://" + website
    if not website: notes.append("Website not found.")

    deck_urls = []
    for tag in soup.find_all(["iframe", "a", "embed"]):
        h = tag.get("src") or tag.get("href") or ""
        if not h.startswith("http"): continue
        host = urlparse(h).netloc
        if any(d in host for d in DECK_HOSTS) or re.search(r"\.(pdf|pptx?|key)(\?|$)", h, re.I):
            if h not in deck_urls: deck_urls.append(h)
    if not deck_urls: notes.append("No deck link found on page.")

    og = soup.find("meta", property="og:image")
    body_text = "\n".join(l.strip() for l in lines if l.strip())
    return dict(name=name, industry=industry or "Healthcare", amount=amount, year=year,
                stage=stage, website=website, deck_urls=deck_urls,
                cover=og["content"] if og and og.get("content") else "",
                text=body_text, source=url, notes=notes)

# ---------- assets ----------
def save_bin(url, path):
    r = get(url, stream=True)
    with open(path, "wb") as f:
        for chunk in r.iter_content(65536): f.write(chunk)
    return path

def ocr_image(path):
    try:
        import pytesseract; from PIL import Image
        return pytesseract.image_to_string(Image.open(path))
    except Exception as e:
        return f"[OCR failed: {e}]"

def pdf_text(path):
    txt = ""
    try:
        from pypdf import PdfReader
        txt = "\n\n".join((p.extract_text() or "") for p in PdfReader(path).pages)
    except Exception: pass
    if len(txt.strip()) < 200:  # likely scanned/image-only -> OCR
        try:
            from pdf2image import convert_from_path; import pytesseract
            pages = convert_from_path(path, dpi=150)
            txt = "\n\n".join(f"--- slide {i+1} ---\n" + pytesseract.image_to_string(p)
                               for i, p in enumerate(pages))
        except Exception as e:
            txt += f"\n[OCR failed: {e}]"
    return txt

def slideshare_images(embed_url):
    """Best effort: pull slide image URLs from the embed page. May return nothing."""
    try:
        html = get(embed_url).text
    except Exception:
        return []
    urls = re.findall(r"https://image\.slidesharecdn\.com/[^\"'\s\\]+?\.jpg", html)
    seen, out = set(), []
    for u in urls:
        key = re.sub(r"-\d+-1024", "", u)
        if u not in seen and key not in seen:
            seen.add(u); seen.add(key); out.append(u)
    return out

def collect_assets(d, folder):
    os.makedirs(folder, exist_ok=True)
    got = []
    with open(os.path.join(folder, "page_text.txt"), "w", encoding="utf-8") as f: f.write(d["text"])
    if d["cover"]:
        try:
            p = save_bin(d["cover"], os.path.join(folder, "cover.png")); got.append("cover image")
            t = ocr_image(p)
            if t.strip(): open(os.path.join(folder, "cover_ocr.txt"), "w", encoding="utf-8").write(t)
        except Exception as e: d["notes"].append(f"Cover image failed: {e}")
    for i, u in enumerate(d["deck_urls"]):
        try:
            if re.search(r"\.(pdf|pptx?|key)(\?|$)", u, re.I):
                ext = re.search(r"\.(pdf|pptx?|key)", u, re.I).group(1).lower()
                p = save_bin(u, os.path.join(folder, f"deck_{i}.{ext}")); got.append(f"deck_{i}.{ext}")
                if ext == "pdf":
                    open(os.path.join(folder, f"deck_{i}_text.txt"), "w", encoding="utf-8").write(pdf_text(p))
            elif "slideshare.net" in u:
                imgs = slideshare_images(u)
                if not imgs:
                    d["notes"].append(f"SlideShare slides not downloadable automatically: {u}"); continue
                sdir = os.path.join(folder, "slides"); os.makedirs(sdir, exist_ok=True)
                texts = []
                for n, iu in enumerate(imgs, 1):
                    ip = save_bin(iu, os.path.join(sdir, f"slide_{n:02d}.jpg"))
                    texts.append(f"--- slide {n} ---\n{ocr_image(ip)}")
                open(os.path.join(folder, "slides_ocr.txt"), "w", encoding="utf-8").write("\n\n".join(texts))
                got.append(f"{len(imgs)} slide images + OCR")
                try:  # bundle slides into a PDF
                    from PIL import Image
                    ims = [Image.open(os.path.join(sdir, f"slide_{n:02d}.jpg")).convert("RGB")
                           for n in range(1, len(imgs) + 1)]
                    ims[0].save(os.path.join(folder, "deck.pdf"), save_all=True, append_images=ims[1:])
                    got.append("deck.pdf")
                except Exception as e: d["notes"].append(f"PDF bundle failed: {e}")
            else:
                d["notes"].append(f"Deck hosted at {u} (needs manual download or login).")
        except Exception as e:
            d["notes"].append(f"Deck asset failed ({u}): {e}")
    return got

# ---------- excel ----------
def write_xlsx(rows, path):
    cols = ["Name", "Industry", "Amount Raised", "Year", "Stage", "Website",
            "Source URL", "Deck Link(s)", "Assets Saved", "Notes"]
    wb = Workbook(); ws = wb.active; ws.title = "Healthcare Decks"; ws.append(cols)
    for r in rows:
        ws.append([r["name"], r["industry"], r["amount"],
                   int(r["year"]) if r["year"] else "", r["stage"], r["website"],
                   r["source"], "\n".join(r["deck_urls"]), ", ".join(r.get("assets", [])),
                   " ".join(r["notes"])])
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78"); c.alignment = Alignment(horizontal="center")
    for row in ws.iter_rows(min_row=2):
        for c in row: c.font = Font(name="Arial"); c.alignment = Alignment(vertical="top", wrap_text=True)
    for col, w in zip("ABCDEFGHIJ", [26, 12, 14, 8, 12, 30, 48, 48, 28, 70]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
    wb.save(path)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="healthcare_decks"); ap.add_argument("--limit", type=int)
    ap.add_argument("--delay", type=float, default=1.5, help="seconds between requests")
    ap.add_argument("--no-assets", action="store_true")
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    urls = listing_urls()
    if a.limit: urls = urls[:a.limit]
    print(f"{len(urls)} deck pages")
    rows = []
    for i, u in enumerate(urls, 1):
        try:
            d = parse_detail(u)
            if not a.no_assets:
                d["assets"] = collect_assets(d, os.path.join(a.out, slugify(d["name"] or u.rsplit('/',1)[-1])))
            rows.append(d); print(f"[{i}/{len(urls)}] {d['name']}  {d['amount']} {d['year']} {d['stage']}")
        except Exception as e:
            print(f"[{i}/{len(urls)}] FAILED {u}: {e}")
            rows.append(dict(name=u.rsplit('/',1)[-1], industry="", amount="", year="", stage="",
                             website="", deck_urls=[], source=u, notes=[f"Page failed: {e}"], assets=[]))
        write_xlsx(rows, os.path.join(a.out, "decks_healthcare.xlsx"))  # save progress each row
        time.sleep(a.delay)
    json.dump(rows, open(os.path.join(a.out, "raw_data.json"), "w"), indent=1, default=str)
    print("Done ->", os.path.join(a.out, "decks_healthcare.xlsx"))

if __name__ == "__main__":
    main()
