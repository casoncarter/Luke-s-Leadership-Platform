import os
import re
import zipfile
from pathlib import Path
from html import escape

# ----------------------------
# CONFIG: point to your DOCX files
# ----------------------------
DOCS = [
    ("business", "Business and Entrepenurship Platform.docx"),
    ("sports", "THE SPORTS LEADERSHIP PLATFORM.docx"),
    ("if", "If Leadership Platform.docx"),
    ("guitar", "THE GUITAR TRAINING PLATFORM.docx"),
    ("spiritual", "Spiritual Leadership.docx"),
    ("dadjokes", "THE DAD JOKE MASTERY PLATFORM.docx"),
    ("power", "Power platform.docx"),
    ("art", "🎨 THE ADVANCED SKETCH PLATFORM.docx"),
]

# Optional: if a DOCX contains extra platforms appended, cut at these markers.
# (Your "Sports" + "Spiritual" docs appear to include an appended Art/Sketch section.)
END_MARKERS = {
    "sports": ["🎨", "THE ADVANCED SKETCH", "See More. Choose More. Create More."],
    "spiritual": ["See More. Choose More. Create More.", "🎨", "THE ADVANCED SKETCH"],
}

OUT_DIR = Path("platforms")
IMG_DIR = Path("images")

# ----------------------------
# Helpers
# ----------------------------
def read_docx_paragraphs(docx_path: Path) -> list[str]:
    """Lightweight .docx text extraction without python-docx, for portability."""
    # .docx is a zip; main text is word/document.xml
    import xml.etree.ElementTree as ET

    with zipfile.ZipFile(docx_path, "r") as z:
        xml = z.read("word/document.xml")
    root = ET.fromstring(xml)

    # Word uses namespaces
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    paras = []
    for p in root.findall(".//w:p", ns):
        texts = []
        for t in p.findall(".//w:t", ns):
            if t.text:
                texts.append(t.text)
        line = "".join(texts).strip()
        if line:
            # normalize whitespace
            line = re.sub(r"[ \t]+", " ", line)
            paras.append(line)
    return paras

def cut_section(paras: list[str], end_markers: list[str] | None) -> list[str]:
    if not end_markers:
        return paras
    for i, line in enumerate(paras):
        for m in end_markers:
            if m.lower() in line.lower():
                return paras[:i]
    return paras

def extract_docx_images(docx_path: Path, out_folder: Path) -> list[str]:
    """
    Extract all images from /word/media inside the DOCX.
    Returns list of extracted image filenames (relative to out_folder).
    """
    out_folder.mkdir(parents=True, exist_ok=True)
    extracted = []
    with zipfile.ZipFile(docx_path, "r") as z:
        media_files = [f for f in z.namelist() if f.startswith("word/media/")]
        for f in media_files:
            data = z.read(f)
            name = Path(f).name  # e.g. image1.jpeg
            target = out_folder / name
            target.write_bytes(data)
            extracted.append(name)
    return extracted

def classify_line(line: str) -> str:
    """
    Heuristics to map doc lines to headings/lists.
    """
    if re.match(r"^(MODULE|Module)\s+\d+\b", line):
        return "h2"
    if re.match(r"^(Executive Summary|EXECUTIVE SUMMARY)\b", line):
        return "h2"
    if re.match(r"^(Core Principle:)\b", line):
        return "core"
    if line.strip() in {"Framework", "Case Studies", "Challenge Exercise", "What This Module Builds", "Real-World Applications", "Key Takeaways", "Closing", "Conclusion"}:
        return "h3"
    # lines that look like bullets in docs (often sentence fragments listed line-by-line)
    # We'll treat short lines after a "Framework"/"Key Takeaways" heading as bullets downstream.
    return "p"

def build_content_html(paras: list[str]) -> str:
    """
    Converts lines into semantic HTML sections with automatic bullet grouping.
    """
    out = []
    in_ul = False

    def close_ul():
        nonlocal in_ul
        if in_ul:
            out.append("</ul>")
            in_ul = False

    prev_kind = None
    for idx, raw in enumerate(paras):
        line = raw.strip()
        kind = classify_line(line)

        # auto-bullet grouping:
        # If we had an h3 like "Framework" or "Key Takeaways", then a run of
        # short-ish lines will become bullets until the next heading.
        if kind in {"h2", "h3"}:
            close_ul()

        if kind == "h2":
            out.append(f'<h2 class="sec">{escape(line)}</h2>')
        elif kind == "h3":
            out.append(f'<h3 class="subsec">{escape(line)}</h3>')
        elif kind == "core":
            close_ul()
            out.append(f'<p class="core"><strong>{escape("Core Principle:")}</strong> {escape(line.split("Core Principle:",1)[1].strip())}</p>')
        else:
            # decide bullet vs paragraph based on context
            bullet_context = (prev_kind == "h3")  # immediately after a subsection header
            looks_bulleted = (len(line) <= 110 and line.endswith((".", ")", "—")))

            if bullet_context and looks_bulleted:
                if not in_ul:
                    out.append('<ul class="bul">')
                    in_ul = True
                out.append(f"<li>{escape(line)}</li>")
            else:
                close_ul()
                out.append(f'<p class="p">{escape(line)}</p>')

        prev_kind = kind

    close_ul()
    return "\n".join(out)

def wrap_page(platform_key: str, title: str, content_html: str, extracted_images: list[str]) -> str:
    """
    Full standalone HTML page wrapper. Uses image gallery block so all extracted images
    can be placed visibly (you can later move them into specific modules).
    """
    # basic nav linking to your other pages
    nav = """
      <a href="index.html">All Modules</a>
      <a href="platforms/business.html">Business</a>
      <a href="platforms/sports.html">Sports</a>
      <a href="platforms/if.html">If—</a>
      <a href="platforms/guitar.html">Guitar</a>
      <a href="platforms/spiritual.html">Spiritual</a>
      <a href="platforms/dadjokes.html">Dad Jokes</a>
      <a href="platforms/power.html">Power</a>
      <a href="platforms/art.html">Art/Sketch</a>
    """

    # image gallery (so nothing is “missing” even before you hand-place images)
    gallery = ""
    if extracted_images:
        imgs = "\n".join(
            f'<figure class="img"><img src="../images/{platform_key}/{escape(name)}" alt="{escape(name)}"><figcaption>{escape(name)}</figcaption></figure>'
            for name in extracted_images
        )
        gallery = f"""
        <section class="card">
          <h2 class="sec">Images from the Word document</h2>
          <p class="note">These were extracted automatically. You can later move each image into the exact module where it belongs.</p>
          <div class="imggrid">
            {imgs}
          </div>
        </section>
        """

    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>{escape(title)}</title>
  <style>
    :root {{
      --bg:#0b0f17; --panel:#101827; --text:#e9eefc; --muted:#b7c2df; --line:rgba(255,255,255,.12);
      --accent:#ff7a18; --shadow:0 18px 60px rgba(0,0,0,.45);
      --r:18px; --r2:24px; --max:1100px;
      --sans: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial;
    }}
    *{{box-sizing:border-box}}
    body{{margin:0;font-family:var(--sans);color:var(--text);background:linear-gradient(180deg,#070a11,var(--bg));}}
    .wrap{{max-width:var(--max);margin:0 auto;padding:0 22px;}}
    header{{position:sticky;top:0;z-index:40;backdrop-filter: blur(12px);background:rgba(10,15,24,.75);border-bottom:1px solid var(--line);}}
    .top{{display:flex;gap:14px;align-items:center;justify-content:space-between;padding:14px 0;flex-wrap:wrap;}}
    .brand{{font-weight:800;letter-spacing:.2px}}
    nav a{{color:var(--muted);text-decoration:none;border:1px solid var(--line);padding:8px 12px;border-radius:999px;font-size:13px;}}
    nav a:hover{{color:var(--text);border-color:rgba(255,255,255,.22)}}
    .hero{{padding:26px 0 10px;}}
    .hero h1{{margin:0 0 8px;font-size:34px;letter-spacing:-.02em}}
    .hero p{{margin:0;color:var(--muted);max-width:75ch;line-height:1.55}}
    .card{{border:1px solid var(--line);background:rgba(16,24,39,.58);border-radius:var(--r2);box-shadow:var(--shadow);padding:18px;margin:14px 0;}}
    .sec{{margin:0 0 10px;font-size:20px;letter-spacing:-.01em}}
    .subsec{{margin:14px 0 8px;font-size:16px;color:var(--text)}}
    .p{{margin:0 0 10px;color:var(--muted);line-height:1.6}}
    .core{{margin:0 0 10px;color:var(--muted);line-height:1.6;border-left:3px solid rgba(255,122,24,.55);padding-left:12px}}
    .bul{{margin:6px 0 10px;padding-left:18px;color:var(--muted);line-height:1.55}}
    .bul li{{margin:6px 0}}
    .note{{color:rgba(255,255,255,.65);font-size:13px;line-height:1.5}}
    .imggrid{{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px}}
    .img{{border:1px solid var(--line);border-radius:16px;overflow:hidden;background:rgba(0,0,0,.12);margin:0}}
    .img img{{width:100%;height:160px;object-fit:cover;display:block}}
    .img figcaption{{padding:8px 10px;color:rgba(255,255,255,.65);font-size:12px}}
    @media (max-width: 920px){{ .imggrid{{grid-template-columns:1fr}} }}
    footer{{border-top:1px solid var(--line);padding:26px 0 44px;color:rgba(255,255,255,.62);font-size:13px;margin-top:24px}}
  </style>
</head>
<body>
<header>
  <div class="wrap">
    <div class="top">
      <div class="brand">Luke Carter · Platforms</div>
      <nav aria-label="platform navigation">{nav}</nav>
    </div>
  </div>
</header>

<div class="wrap">
  <section class="hero">
    <h1>{escape(title)}</h1>
    <p>This page is generated directly from the Word document so no module content gets lost. If you update the doc, re-run the script.</p>
  </section>

  <section class="card">
    {content_html}
  </section>

  {gallery}

  <footer>
    Tip: If a doc contains multiple platforms appended, adjust END_MARKERS in the script so each page cuts off at the correct boundary.
  </footer>
</div>
</body>
</html>
"""

def main():
    OUT_DIR.mkdir(exist_ok=True)
    IMG_DIR.mkdir(exist_ok=True)

    for key, filename in DOCS:
        docx = Path(filename)
        if not docx.exists():
            print(f"[SKIP] Missing: {filename}")
            continue

        print(f"[READ] {filename}")
        paras = read_docx_paragraphs(docx)
        paras = cut_section(paras, END_MARKERS.get(key))

        # Title: first non-empty line
        title = paras[0] if paras else f"{key.title()} Platform"

        # Extract images
        img_out = IMG_DIR / key
        extracted = extract_docx_images(docx, img_out)

        content_html = build_content_html(paras)
        html = wrap_page(key, title, content_html, extracted)

        out_file = OUT_DIR / f"{key}.html"
        out_file.write_text(html, encoding="utf-8")
        print(f"[WRITE] {out_file} (+ {len(extracted)} images)")

    print("\nDone. Upload /platforms and /images to GitHub.")

if __name__ == "__main__":
    main()
