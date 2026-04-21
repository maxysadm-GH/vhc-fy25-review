"""Measure slide heights, screenshot each slide, then produce a multi-page PDF where each page's size matches the slide's natural content."""
import sys, os, io
from pathlib import Path
from playwright.sync_api import sync_playwright

REPORTS = Path(r"C:/Users/maxys/OneDrive - MBACIO/_CODE/reports")
TMP = Path(r"C:/Users/maxys/tmp")
SLIDE_W = 1280

def run(base):
    url = f"file:///{(REPORTS / (base + '.html')).as_posix()}?print=1"
    out_pdf = REPORTS / f"{base}-v7.pdf"
    shots_dir = TMP / f"{base}_shots"
    shots_dir.mkdir(exist_ok=True)

    with sync_playwright() as pw:
        browser = pw.chromium.launch()
        page = browser.new_page(viewport={"width": SLIDE_W, "height": 2000}, device_scale_factor=2)
        page.goto(url, wait_until="networkidle")
        page.wait_for_timeout(1500)  # let Chart.js finish first-frame even with animation:false
        page.emulate_media(media="print")  # measure under print CSS so heights match page.pdf
        page.wait_for_timeout(400)

        # Measure each slide
        heights = page.evaluate("""() => {
            return [...document.querySelectorAll('.slide')].map(s => {
                const r = s.getBoundingClientRect();
                return { w: Math.round(r.width), h: Math.round(s.scrollHeight) };
            });
        }""")
        print(f"[{base}] heights: {heights}")

        # Per-slide: hide others, PDF at natural height, screenshot for eyeball
        pdf_pages = []
        for i, dim in enumerate(heights):
            page.evaluate(
                """(idx) => {
                    document.querySelectorAll('.slide').forEach((s,i)=>{
                        s.style.display = (i === idx) ? 'flex' : 'none';
                    });
                    const deck = document.querySelector('.deck');
                    if(deck){ deck.style.gap = '0'; deck.style.padding = '0'; }
                    document.body.style.background = '#fff';
                }""",
                i,
            )
            page.wait_for_timeout(400)
            # Re-measure after display toggle (layout can shift)
            real_h = page.evaluate("""(idx) => {
                const s = document.querySelectorAll('.slide')[idx];
                return Math.max(s.scrollHeight, s.getBoundingClientRect().height);
            }""", i)
            real_h = int(real_h) + 2  # +2px safety to avoid rounding overflow
            # Screenshot for visual QA (full page, just this slide visible)
            png_path = shots_dir / f"slide{i+1}.png"
            page.screenshot(path=str(png_path), full_page=True)
            print(f"  slide{i+1}: measured={dim['w']}x{dim['h']}  real={real_h}  -> {png_path.name}")
            # PDF per slide at exact content size
            pdf_path = shots_dir / f"slide{i+1}.pdf"
            ph = real_h  # fit content, no dead-space floor
            page.pdf(
                path=str(pdf_path),
                width=f"{SLIDE_W}px",
                height=f"{ph}px",
                margin={"top":"0","right":"0","bottom":"0","left":"0"},
                print_background=True,
                prefer_css_page_size=False,
            )
            pdf_pages.append(pdf_path)

        browser.close()

    # Merge per-slide PDFs into one
    try:
        from pypdf import PdfWriter, PdfReader
    except ImportError:
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", "pypdf"])
        from pypdf import PdfWriter, PdfReader
    w = PdfWriter()
    for p in pdf_pages:
        r = PdfReader(str(p))
        for pg in r.pages:
            w.add_page(pg)
    with open(out_pdf, "wb") as f:
        w.write(f)
    # page count
    pc = len(PdfReader(str(out_pdf)).pages)
    print(f"[{base}] -> {out_pdf}  ({pc} pages)")

for base in ("vhc-board-2025-story-2026-04-20-editorial", "vhc-board-2025-story-2026-04-20-brand"):
    run(base)
