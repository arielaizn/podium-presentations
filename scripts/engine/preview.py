"""Build a deck and render every slide to PNG + one contact sheet, so the result is checked visually.

usage: python preview.py <deck.json> <out.pptx>
Renders with PowerPoint (Windows) or LibreOffice (any OS) and writes <out>_render/sheet.jpg.
"""
import json, os, platform, shutil, subprocess, sys
from pathlib import Path

from PIL import Image

sys.path.insert(0, str(Path(__file__).parent))
import deck  # noqa: E402


def render_powerpoint(pptx, out_dir):
    ps = f"""
$pp = New-Object -ComObject PowerPoint.Application
$p = $pp.Presentations.Open('{pptx}', -1, 0, 0)
$i = 0; foreach ($s in $p.Slides) {{ $i++; $s.Export(('{out_dir}\\p{{0:D2}}.png' -f $i), 'PNG', 1600, 900) }}
$p.Close(); $pp.Quit()
"""
    subprocess.run(["powershell", "-NoProfile", "-Command", ps], check=True, capture_output=True)


def render_libreoffice(pptx, out_dir):
    soffice = shutil.which("soffice") or shutil.which("libreoffice") or "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(pptx)], check=True, capture_output=True)
    pdf = out_dir / (Path(pptx).stem + ".pdf")
    try:
        import fitz  # PyMuPDF
        doc = fitz.open(pdf)
        for i, page in enumerate(doc, 1):
            page.get_pixmap(dpi=120).save(out_dir / f"p{i:02d}.png")
    except ImportError:
        subprocess.run(["pdftoppm", "-png", "-r", "120", str(pdf), str(out_dir / "p")], check=True)
        for f in out_dir.glob("p-*.png"):
            f.rename(out_dir / f"p{int(f.stem.split('-')[1]):02d}.png")


def sheet(out_dir, cols=2):
    ims = [Image.open(p).convert("RGB").resize((800, 450)) for p in sorted(out_dir.glob("p[0-9]*.png"))]
    rows = (len(ims) + cols - 1) // cols
    sh = Image.new("RGB", (cols * 810 + 10, rows * 460 + 10), (128, 128, 128))
    for i, im in enumerate(ims):
        sh.paste(im, (10 + (i % cols) * 810, 10 + (i // cols) * 460))
    sh.save(out_dir / "sheet.jpg", quality=86)
    return out_dir / "sheet.jpg"


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    spec_path, out = Path(sys.argv[1]).resolve(), Path(sys.argv[2]).resolve()
    deck.build(json.loads(spec_path.read_text(encoding="utf8")), str(out), spec_path.parent)
    rd = out.with_name(out.stem + "_render")
    if rd.exists():
        shutil.rmtree(rd)
    rd.mkdir()
    if platform.system() == "Windows":
        try:
            render_powerpoint(str(out), str(rd))
        except Exception:
            render_libreoffice(out, rd)
    else:
        render_libreoffice(out, rd)
    n_slides = len(json.loads(spec_path.read_text(encoding="utf8"))["slides"])
    n_png = len(list(rd.glob("p[0-9]*.png")))
    if n_png != n_slides:
        print(f"ERROR: rendered {n_png} of {n_slides} slides - PowerPoint could not open the file "
              f"(invalid XML from an engine change). Build slides one at a time to find the culprit.")
        sys.exit(1)
    print(sheet(rd))

