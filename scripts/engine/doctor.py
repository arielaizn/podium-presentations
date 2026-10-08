"""Setup check for Podium. Tells the student exactly what is missing and how to fix it.

usage: python engine/doctor.py [brand.json]
"""
import importlib, json, platform, shutil, subprocess, sys
from pathlib import Path

ok = True


def line(good, what, fix=""):
    global ok
    ok &= good
    print(("  OK    " if good else "  FIX   ") + what + ("" if good or not fix else f"\n        -> {fix}"))


print(f"Podium doctor · Python {sys.version.split()[0]} · {platform.system()}")
line(sys.version_info >= (3, 10), "Python 3.10 or newer", "install Python 3.12 from python.org")
for mod, pipname in (("pptx", "python-pptx"), ("PIL", "pillow"), ("lxml", "lxml"), ("fontTools", "fonttools"), ("fitz", "pymupdf")):
    try:
        importlib.import_module(mod)
        line(True, f"package {pipname}")
    except ImportError:
        line(False, f"package {pipname}", f"pip install {pipname}")

renderer = None
if platform.system() == "Windows":
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-Command",
                            "$p = New-Object -ComObject PowerPoint.Application; $p.Version; $p.Quit()"],
                           capture_output=True, text=True, timeout=60)
        if r.returncode == 0 and r.stdout.strip():
            renderer = f"PowerPoint {r.stdout.strip()}"
    except Exception:
        pass
if not renderer:
    so = shutil.which("soffice") or shutil.which("libreoffice")
    mac = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    if so or mac.exists():
        renderer = "LibreOffice"
        try:
            importlib.import_module("fitz")
            line(True, "PDF rasterizer (pymupdf)")
        except ImportError:
            line(bool(shutil.which("pdftoppm")), "PDF rasterizer (pymupdf or pdftoppm)", "pip install pymupdf")
line(bool(renderer), f"slide renderer: {renderer or 'none found'}",
     "install Microsoft PowerPoint (Windows) or LibreOffice (free, any OS: libreoffice.org)")
if renderer == "LibreOffice":
    print("        note: LibreOffice previews ignore animations and may differ slightly from PowerPoint.")

if len(sys.argv) > 1:
    sys.path.insert(0, str(Path(__file__).parent))
    import deck
    print("fonts for", sys.argv[1])
    ok &= deck.check_fonts(json.loads(Path(sys.argv[1]).read_text(encoding="utf8")))

print("\nready to build decks." if ok else "\nfix the items above, then run doctor again.")
sys.exit(0 if ok else 1)

