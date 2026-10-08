"""Install Google Fonts families for the current Windows user (no admin needed).

Variable fonts are turned into static Regular + Bold files named exactly after the family,
so PowerPoint lists them under the plain name the brand uses (a variable font with an optical-size
axis would otherwise show up as e.g. "Fraunces 9pt").

usage: python install_fonts.py "Family One" "Family Two" ...      (needs: pip install fonttools)
"""
import io, json, os, sys, urllib.request
from pathlib import Path

DST = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "Microsoft" / "Windows" / "Fonts"
API = "https://api.github.com/repos/google/fonts/contents/{lic}/{d}"


def listing(d):
    for lic in ("ofl", "apache", "ufl"):
        try:
            req = urllib.request.Request(API.format(lic=lic, d=d), headers={"User-Agent": "font-installer"})
            return json.loads(urllib.request.urlopen(req, timeout=30).read())
        except Exception:
            continue
    return None


def register(path, display):
    import winreg
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows NT\CurrentVersion\Fonts", 0,
                        winreg.KEY_SET_VALUE) as key:
        winreg.SetValueEx(key, display + " (TrueType)", 0, winreg.REG_SZ, str(path))


def static_cuts(blob, family):
    """Variable font bytes -> {style: static font bytes}, family name set to `family`."""
    from fontTools.ttLib import TTFont
    from fontTools.varLib.instancer import instantiateVariableFont
    out = {}
    probe = TTFont(io.BytesIO(blob))
    axes = {a.axisTag: a for a in probe["fvar"].axes}
    for style, wght in (("Regular", 400), ("Bold", 700)):
        f = TTFont(io.BytesIO(blob))
        loc = {}
        for tag, a in axes.items():
            if tag == "wght":
                loc[tag] = max(a.minValue, min(a.maxValue, wght))
            elif tag == "opsz":
                loc[tag] = max(a.minValue, min(a.maxValue, 48))  # display-friendly optical size
            else:
                loc[tag] = a.defaultValue
        inst = instantiateVariableFont(f, loc)
        name = inst["name"]
        for rec in list(name.names):
            if rec.nameID in (16, 17, 21, 22, 25):
                name.removeNames(nameID=rec.nameID)
        ps = family.replace(" ", "") + "-" + style
        for nid, val in ((1, family), (2, style), (4, f"{family} {style}" if style != "Regular" else family), (6, ps)):
            name.setName(val, nid, 3, 1, 0x409)
            name.setName(val, nid, 1, 0, 0)
        inst["OS/2"].usWeightClass = wght
        inst["OS/2"].fsSelection = (inst["OS/2"].fsSelection & ~0b1100001) | (0b100000 if style == "Bold" else 0b1000000)
        inst["head"].macStyle = 1 if style == "Bold" else 0
        buf = io.BytesIO()
        inst.save(buf)
        out[style] = buf.getvalue()
    return out


def install(family):
    d = family.lower().replace(" ", "")
    files = listing(d)
    if not files:
        return f"{family}: NOT FOUND on Google Fonts"
    DST.mkdir(parents=True, exist_ok=True)
    ttfs = [f for f in files if f["name"].endswith(".ttf") and "Italic" not in f["name"]]
    if any(f.get("type") == "dir" and f["name"] == "static" for f in files):
        st = [f for f in (listing(d + "/static") or []) if f["name"].endswith(".ttf") and "Italic" not in f["name"]]
        if st:
            ttfs = st
    n = 0
    for f in ttfs:
        blob = urllib.request.urlopen(f["download_url"], timeout=60).read()
        if "[" in f["name"]:  # variable font -> static cuts with clean names
            try:
                for style, data in static_cuts(blob, family).items():
                    path = DST / f"{family.replace(' ', '')}-{style}.ttf"
                    path.write_bytes(data)
                    register(path, f"{family} {style}" if style != "Regular" else family)
                    n += 1
                continue
            except ImportError:
                print("  (pip install fonttools for clean static cuts; installing the variable file as-is)")
        safe = f["name"].replace("[", "-").replace("]", "").replace(",", "-")
        path = DST / safe
        path.write_bytes(blob)
        register(path, safe.rsplit(".", 1)[0])
        n += 1
    return f"{family}: {n} files"


if __name__ == "__main__":
    if os.name != "nt":
        sys.exit("Windows only. On macOS/Linux download the family from fonts.google.com and install it normally.")
    for fam in sys.argv[1:]:
        print(install(fam))
