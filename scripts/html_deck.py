"""Export canonical Podium JSON as an offline animated HTML text scaffold.

Usage: python html_deck.py deck.json output/index.html
Complex media/chart layouts require design and browser review after export.
"""
import html
import json
import sys
from pathlib import Path


def escape(value):
    return html.escape(str(value), quote=True)


def entries(value):
    if isinstance(value, dict):
        return [f'{key}: {item}' for key, item in value.items() if isinstance(item, (str, int, float))]
    if isinstance(value, list):
        return [line for item in value for line in entries(item)]
    return [str(value)]


def build(spec):
    slides = spec.get('slides', [])
    if not slides:
        raise ValueError('Deck has no slides')
    rtl = str(spec.get('language', spec.get('lang', 'en'))).lower().startswith(('he', 'heb')) or spec.get('rtl', False)
    sections = []
    supported = {'type', 'variant', 'tone', 'motion', 'transition', 'notes', 'title', 'kicker', 'text', 'subtitle', 'points', 'items', 'steps', 'quote', 'takeaway'}
    for i, slide in enumerate(slides):
        unknown = sorted(set(slide) - supported)
        if unknown:
            print(f'Slide {i + 1}: adapt additional fields manually: {", ".join(unknown)}', file=sys.stderr)
        body = []
        for key in ('subtitle', 'text', 'quote', 'points', 'items', 'steps', 'takeaway'):
            if key in slide:
                body.extend(f'<p class="reveal">{escape(v)}</p>' for v in entries(slide[key]))
        sections.append(f'<section class="slide" aria-hidden="{str(i != 0).lower()}"><div class="inner"><p class="kicker reveal">{escape(slide.get("kicker", ""))}</p><h1 class="reveal">{escape(slide.get("title", "")).replace(chr(10), "<br>")}</h1>{"".join(body)}</div></section>')
    title = escape(spec.get('title', 'Presentation'))
    return '''<!doctype html><html lang="LANG" dir="DIR"><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>TITLE</title>
<style>
*{box-sizing:border-box}body{margin:0;background:#0b0b0c;color:#fff;font-family:Arial,sans-serif}.slide{position:absolute;inset:0;overflow:auto;display:none;padding:6vw 8vw 100px}.slide.active{display:flex;align-items:center;animation:slideIn .65s cubic-bezier(.2,.8,.2,1)}.inner{width:100%;max-width:1200px;margin:auto}h1{font-size:clamp(32px,5vw,76px);line-height:1.12;margin:0 0 30px}p{font-size:clamp(18px,2vw,28px);line-height:1.5;white-space:pre-wrap}.kicker{color:#7fd4ff;font-size:16px;letter-spacing:.08em}.active .reveal{animation:textIn .7s both}.active .reveal:nth-child(2){animation-delay:.12s}.active .reveal:nth-child(3){animation-delay:.24s}.active .reveal:nth-child(n+4){animation-delay:.36s}nav{position:fixed;bottom:20px;left:0;right:0;display:flex;justify-content:center;align-items:center;gap:16px;z-index:10}button{background:#24262a;color:white;border:1px solid #7fd4ff;padding:12px 18px;border-radius:12px;cursor:pointer}button:focus-visible{outline:3px solid #fff}progress{position:fixed;top:0;left:0;width:100%;height:5px;border:0;z-index:10;accent-color:#7fd4ff}@keyframes slideIn{from{opacity:0;transform:translateY(35px) scale(.98)}to{opacity:1;transform:none}}@keyframes textIn{from{opacity:0;transform:translateY(22px);filter:blur(4px)}to{opacity:1;transform:none;filter:none}}@media(prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important}}@media print{.slide,.slide.active{display:block;position:relative;min-height:100vh;break-after:page;animation:none}nav,progress{display:none}.reveal{opacity:1!important;animation:none!important}}
</style><progress value="1" max="COUNT" aria-label="Presentation progress"></progress>SECTIONS
<nav><button id="prev-slide" aria-label="Previous slide">←</button><span id="slide-status" aria-live="polite"></span><button id="next-slide" aria-label="Next slide">→</button><button id="fullscreen" aria-label="Toggle fullscreen">⛶</button></nav>
<script>RUNTIME</script></html>'''.replace('LANG', 'he' if rtl else 'en').replace('DIR', 'rtl' if rtl else 'ltr').replace('TITLE', title).replace('COUNT', str(len(slides))).replace('SECTIONS', ''.join(sections)).replace('RUNTIME', (Path(__file__).resolve().parents[1] / 'assets/html-runtime.js').read_text(encoding='utf-8'))


if __name__ == '__main__':
    source, output = map(Path, sys.argv[1:3])
    spec = json.loads(source.read_text(encoding='utf-8'))
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(build(spec), encoding='utf-8')
    print(output)
