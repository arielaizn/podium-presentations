"""Verify the structure of a presentation package; visual and hosted-access QA remain mandatory.

Manifest example:
{"pptx":"presentation.pptx","pdf":"presentation.pdf","html":"site/index.html",
 "bonuses":["bonuses/01.pdf","bonuses/02.pdf","bonuses/03.pdf","bonuses/04.pdf","bonuses/05.pdf","bonuses/06.pdf"],
 "published_url":"https://verified-provider-url.example"}
Use --require-url only after successful independent browser verification. No network request is made.
"""
import argparse
import hashlib
import json
import re
import sys
from pathlib import Path
from html.parser import HTMLParser
from urllib.parse import urlparse
from pptx import Presentation
try:
    import fitz
except ImportError:
    raise SystemExit("Package QA requires PyMuPDF. Install it through the authorized package workflow before final verification.")


def verify(root, manifest, require_url=False):
    errors, facts = [], {}
    root = root.resolve()
    def path(value):
        if not isinstance(value, str) or not value:
            raise ValueError('Every deliverable must have a nonempty relative file path')
        p = (root/value).resolve()
        if not p.is_relative_to(root) or not p.is_file():
            raise ValueError(f'Missing or out-of-package file: {value}')
        if not p.stat().st_size:
            raise ValueError(f'Empty file: {value}')
        return p
    try:
        pptx, pdf, html = (path(manifest.get(k)) for k in ('pptx','pdf','html'))
        slides = len(Presentation(pptx).slides)
        with fitz.open(pdf) as doc:
            pages = len(doc)
            if not pages or any(not p.get_text().strip() for p in doc):
                errors.append('Main PDF has empty/non-selectable-text pages; inspect rendering and accessibility')
        facts.update(slides=slides,pdf_pages=pages)
        if pages != slides: errors.append('PPTX slide count and main PDF page count differ')
        content = html.read_text(encoding='utf-8')
        class SlideParser(HTMLParser):
            count = 0
            def handle_starttag(self, tag, attrs):
                if tag == 'section' and 'slide' in dict(attrs).get('class','').split():
                    self.count += 1
        parsed = SlideParser()
        parsed.feed(content)
        html_slides = parsed.count
        facts['html_slides'] = html_slides
        if html_slides != slides: errors.append('HTML section.slide count differs from PPTX slide count')
        if 'reveal' not in content: errors.append('HTML lacks staged text elements')
        if '<script' not in content: errors.append('HTML lacks a presentation navigation runtime')
        bonuses = manifest.get('bonuses',[])
        if not isinstance(bonuses,list) or len(bonuses) < 6: errors.append('At least six individual companion PDFs required')
        hashes, text_hashes, paths = set(), set(), set()
        for value in bonuses:
            p = path(value)
            if p in paths or p == pdf: errors.append(f'Repeated PDF or main deck counted as bonus: {value}')
            paths.add(p)
            h = hashlib.sha256(p.read_bytes()).hexdigest()
            if h in hashes: errors.append(f'Byte-identical bonus PDF: {value}')
            hashes.add(h)
            with fitz.open(p) as doc:
                txt = ' '.join(page.get_text() for page in doc)
                if len(doc) < 1: errors.append(f'No pages in bonus: {value}')
            normalized = re.sub(r'\s+',' ',txt).strip()
            if len(normalized) < 180: errors.append(f'Bonus has too little extractable text: {value}')
            th = hashlib.sha256(normalized.encode()).hexdigest()
            if th in text_hashes: errors.append(f'Text-identical bonus PDF: {value}')
            text_hashes.add(th)
        facts['bonus_count'] = len(bonuses)
        url = manifest.get('published_url')
        if require_url and not url: errors.append('Verified publication URL required')
        if url:
            u = urlparse(url)
            if u.scheme != 'https' or not u.netloc or u.username or u.password:
                errors.append('Published URL must be a credential-free HTTPS URL')
        facts['published_url_recorded'] = bool(url)
    except Exception as exc:
        errors.append(str(exc))
    return {'ok':not errors,'facts':facts,'errors':errors,
            'manual_checks_required':['visual layout and RTL','six substantive and distinct companion outcomes','animated HTML browser interaction','published URL and recipient access']}

if __name__ == '__main__':
    parser=argparse.ArgumentParser(description=__doc__,formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('--root',type=Path,required=True)
    parser.add_argument('--manifest',type=Path,required=True)
    parser.add_argument('--require-url',action='store_true')
    args=parser.parse_args()
    result=verify(args.root,json.loads(args.manifest.read_text(encoding='utf-8')),args.require_url)
    print(json.dumps(result,ensure_ascii=False,indent=2))
    sys.exit(0 if result['ok'] else 1)
