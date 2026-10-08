"""Render authored semantic HTML companion documents to separate PDFs with PyMuPDF Story.

Usage: python render_companion_pdfs.py --input docs --output output/bonuses
Each *.html becomes one PDF. Include lang/dir in the input. Supply substantive authored
content; this renderer does not invent the six documents. Local assets resolve under --input.
Remote resources are not fetched. Review every rendered PDF, especially RTL and mixed scripts.
"""
import argparse
import re
import html as html_module
from html.parser import HTMLParser
from pathlib import Path
import fitz

CSS = '''body { font-family: sans-serif; font-size:11pt; line-height:1.5; color:#18252b; }
h1 { font-size:25pt; color:#18252b; margin-bottom:18pt; }
h2 { font-size:17pt; color:#a83224; margin-top:20pt; }
h3 { font-size:13pt; } p,li { margin-bottom:8pt; }
table { border-collapse:collapse; width:100%; } th,td { padding:6pt; border:0.5pt solid #999; }
.rtl { direction:rtl; text-align:right; } .ltr { direction:ltr; }'''

def render(source, output):
    html=source.read_text(encoding='utf-8')
    css = CSS
    if re.search(r'dir\s*=\s*["\']rtl["\']',html,re.I):
        # Story does not consistently inherit RTL alignment from user_css.
        class RTLBlocks(HTMLParser):
            def __init__(self):
                super().__init__(convert_charrefs=False)
                self.parts=[]
                self.lists=[]
            def handle_starttag(self,tag,attrs):
                attrs=dict(attrs)
                if tag in ('ol','ul'):
                    self.lists.append([tag,0])
                    tag='div'
                elif tag == 'li':
                    tag='p'
                    if self.lists:
                        self.lists[-1][1] += 1
                # MuPDF mirrors explicit right alignment under inherited dir=rtl.
                # Keep Unicode shaping, use CSS alignment, and remove only that HTML flag.
                if attrs.get('dir','').lower() == 'rtl': attrs.pop('dir')
                if tag in ('h1','h2','h3','h4','p','li','td','th'):
                    attrs['style']=attrs.get('style','').rstrip(';')+';text-align:right;direction:rtl;' if attrs.get('style') else 'text-align:right;direction:rtl;'
                rendered=''.join(' '+k+'="'+html_module.escape(v or '',quote=True)+'"' for k,v in attrs.items())
                self.parts.append('<'+tag+rendered+'>')
                if tag == 'p' and self.lists:
                    label = str(self.lists[-1][1])+'. ' if self.lists[-1][0]=='ol' else '• '
                    self.parts.append(label)
            def handle_endtag(self,tag):
                if tag in ('ol','ul'):
                    if self.lists:self.lists.pop()
                    tag='div'
                elif tag=='li':tag='p'
                self.parts.append('</'+tag+'>')
            def handle_data(self,data):self.parts.append(data)
            def handle_entityref(self,name):self.parts.append('&'+name+';')
            def handle_charref(self,name):self.parts.append('&#'+name+';')
            def handle_decl(self,decl):self.parts.append('<!'+decl+'>')
        parser=RTLBlocks()
        parser.feed(html)
        html=''.join(parser.parts)
    story=fitz.Story(html=html,user_css=css,archive=fitz.Archive(str(source.parent)))
    media=fitz.paper_rect('a4')
    content=media+(42,42,-42,-42)
    def rectfn(rect_num,filled):
        return media,content,None
    with fitz.DocumentWriter(str(output)) as writer:
        story.write(writer,rectfn)
    with fitz.open(output) as doc:
        if not len(doc) or not any(p.get_text().strip() for p in doc):
            raise ValueError(f'Empty or non-text PDF: {output}')
        return len(doc)

if __name__=='__main__':
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--input',type=Path,required=True)
    ap.add_argument('--output',type=Path,required=True)
    args=ap.parse_args()
    sources=sorted(args.input.glob('*.html'))
    if len(sources)<6:ap.error('Provide at least six distinct authored HTML documents')
    args.output.mkdir(parents=True,exist_ok=True)
    for source in sources:
        out=args.output/(source.stem+'.pdf')
        if out.exists():ap.error(f'Refusing to overwrite existing output: {out}')
        print(f'{out.name}: {render(source,out)} pages')
