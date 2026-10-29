"""Read-only OOXML inspection; not a visual or factual validation."""
import argparse
import json
import re
import zipfile
from xml.etree import ElementTree as ET

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
W = '{' + NS['w'] + '}'

def enabled(node):
    return node is not None and node.get(W+'val', 'true').lower() not in ('0', 'false', 'off')

def inspect(path):
    with zipfile.ZipFile(path) as z:
        doc = ET.fromstring(z.read('word/document.xml'))
        styles = ET.fromstring(z.read('word/styles.xml'))
        media = [n for n in z.namelist() if n.startswith('word/media/')]
    headings, spaces, breaks = [], [], []
    for p in doc.findall('.//w:p', NS):
        text = ''.join(t.text or '' for t in p.findall('.//w:t', NS))
        style = p.find('w:pPr/w:pStyle', NS)
        sid = style.get(W+'val') if style is not None else None
        if re.match(r'^\d+(?:\.\d+)*\s+\S', text) or (sid and 'heading' in sid.lower()) or p.find('w:pPr/w:outlineLvl', NS) is not None:
            headings.append({'text':text, 'style_id':sid})
        if enabled(p.find('w:pPr/w:pageBreakBefore', NS)):
            breaks.append(text[:100])
        if re.search(r'[\u4e00-\u9fff] +[A-Za-z0-9]|[A-Za-z0-9] +[\u4e00-\u9fff]', text):
            spaces.append(text[:120])
    info=[]
    for st in styles.findall('w:style', NS):
        name=st.find('w:name',NS)
        if name is None or name.get(W+'val','').lower() not in ('normal','heading 1','heading 2','caption','toc 1'): continue
        fonts=st.find('w:rPr/w:rFonts',NS)
        size=st.find('w:rPr/w:sz',NS)
        info.append({'name':name.get(W+'val'),'fonts':{} if fonts is None else {k.split('}')[-1]:v for k,v in fonts.attrib.items()},'half_points':None if size is None else size.get(W+'val'),'page_break_before':enabled(st.find('w:pPr/w:pageBreakBefore',NS))})
    return {'file':str(path),'headings_candidates':headings,'styles':info,'media_files':len(media),'tables':len(doc.findall('.//w:tbl',NS)),'manual_page_breaks':len(doc.findall('.//w:br[@w:type="page"]',NS)),'direct_page_break_before':breaks,'possible_unnecessary_spaces':spaces,'note':'Read-only hints. Fonts may inherit from base styles/docDefaults. Spaces require review. Rendered layout, TOC page numbers and experimental claims are not validated.'}

if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('docx')
    args=parser.parse_args()
    print(json.dumps(inspect(args.docx),ensure_ascii=True,indent=2))
