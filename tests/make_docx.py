"""Сборка минимального .docx для тестов (только document.xml и, при необходимости, numbering.xml)."""
import zipfile, io
from xml.sax.saxutils import escape

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

def para(text, auto_num=False, center=False, ind=None, lrpb=False):
    ppr = ''
    if auto_num:
        ppr += '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
    if center:
        ppr += '<w:jc w:val="center"/>'
    if ind is not None:
        ppr += '<w:ind w:left="0" w:firstLine="%d"/>' % ind
    pp = '<w:pPr>%s</w:pPr>' % ppr if ppr else ''
    br = '<w:r><w:lastRenderedPageBreak/></w:r>' if lrpb else ''
    return '<w:p>%s%s<w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (pp, br, escape(text))

NUMBERING = ('<w:numbering xmlns:w="%s"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>'
             '<w:numFmt w:val="decimal"/><w:lvlText w:val="%%1."/><w:pPr><w:ind w:left="709" w:firstLine="0"/></w:pPr></w:lvl></w:abstractNum>'
             '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>') % W

def build(paras, numbering=True):
    body = ''.join(paras)
    doc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="%s"><w:body>%s</w:body></w:document>' % (W, body)
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('word/document.xml', doc)
        if numbering:
            z.writestr('word/numbering.xml', NUMBERING)
    bio.seek(0)
    return bio
