"""Лист самопроверки ВКР (.docx): проверка решений студента и сборка файла.

Только стандартная библиотека: файл .docx собирается как zip из вручную написанной разметки WordprocessingML.
В лист не попадают имя файла, титульный лист и метаданные об авторе. Во фрагментах маскируются e-mail,
телефоны и длинные числовые последовательности; ФИО в тексте надёжно найти нельзя, поэтому студент может
отключить фрагменты целиком (include_quotes = False).
"""
import base64
import io
import json
import re
import zipfile
from xml.sax.saxutils import escape

MIN_CLAUSE, MAX_CLAUSE = 3, 120      # «Пункт ГОСТ»: не короче 3 знаков, есть цифра
MIN_REASON, MAX_REASON = 15, 1000    # «Почему оставляете»
LEVEL = {'fix': 'Исправьте', 'check': 'Проверьте'}

_BAD_XML = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff￾￿]')
_EMAIL = re.compile(r'[\w.+-]+@[\w-]+(?:\.[\w-]+)+')
_PHONE = re.compile(r'(?<!\w)(?:\+7|8)[\s\-]*\(?\d{3}\)?[\s\-]*\d{3}[\s\-]*\d{2}[\s\-]*\d{2}(?!\d)')
_LONGNUM = re.compile(r'(?<!\d)\d(?:[\s\-]?\d){9,}(?!\d)')
_QUOTED = re.compile(r'«[^»]*»')


def xml_text(s):
    """Строка без символов, недопустимых в XML 1.0 (Word отказывается открывать файл с ними)."""
    return _BAD_XML.sub('', '' if s is None else str(s)).replace('\r\n', '\n').replace('\r', '\n')


def mask(s):
    s = xml_text(s)
    s = _EMAIL.sub('[e-mail]', s)
    s = _PHONE.sub('[телефон]', s)
    return _LONGNUM.sub('[число]', s)


def _clip(s, n):
    s = ' '.join(xml_text(s).split())
    return s if len(s) <= n else s[:n - 1].rstrip() + '…'


def clause_ok(clause):
    c = ' '.join(xml_text(clause).split())
    return MIN_CLAUSE <= len(c) <= MAX_CLAUSE and bool(re.search(r'\d', c))


def reason_ok(reason):
    r = ' '.join(xml_text(reason).split())
    return MIN_REASON <= len(r) <= MAX_REASON


def validate(payload):
    """Возвращает {'undecided': [id], 'bad_keep': [id]}; лист формируется, когда оба списка пусты."""
    dec = payload.get('decisions') or {}
    undecided, bad = [], []
    for g in payload.get('groups') or []:
        for it in g.get('items') or []:
            d = dec.get(it.get('id')) or {}
            kind = d.get('d')
            if kind == 'fixed':
                continue
            if kind == 'keep':
                if not (clause_ok(d.get('clause')) and reason_ok(d.get('reason'))):
                    bad.append(it['id'])
            else:
                undecided.append(it['id'])
    return {'undecided': undecided, 'bad_keep': bad}


# --- разметка WordprocessingML ---------------------------------------------------------------
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
FONT = 'Times New Roman'


def _rpr(b=False, i=False, sz=None, color=None):
    p = ''
    if b:
        p += '<w:b/><w:bCs/>'
    if i:
        p += '<w:i/><w:iCs/>'
    if color:
        p += '<w:color w:val="%s"/>' % color
    if sz:
        p += '<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (sz, sz)
    return '<w:rPr>%s</w:rPr>' % p if p else ''


def run(text, **kw):
    parts = xml_text(text).split('\n')
    out = []
    for k, part in enumerate(parts):
        if k:
            out.append('<w:r>%s<w:br/></w:r>' % _rpr(**kw))
        if part:
            out.append('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (_rpr(**kw), escape(part)))
    return ''.join(out)


def para(content='', style=None, jc=None, before=None, after=None, keep_next=False):
    ppr = ''
    if style:
        ppr += '<w:pStyle w:val="%s"/>' % style
    if keep_next:
        ppr += '<w:keepNext/>'
    if before is not None or after is not None:
        ppr += '<w:spacing%s%s/>' % ((' w:before="%d"' % before) if before is not None else '',
                                      (' w:after="%d"' % after) if after is not None else '')
    if jc:
        ppr += '<w:jc w:val="%s"/>' % jc
    return '<w:p>%s%s</w:p>' % ('<w:pPr>%s</w:pPr>' % ppr if ppr else '', content)


def cell(width, paras, fill=None):
    shd = '<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % fill if fill else ''
    return '<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>%s</w:tcPr>%s</w:tc>' % (width, shd, paras or para())


def table(grid, header, rows):
    total = sum(grid)
    brd = ''.join('<w:%s w:val="single" w:sz="4" w:space="0" w:color="808080"/>' % n
                  for n in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
    x = ['<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/><w:tblBorders>%s</w:tblBorders><w:tblLayout w:type="fixed"/>'
         '<w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="80" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/>'
         '<w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>%s</w:tblGrid>'
         % (total, brd, ''.join('<w:gridCol w:w="%d"/>' % g for g in grid))]
    x.append('<w:tr><w:trPr><w:cantSplit/><w:tblHeader/></w:trPr>%s</w:tr>'
             % ''.join(cell(w, para(run(h, b=True, sz=20)), 'E7EAEE') for w, h in zip(grid, header)))
    for r in rows:
        x.append('<w:tr><w:trPr><w:cantSplit/></w:trPr>%s</w:tr>' % ''.join(cell(w, c) for w, c in zip(grid, r)))
    x.append('</w:tbl>')
    return ''.join(x)


def tpara(*runs_xml):
    return para(''.join(runs_xml))


# --- содержимое листа -----------------------------------------------------------------------------
def _where(it, offset, pages_known, quotes):
    parts = []
    if pages_known and isinstance(it.get('page'), int):
        parts.append('стр. %d' % (it['page'] + offset))
    head = _clip(mask(it.get('head')), 120)
    if head:
        parts.append('раздел «%s»' % head)
    return ' · '.join(parts) or 'место не определено'


def _finding_cell(g, it, quotes):
    msg = mask(it.get('msg'))
    if not quotes:
        msg = _QUOTED.sub('«…»', msg)
    ps = [tpara(run('[%s] %s (%s)' % (LEVEL.get(g.get('level'), ''), _clip(g.get('title'), 120), g.get('rule', '')), b=True, sz=20)),
          tpara(run(_clip(msg, 400), sz=20))]
    if quotes:
        q = _clip(mask(it.get('quote')), 120)
        if q:
            ps.append(tpara(run('«%s»' % q, i=True, sz=20)))
        for ex in (it.get('examples') or [])[:3]:
            e = _clip(mask(ex.get('quote')), 100)
            if e:
                ps.append(tpara(run('Пример: «%s»' % e, i=True, sz=18, color='555555')))
    return ''.join(ps)


def build_document(payload):
    groups = payload.get('groups') or []
    dec = payload.get('decisions') or {}
    offset = max(0, int(payload.get('offset') or 0))
    pages_known = bool(payload.get('pages_known'))
    quotes = bool(payload.get('include_quotes', True))
    keep, fixed = [], []
    levels = {'fix': [0, 0, 0], 'check': [0, 0, 0]}      # всего, исправлено, оставлено
    for g in groups:
        for it in g.get('items') or []:
            d = dec.get(it['id']) or {}
            lv = levels.setdefault(g.get('level'), [0, 0, 0])
            lv[0] += 1
            if d.get('d') == 'keep':
                lv[2] += 1
                keep.append((g, it, d))
            else:
                lv[1] += 1
                fixed.append((g, it, d))
    total = len(keep) + len(fixed)

    body = [para(run('Лист самопроверки оформления ВКР'), style='Title')]
    body.append(para(run('Сформирован автоматически сервисом самопроверки, версия %s. Дата проверки: %s. '
                         'Лист не содержит титульного листа, ФИО и других персональных данных.'
                         % (xml_text(payload.get('version', '')), xml_text(payload.get('date_ru', ''))), sz=20), after=120))
    body.append(para(run('Итог', b=True), keep_next=True, before=60, after=40))
    body.append(para(run('Всего замечаний: %d. Оставлено без изменений: %d. Исправлено: %d.' % (total, len(keep), len(fixed)))))
    for key in ('fix', 'check'):
        t = levels.get(key, [0, 0, 0])
        body.append(para(run('«%s»: всего %d, исправлено %d, оставлено %d.' % (LEVEL[key], t[0], t[1], t[2]))))
    if not pages_known:
        body.append(para(run('В файле нет меток страниц Word (он сохранён не в Word), поэтому вместо страниц указаны разделы и фрагменты.', sz=20)))
    elif offset:
        body.append(para(run('Номера страниц указаны с учётом титульного листа и задания: сдвиг %d стр.' % offset, sz=20)))
    sc = [_clip(c, 80) for c in (payload.get('self_checks') or [])]
    if sc:
        body.append(para(run('Сервис не проверяет (это делает нормоконтролёр): %s.' % '; '.join(sc), sz=20), after=120))

    # 1. Оставлено без изменений — первым
    body.append(para(run('1. Оставлено без изменений (%d)' % len(keep)), style='Heading1'))
    if keep:
        grid = [500, 2500, 5100, 2300, 4738]
        rows = []
        for n, (g, it, d) in enumerate(keep, 1):
            reason = _clip(mask(d.get('reason')), MAX_REASON)
            clause = _clip(mask(d.get('clause')), MAX_CLAUSE)
            rows.append([para(run(str(n), sz=20), jc='center'),
                         para(run(_where(it, offset, pages_known, quotes), sz=20)),
                         _finding_cell(g, it, quotes),
                         para(run(_clip(g.get('source'), 200), sz=20)),
                         tpara(run('Пункт: ', b=True, sz=20), run(clause, sz=20)) + tpara(run('Причина: ', b=True, sz=20), run(reason, sz=20))])
        body.append(table(grid, ['№', 'Где', 'Замечание сервиса', 'Требование', 'Решение студента: пункт и причина'], rows))
    else:
        body.append(para(run('Замечаний, оставленных без изменений, нет.')))

    # 2. Исправлено
    body.append(para(run('2. Исправлено студентом (%d)' % len(fixed)), style='Heading1'))
    if fixed:
        grid = [500, 2500, 7000, 3438, 1700]
        rows = []
        for n, (g, it, d) in enumerate(fixed, 1):
            rows.append([para(run(str(n), sz=20), jc='center'),
                         para(run(_where(it, offset, pages_known, quotes), sz=20)),
                         _finding_cell(g, it, quotes),
                         para(run(_clip(g.get('source'), 200), sz=20)),
                         para(run('Исправил(а)', sz=20))])
        body.append(table(grid, ['№', 'Где', 'Замечание сервиса', 'Требование', 'Решение'], rows))
    else:
        body.append(para(run('Исправленных замечаний нет.')))

    sect = ('<w:sectPr><w:footerReference w:type="default" r:id="rId3"/><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
            '<w:pgMar w:top="850" w:right="850" w:bottom="850" w:left="850" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="%s" xmlns:r="%s"><w:body>%s%s</w:body></w:document>'
            % (W_NS, R_NS, ''.join(body), sect))


STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="%s">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:eastAsia="{f}" w:cs="{f}"/>'
          '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="ru-RU" w:eastAsia="ru-RU" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="60" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
          '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>'
          '<w:pPr><w:spacing w:before="0" w:after="120"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>'
          '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>'
          '<w:pPr><w:keepNext/><w:spacing w:before="280" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr>'
          '<w:rPr><w:b/><w:bCs/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>'
          '</w:styles>') % W_NS
STYLES = STYLES.replace('{f}', FONT)

SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="%s"><w:defaultTabStop w:val="709"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
            '</w:settings>') % W_NS

FOOTER = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="%s"><w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
          '<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">Лист самопроверки ВКР · страница </w:t></w:r>'
          '<w:fldSimple w:instr=" PAGE "><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>'
          '<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve"> из </w:t></w:r>'
          '<w:fldSimple w:instr=" NUMPAGES "><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>') % W_NS

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>')

RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>')

DOC_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>')

APP = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
       '<Application>Сервис самопроверки ВКР</Application></Properties>')


def _core(date_iso):
    created = ''
    if re.fullmatch(r'\d{4}-\d{2}-\d{2}', date_iso or ''):
        created = '<dcterms:created xsi:type="dcterms:W3CDTF">%sT00:00:00Z</dcterms:created>' % date_iso
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties '
            'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" '
            'xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Лист самопроверки оформления ВКР</dc:title>%s</cp:coreProperties>' % created)


def build_docx(payload):
    parts = [('[Content_Types].xml', CT), ('_rels/.rels', RELS), ('word/document.xml', build_document(payload)),
             ('word/_rels/document.xml.rels', DOC_RELS), ('word/styles.xml', STYLES), ('word/settings.xml', SETTINGS),
             ('word/footer1.xml', FOOTER), ('docProps/core.xml', _core(payload.get('date_iso'))), ('docProps/app.xml', APP)]
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, xml in parts:
            zi = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(zi, xml.encode('utf-8'))
    return bio.getvalue()


def make(payload_json):
    """Точка входа для браузера: JSON с решениями → JSON {ok, b64} или {ok: false, undecided, bad_keep}."""
    try:
        payload = json.loads(payload_json)
        for g in payload.get('groups') or []:           # идентификаторы — строки вида «3-12»
            for it in g.get('items') or []:
                it['id'] = str(it.get('id'))
        problems = validate(payload)
        if problems['undecided'] or problems['bad_keep']:
            return json.dumps(dict(ok=False, **problems), ensure_ascii=False)
        return json.dumps(dict(ok=True, b64=base64.b64encode(build_docx(payload)).decode('ascii')))
    except Exception as e:
        return json.dumps(dict(ok=False, error='Не удалось сформировать лист: %s' % (e,)), ensure_ascii=False)
