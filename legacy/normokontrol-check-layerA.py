"""Проверка слоя A для ВКР (docx). Правила: реестр v0.4, приоритет ГОСТ 7.32-2017.
Выход: словарь {правило: [описание, ...]} для одной работы."""
import zipfile, re, sys, json, glob, os
from lxml import etree

import os
TXS = 0.03   # доля абзацев основного текста с отклонением, при которой правило TX1 срабатывает
PROFILE = os.environ.get('NK_PROFILE', 'metod')   # metod: правое поле 10 мм и номер справа (до нового ГОСТ); gost: 15 мм и по центру
PROFILE_T = os.environ.get('NK_TOC', 'gost')   # gost: содержание по 5.4.1 и 6.13 (влево, сдвиг подразделов); practice: по ширине, без отступов, как требовал нормоконтролёр
PROFILE_H = os.environ.get('NK_HEADINGS', 'gost')   # gost: заголовки разделов по ГОСТ 6.2.3; practice: по центру, как требовал нормоконтролёр
NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
W = '{%s}' % NS['w']
def q(t): return W + t
def attr(el, name):
    return None if el is None else el.get(q(name))

class Doc:
    def __init__(self, path):
        self.z = zipfile.ZipFile(path)
        self.doc = etree.fromstring(self.z.read('word/document.xml'))
        self.styles = {}
        self.defp = None; self.defr = None; self.default_pstyle = None
        if 'word/styles.xml' in self.z.namelist():
            st = etree.fromstring(self.z.read('word/styles.xml'))
            dd = st.find('w:docDefaults', NS)
            if dd is not None:
                self.defp = dd.find('w:pPrDefault/w:pPr', NS)
                self.defr = dd.find('w:rPrDefault/w:rPr', NS)
            for s in st.findall('w:style', NS):
                sid = s.get(q('styleId'))
                self.styles[sid] = dict(
                    type=s.get(q('type')),
                    name=(attr(s.find('w:name', NS), 'val') or ''),
                    based=attr(s.find('w:basedOn', NS), 'val'),
                    ppr=s.find('w:pPr', NS), rpr=s.find('w:rPr', NS))
                if s.get(q('type')) == 'paragraph' and s.get(q('default')) == '1':
                    self.default_pstyle = sid
        self.theme_minor = None
        if 'word/theme/theme1.xml' in self.z.namelist():
            th = etree.fromstring(self.z.read('word/theme/theme1.xml'))
            m = th.find('.//a:minorFont/a:latin', NS)
            if m is not None: self.theme_minor = m.get('typeface')
        self.settings = etree.fromstring(self.z.read('word/settings.xml')) if 'word/settings.xml' in self.z.namelist() else None

    def chain(self, sid):
        seen = set()
        while sid and sid in self.styles and sid not in seen:
            seen.add(sid); yield self.styles[sid]; sid = self.styles[sid]['based']

    def pprop(self, p, getter):
        ppr = p.find('w:pPr', NS)
        v = getter(ppr) if ppr is not None else None
        if v is not None: return v
        sid = attr(ppr.find('w:pStyle', NS), 'val') if ppr is not None and ppr.find('w:pStyle', NS) is not None else self.default_pstyle
        for s in self.chain(sid):
            if s['ppr'] is not None:
                v = getter(s['ppr'])
                if v is not None: return v
        if self.defp is not None: return getter(self.defp)
        return None

    def rprop(self, r, p, getter):
        rpr = r.find('w:rPr', NS)
        if rpr is not None:
            v = getter(rpr)
            if v is not None: return v
            rs = rpr.find('w:rStyle', NS)
            if rs is not None:
                for s in self.chain(attr(rs, 'val')):
                    if s['rpr'] is not None:
                        v = getter(s['rpr'])
                        if v is not None: return v
        ppr = p.find('w:pPr', NS)
        sid = attr(ppr.find('w:pStyle', NS), 'val') if ppr is not None and ppr.find('w:pStyle', NS) is not None else self.default_pstyle
        for s in self.chain(sid):
            if s['rpr'] is not None:
                v = getter(s['rpr'])
                if v is not None: return v
        if self.defr is not None: return getter(self.defr)
        return None

    def style_name(self, p):
        ppr = p.find('w:pPr', NS)
        sid = attr(ppr.find('w:pStyle', NS), 'val') if ppr is not None and ppr.find('w:pStyle', NS) is not None else self.default_pstyle
        return (self.styles.get(sid, {}).get('name') or '').lower()

# --- getters -------------------------------------------------------------
def g_jc(ppr): return attr(ppr.find('w:jc', NS), 'val') if ppr.find('w:jc', NS) is not None else None
def g_first(ppr):
    i = ppr.find('w:ind', NS)
    if i is None: return None
    if i.get(q('firstLineChars')): return 'chars'
    if i.get(q('firstLine')) is not None: return int(i.get(q('firstLine')))
    if i.get(q('hanging')) is not None: return -int(i.get(q('hanging')))
    return None
def g_sp(name):
    def f(ppr):
        s = ppr.find('w:spacing', NS)
        if s is None: return None
        return s.get(q(name))
    return f
def g_flag(tag):
    def f(ppr):
        e = ppr.find('w:' + tag, NS)
        if e is None: return None
        return attr(e, 'val') not in ('0', 'false', 'off')
    return f
def g_font(rpr):
    f = rpr.find('w:rFonts', NS)
    if f is None: return None
    if f.get(q('ascii')): return f.get(q('ascii'))
    if f.get(q('hAnsi')): return f.get(q('hAnsi'))
    if f.get(q('asciiTheme')): return 'THEME:' + f.get(q('asciiTheme'))
    return None
def g_sz(rpr):
    e = rpr.find('w:sz', NS)
    return int(attr(e, 'val')) if e is not None else None
def g_b(rpr):
    e = rpr.find('w:b', NS)
    if e is None: return None
    return attr(e, 'val') not in ('0', 'false', 'off')

# --- извлечение текста -------------------------------------------------------
def ptext(p):
    out = []
    for e in p.iter():
        if e.tag == q('t'): out.append(e.text or '')
        elif e.tag == q('tab'): out.append('\t')
        elif e.tag == q('br') and e.get(q('type')) in (None, 'textWrapping'): out.append('\n')
    return ''.join(out)

RE_STRUCT = re.compile(r'^(ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|СПИСОК\s+(ИСПОЛЬЗОВАННЫХ\s+)?(ИСТОЧНИКОВ|ЛИТЕРАТУРЫ)|СОДЕРЖАНИЕ|ОГЛАВЛЕНИЕ|РЕФЕРАТ|ABSTRACT|ПРИЛОЖЕНИЯ?(\s+[А-ЯA-Z0-9])?)\b', re.I)
RE_CHAP = re.compile(r'^(ГЛАВА|Глава)\s+(\d+|[IVXLC]+)\b')
RE_SUB = re.compile(r'^\d+(\.\d+)+\.?\s+\S')
RE_BIB = re.compile(r'^(СПИСОК\s+(ИСПОЛЬЗ\w+\s+)?(ИСТОЧНИК\w*|ЛИТЕРАТУР\w*)|БИБЛИОГРАФИЧЕСКИЙ\s+СПИСОК|ИСПОЛЬЗ\w+\s+(ИСТОЧНИК\w*|ЛИТЕРАТУР\w*)|ЛИТЕРАТУРА|ИСТОЧНИКИ|REFERENCES|BIBLIOGRAPHY)')
RE_VYV = re.compile(r'^(ВЫВОД[ЫА]?|Вывод[ыа]?)\s+(ПО|по)\s+(ГЛАВ[ЕА]|глав[еа])', re.I)

# --- перечисления: ГОСТ 7.32-2017, п. 6.4.6 (LS1) и п. 6.16 (LS2) ---------------------
DASHES = set('-–—−‒―⎯‐‑')            # допустимые «тире» (в том числе из Symbol: F02D)
RU_LET = 'абвгдеёжзийклмнопрстуфхцчшщъыьэюя'
RU_BAD = set('ёзйочъыь')                              # п. 6.4.6: буквы, которые не используются
WORD_RU = 'абвгдежзиклмнопрстуфхцчшщэюя'             # порядок numFmt=russianLower в Word
LS_IND = 709; LS_TOL = int(os.environ.get('NK_LS_TOL', 60))   # абзацный отступ 1,25 см, допуск ~1 мм

def load_numbering(D):
    nums, absn = {}, {}
    if 'word/numbering.xml' not in D.z.namelist(): return nums, absn
    nx = etree.fromstring(D.z.read('word/numbering.xml'))
    for a in nx.findall('w:abstractNum', NS):
        lv = {}
        for l in a.findall('w:lvl', NS):
            lv[int(l.get(q('ilvl')))] = dict(fmt=attr(l.find('w:numFmt', NS), 'val'), txt=attr(l.find('w:lvlText', NS), 'val'),
                                              ind=l.find('w:pPr/w:ind', NS), start=attr(l.find('w:start', NS), 'val'))
        absn[a.get(q('abstractNumId'))] = lv
    for n in nx.findall('w:num', NS):
        nums[n.get(q('numId'))] = dict(abs=attr(n.find('w:abstractNumId', NS), 'val'),
                                       ov={int(o.get(q('ilvl'))): o for o in n.findall('w:lvlOverride', NS)})
    return nums, absn

def _ind_vals(ind):
    """w:ind → (left, first), first>0 — красная строка, <0 — выступ; None — не задано."""
    if ind is None: return None, None
    def iv(n):
        for k in (n, {'left': 'start'}.get(n)):
            if k and ind.get(q(k)) is not None: return int(float(ind.get(q(k))))
        return None
    left = iv('left')
    if ind.get(q('firstLineChars')): first = int(int(ind.get(q('firstLineChars'))) * 2.8)   # 1/100 знака при 14 пт ≈ 2,8 твипа
    elif ind.get(q('firstLine')) is not None: first = int(ind.get(q('firstLine')))
    elif ind.get(q('hanging')) is not None: first = -int(ind.get(q('hanging')))
    else: first = None
    return left, first

def para_numpr(D, p):
    def g(ppr):
        n = ppr.find('w:numPr', NS)
        if n is None: return None
        i = n.find('w:ilvl', NS); nid = n.find('w:numId', NS)
        return (int(attr(i, 'val')) if i is not None else 0, attr(nid, 'val') if nid is not None else None)
    return D.pprop(p, g)

def eff_marker_pos(D, p, lvl_ind):
    """Положение маркера от левого поля, твипы: стиль < уровень нумерации < прямое форматирование."""
    left = first = None
    ppr = p.find('w:pPr', NS)
    sid = attr(ppr.find('w:pStyle', NS), 'val') if ppr is not None and ppr.find('w:pStyle', NS) is not None else D.default_pstyle
    layers = [s['ppr'].find('w:ind', NS) for s in reversed(list(D.chain(sid))) if s['ppr'] is not None]
    layers.append(lvl_ind)
    if ppr is not None: layers.append(ppr.find('w:ind', NS))
    for ind in layers:
        l, f = _ind_vals(ind)
        if l is not None: left = l
        if f is not None: first = f
    left = left or 0; first = first or 0
    return left + first if first > 0 else left + first if first < 0 else left

def list_marker(D, p, text, nums, absn):
    """Описание маркера абзаца или None. kind: dash | bullet | digit_p (1)) | digit_d (1.) | ru_p | ru_d | lat | multi."""
    v = para_numpr(D, p)
    if v and v[1] not in (None, '0'):
        n = nums.get(v[1]); lv = absn.get(n['abs'], {}) if n else {}
        l = lv.get(v[0])
        if l is None: return None
        fmt, txt = l['fmt'], l['txt'] or ''
        ov = n['ov'].get(v[0]) if n else None
        if ov is not None and ov.find('w:lvl', NS) is not None:
            lo = ov.find('w:lvl', NS)
            fmt = attr(lo.find('w:numFmt', NS), 'val') or fmt; txt = attr(lo.find('w:lvlText', NS), 'val') or txt
        pos = eff_marker_pos(D, p, l['ind'])
        m = dict(auto=True, numid=v[1], lvl=v[0], pos=pos, txt=txt, fmt=fmt, start=int(l['start'] or 1))
        if fmt == 'bullet':
            m['kind'] = 'dash' if txt and txt[0] in DASHES else 'bullet'; m['ch'] = txt[:1]
        elif txt.count('%') >= 2: m['kind'] = 'multi'
        else:
            paren = ')' in txt; dot = '.' in txt
            m['kind'] = {'decimal': 'digit', 'russianLower': 'ru', 'lowerLetter': 'lat', 'upperLetter': 'lat', 'lowerRoman': 'lat', 'upperRoman': 'lat',
                         'decimalZero': 'digit'}.get(fmt, 'other')
            if m['kind'] in ('digit', 'ru', 'lat'): m['kind'] += '_p' if paren else '_d' if dot else '_n'
        return m
    mm = re.match(r'^\s*([-–—−•·●▪■◦*])\s+(\S)', text)
    if mm:
        ch = mm.group(1)
        return dict(auto=False, kind='dash' if ch in DASHES else 'bullet', ch=ch, pos=eff_marker_pos(D, p, None))
    mm = re.match(r'^\s*(\d{1,2})([).])\s+\S', text)
    if mm: return dict(auto=False, kind='digit_p' if mm.group(2) == ')' else 'digit_d', ord=int(mm.group(1)), pos=eff_marker_pos(D, p, None))
    mm = re.match(r'^\s*([а-яёa-z])([).])\s+\S', text)
    if mm:
        ch = mm.group(1); ru = ch in RU_LET
        return dict(auto=False, kind=('ru' if ru else 'lat') + ('_p' if mm.group(2) == ')' else '_d'), ch=ch, ord=(RU_LET.index(ch) + 1 if ru else ord(ch) - 96), pos=eff_marker_pos(D, p, None))
    return None

def ls_check(runs, add, bib=False):
    """runs — список последовательных абзацев-перечислений [(marker, text)]. Нарушения п. 6.4.6 (LS1) / п. 6.16 (LS2)."""
    if not runs: return
    k = runs[0][0]['kind']; head = runs[0][1][:30]
    if not bib:
        if k == 'bullet': add('LS1', 'перечисление маркером-точкой вместо тире (6.4.6): «%s»' % head)
        elif k in ('digit_d', 'ru_d', 'lat_d', 'digit_n', 'ru_n', 'lat_n'): add('LS1n', 'нумерация без скобки (%s): «%s»' % ('1.' if k[:5] == 'digit' else 'а.', head))
        elif k.startswith('lat'): add('LS1', 'латинские буквы вместо русских (6.4.6): «%s»' % head)
        elif k == 'other': add('LS1', 'нестандартный формат нумерации: «%s»' % head)
        if k in ('ru_p', 'ru_d'):
            letters = [m.get('ch') for m, t in runs if m.get('ch')]
            if not letters and runs[0][0].get('auto'):
                st = runs[0][0].get('start', 1) - 1
                letters = [WORD_RU[(st + i) % len(WORD_RU)] for i in range(len(runs))]
            if any(c in RU_BAD for c in letters): add('LS1', 'буква вне допустимого набора (ё з й о ч ъ ы ь): «%s»' % head)
            if letters and letters[0] != 'а': add('LS1', 'перечисление буквами начато не с «а»: «%s»' % head)
        if k == 'digit_p':
            ords = [m.get('ord') for m, t in runs if m.get('ord')]
            if ords and ords[0] != 1: add('LS1', 'нумерация перечисления начата не с 1: «%s»' % head)
        if k in ('digit_p', 'ru_p', 'dash', 'digit_d', 'ru_d', 'bullet'):
            bad = [m for m, t in runs if abs(m['pos'] - LS_IND) > LS_TOL]
            if bad: add('LS1i', 'маркер не с абзацного отступа 1,25 см: %.1f см, %d из %d (6.4.6): «%s»' % (bad[0]['pos'] / 567.0, len(bad), len(runs), head))
    else:
        # 6.16: источники — арабские цифры с точкой, с абзацного отступа
        if k in ('digit_p',): add('LS2', 'источники пронумерованы со скобкой, нужна точка (6.16): «%s»' % head)
        bad = [m for m, t in runs if abs(m['pos'] - LS_IND) > LS_TOL]
        if k.startswith('digit') and bad: add('LS2', 'список источников не с абзацного отступа 1,25 см: %.1f см, %d из %d (6.16): «%s»' % (bad[0]['pos'] / 567.0, len(bad), len(runs), head))

# --- содержание: ГОСТ 7.32-2017, п. 5.4.1 и 6.13 -----------------------------------------
def eff_tabs(D, p):
    """Позиции табуляции абзаца: стиль < прямое форматирование; val=clear снимает. {pos: (val, leader)}"""
    tabs = {}
    ppr = p.find('w:pPr', NS)
    sid = attr(ppr.find('w:pStyle', NS), 'val') if ppr is not None and ppr.find('w:pStyle', NS) is not None else D.default_pstyle
    layers = [s['ppr'] for s in reversed(list(D.chain(sid))) if s['ppr'] is not None]
    if ppr is not None: layers.append(ppr)
    for l in layers:
        for t in l.findall('w:tabs/w:tab', NS):
            pos = int(float(t.get(q('pos')) or 0)); v = t.get(q('val'))
            if v == 'clear': tabs.pop(pos, None)
            else: tabs[pos] = (v, t.get(q('leader')) or 'none')
    return tabs

def eff_left(D, p):
    left = 0; first = 0
    ppr = p.find('w:pPr', NS)
    sid = attr(ppr.find('w:pStyle', NS), 'val') if ppr is not None and ppr.find('w:pStyle', NS) is not None else D.default_pstyle
    layers = [s['ppr'].find('w:ind', NS) for s in reversed(list(D.chain(sid))) if s['ppr'] is not None]
    if ppr is not None: layers.append(ppr.find('w:ind', NS))
    for ind in layers:
        l, f = _ind_vals(ind)
        if l is not None: left = l
        if f is not None: first = f
    return left, first

RE_TOC_TAIL = re.compile(r'[\s.…··_\-–—]*\t?\s*\d+\s*$')
def toc_level(text, sname):
    t = RE_TOC_TAIL.sub('', text).strip()
    if re.match(r'^\d+\.\d+\.\d+', t): return 3
    if re.match(r'^\d+\.\d+', t): return 2
    if re.match(r'^(ГЛАВА|Глава|\d+[.\s])', t) or RE_STRUCT.match(t) or re.match(r'^(ВЫВОД|Вывод)', t): return 1
    m = re.search(r'(?:toc|оглавление|содержание)\s*(\d)', sname)
    return min(max(int(m.group(1)), 1), 3) if m else 2

def toc_check(D, toc_paras, textw, add):
    """6.13: запись — отдельный абзац, выровнена влево; номер страницы — по правому краю поля через отточие.
    5.4.1: подразделы сдвинуты относительно разделов (на два знака), пункты — на четыре."""
    ent = []
    for p in toc_paras:
        t = ptext(p).strip()
        if not t or t.upper() in ('СОДЕРЖАНИЕ', 'ОГЛАВЛЕНИЕ'): continue
        ent.append((p, t))
    if len(ent) < 3: return {}
    stat = dict(n=len(ent), notleft=0, notab=0, typed=0, lv={1: [], 2: [], 3: []}, merged=0)
    for p, t in ent:
        jc = D.pprop(p, g_jc)
        if jc not in (None, 'left', 'start'): stat['notleft'] += 1
        tabs = eff_tabs(D, p)
        ok_tab = any(v == 'right' and ld in ('dot', 'middleDot') and (textw is None or abs(pos - textw) <= 340) for pos, (v, ld) in tabs.items())
        typed = bool(re.search(r'(\.{4,}|…{2,}|[.…]{4,})\s*\d+\s*$', t))
        if typed: stat['typed'] += 1
        if not ok_tab and not typed: stat['notab'] += 1
        sname = D.style_name(p)
        stat['lv'][toc_level(t, sname)].append(eff_left(D, p)[0])
        if re.search(r'\d+\s*\n\s*(ВВЕДЕНИЕ|ГЛАВА|Глава|\d+[.\s]|ЗАКЛЮЧЕНИЕ|СПИСОК|ПРИЛОЖЕНИЕ)', t): stat['merged'] += 1
    n = stat['n']
    if PROFILE_T == 'gost':
        if stat['notleft'] / n > .3: add('TOC3', 'запись содержания не выровнена влево (6.13): %d из %d' % (stat['notleft'], n))
    else:
        nb = sum(1 for p, t in ent if D.pprop(p, g_jc) not in ('both', 'distribute'))
        if nb / n > .3: add('TOC3', 'содержание не по ширине (практика нормоконтроля): %d из %d' % (nb, n))
    if stat['notab'] / n > .3: add('TOC4', 'номер страницы не по правому краю через отточие (6.13): %d из %d записей' % (stat['notab'], n))
    if stat['merged']: add('TOC3', 'несколько записей содержания в одном абзаце (6.13: каждая запись — отдельный абзац): %d' % stat['merged'])
    def med(a): a = sorted(a); return a[len(a) // 2]
    l1, l2, l3 = stat['lv'][1], stat['lv'][2], stat['lv'][3]
    if PROFILE_T == 'gost':
        if len(l1) >= 2 and len(l2) >= 2 and med(l2) - med(l1) < 100:
            add('TOC5', 'подразделы в содержании не сдвинуты относительно разделов (5.4.1): %.1f и %.1f см' % (med(l2) / 567.0, med(l1) / 567.0))
        if len(l2) >= 2 and len(l3) >= 2 and med(l3) - med(l2) < 100:
            add('TOC5', 'пункты в содержании не сдвинуты относительно подразделов (5.4.1)')
    elif len(l2) >= 2 and med(l2) - med(l1 or [0]) >= 100:
        add('TOC5', 'в содержании есть отступы слева (практика нормоконтроля — без отступов)')
    return stat

def _tkey(t):
    t = RE_TOC_TAIL.sub('', t).strip()
    t = re.sub(r'^(ГЛАВА|Глава)\s+[\dIVXLC]+\s*[.:]?\s*', '', t)
    t = re.sub(r'^\d+(\.\d+)*\.?\s*', '', t)
    return re.sub(r'[^0-9a-zа-яё]', '', t.lower())

def toc_compose(toc_paras, heads, add):
    """5.4.1: в содержании — введение, все разделы и подразделы, заключение, список источников, приложения."""
    if len(toc_paras) < 3: return
    keys = [_tkey(ptext(p).strip()) for p in toc_paras]
    keys = [k for k in keys if k]
    def present(h):
        k = _tkey(h)
        return (not k) or any(k[:25] in x or (x and x in k) for x in keys)
    miss_struct = []
    for pat, name in ((r'^ВВЕДЕНИЕ', 'ВВЕДЕНИЕ'), (r'^ЗАКЛЮЧЕНИЕ', 'ЗАКЛЮЧЕНИЕ')):
        if any(re.match(pat, h.upper()) for h in heads) and not any(k.startswith(name.lower()) for k in keys): miss_struct.append(name)
    if miss_struct: add('TOC6', 'в содержании нет: %s (5.4.1)' % ', '.join(miss_struct))
    main = []
    for h in heads:                      # заголовки основной части: до первого приложения, только нумерованные (в приложениях — задания и т. п.)
        if re.match(r'^ПРИЛОЖЕНИЕ', h.upper()): break
        if re.match(r'^(ГЛАВА|Глава)\s+\S+', h) or re.match(r'^\d+(\.\d+)*\.?\s+\S', h) and not re.match(r'^(Таблица|Рисунок)', h) and not h.rstrip().endswith(('.', ':', '?')) and '?' not in h and not re.search(r'[.!]\s+\S', re.sub(r'^\d+(\.\d+)*\.?\s+', '', h)): main.append(h)   # задания и предложения — не заголовки
    def lvl(h): return 1 if re.match(r'^(ГЛАВА|Глава)', h) or re.match(r'^\d+\.?\s', h) else (2 if re.match(r'^\d+\.\d+\.?\s', h) else 3)
    for names, code, lab in (([h for h in main if lvl(h) <= 2], 'TOC6', 'разделов и подразделов'), ([h for h in main if lvl(h) == 3], 'TOC7', 'пунктов')):
        absent = [h for h in names if not present(h)]
        if absent and (len(absent) / max(len(names), 1) >= .15 or len(absent) >= 3):
            add(code, 'в содержании нет %d из %d заголовков %s (5.4.1): «%s»' % (len(absent), len(names), lab, absent[0][:30]))

def check(path):
    D = Doc(path)
    F = {}   # rule -> list[str]
    def add(rule, msg): F.setdefault(rule, []).append(msg[:90])
    body = D.doc.find('w:body', NS)
    # --- поля, нумерация
    sects = body.findall('.//w:sectPr', NS)
    textw = None
    for i, s in enumerate(sects):
        m = s.find('w:pgMar', NS); sz = s.find('w:pgSz', NS)
        if m is None: continue
        L, R, T, B = (int(m.get(q(k)) or 0) for k in ('left', 'right', 'top', 'bottom'))
        if sz is not None and textw is None: textw = int(sz.get(q('w'))) - L - R
        tol = 60  # ~1 мм
        exp = dict(left=1701, right=(567 if PROFILE == 'metod' else 850), top=1134, bottom=1134)
        bad = [k for k, v in zip(('left', 'right', 'top', 'bottom'), (L, R, T, B)) if abs(v - exp[k]) > tol]
        if 'right' in bad and i == 0: add('PG1r', 'правое поле %.1f мм вместо %d' % (R / 56.7, 10 if PROFILE == 'metod' else 15))
        bad = [b for b in bad if b != 'right']
        if bad and i == 0 or bad and len(sects) == 1:
            add('PG1', 'поля: ' + ', '.join('%s=%.1f мм' % (k, dict(left=L, right=R, top=T, bottom=B)[k] / 56.7) for k in bad))
    # номер страницы
    has_page = False; page_center = None
    for n in D.z.namelist():
        if re.match(r'word/footer\d*\.xml', n):
            ft = etree.fromstring(D.z.read(n))
            for p in ft.iter(q('p')):
                if p.xpath('.//w:instrText[contains(.,"PAGE")]|.//w:fldSimple[contains(@w:instr,"PAGE")]', namespaces=NS):
                    has_page = True
                    j = D.pprop(p, g_jc)
                    want = 'right' if PROFILE == 'metod' else 'center'
                    if j == want: page_center = True
                    elif page_center is None: page_center = False
    if not has_page: add('NP1', 'нет автонумерации страниц в нижнем колонтитуле')
    elif page_center is False: add('N1c', 'номер страницы стоит не там (%s)' % ('справа' if PROFILE == 'metod' else 'по центру'))
    # номер на титульном листе
    try:
        rels = {}
        rx = etree.fromstring(D.z.read('word/_rels/document.xml.rels'))
        for r_ in rx: rels[r_.get('Id')] = r_.get('Target')
        s0 = sects[0]
        refs = {fr.get(q('type')): rels.get(fr.get('{%s}id' % NS['r'])) for fr in s0.findall('w:footerReference', NS)}
        tp = s0.find('w:titlePg', NS) is not None
        tgt = refs.get('first') if tp else refs.get('default')
        if tgt:
            ft = etree.fromstring(D.z.read('word/' + tgt.lstrip('/').replace('word/', '')))
            if ft.xpath('.//w:instrText[contains(.,"PAGE")]|.//w:fldSimple[contains(@w:instr,"PAGE")]', namespaces=NS):
                add('NP1', 'номер страницы на титульном листе')
    except Exception:
        pass
    # hyphenation
    auto_h = False
    if D.settings is not None:
        e = D.settings.find('w:autoHyphenation', NS)
        auto_h = e is not None and attr(e, 'val') not in ('0', 'false', 'off')
    if not auto_h: add('R1', 'автоперенос выключен')
    # --- обход абзацев
    # элементы верхнего уровня: p, tbl, sdt
    items = []
    for ch in body:
        if ch.tag == q('p') or ch.tag == q('tbl'): items.append(ch)
        elif ch.tag == q('sdt'):
            c = ch.find('w:sdtContent', NS)
            toc_sdt = bool(ch.xpath('.//w:docPartGallery[contains(@w:val,"Table of Contents")]', namespaces=NS))
            if c is not None:
                for e in c:
                    if e.tag in (q('p'), q('tbl')):
                        e.set('_sdt_toc', '1' if toc_sdt else '0'); items.append(e)
    in_toc = False; toc_seen = False; has_toc_field = bool(D.doc.xpath('.//w:instrText[contains(.,"TOC")]|.//w:fldSimple[contains(@w:instr,"TOC")]', namespaces=NS))
    if not has_toc_field: add('TOC2', 'нет поля оглавления (оглавление набрано вручную или отсутствует)')
    body_stats = dict(jc=0, line=0, ind=0, font=0, sz=0, n=0, spb=0)
    prev_empty = 0; empties_run = 0; pg2_hits = 0
    prev_item_was_caption_tab = False
    q_straight = 0; dash_hyph = 0; dbl_sp = 0; hl = 0
    nh_head = 0; nk_head = 0; head_n = 0
    spacing_heads = 0
    heads = []
    last_p = None
    in_app = False; apps = []; pend_app = False
    tab_main = tab_app = fig_main = fig_app = 0
    toc_texts = []; toc_fmt = dict(n=0, bad=0)
    nums, absn = load_numbering(D); seq = []; in_bib = False; toc_paras = []
    seen_intro = False; pre_intro = []; recent_pb = False
    for idx, el in enumerate(items):
        if el.tag == q('tbl'):
            if pend_app: add('AP1', 'у приложения нет заголовка (сразу таблица)'); pend_app = False
            # ширина
            tw = el.find('w:tblPr/w:tblW', NS)
            grid = [int(g.get(q('w')) or 0) for g in el.findall('w:tblGrid/w:gridCol', NS)]
            width = None
            if tw is not None and tw.get(q('type')) == 'dxa': width = int(tw.get(q('w')))
            elif grid: width = sum(grid)
            if width and textw and width > textw + 200: add('TB3', 'таблица шире поля: %d > %d' % (width, textw))
            # подпись выше?
            prev = last_p
            ptxt = ptext(prev).strip() if prev is not None else ''
            rows_ = el.findall('w:tr', NS); cells_ = rows_[0].findall('w:tc', NS) if rows_ else []
            data_like = len(rows_) >= 2 and len(cells_) >= 2 and sum(len(ptext(c)) for c in el.iter(q('tc'))) > 150
            if data_like and not re.match(r'^(Таблица|Продолжение|Окончание)', ptxt, re.I):
                add('TB0', 'таблица без подписи над ней')
            empties_run = 0
            last_p = None
            continue
        p = el
        text = ptext(p).strip()
        sname = D.style_name(p)
        is_toc_style = bool(re.search(r'toc|оглавл|содерж', sname))
        # граница оглавления
        up = text.upper()
        if up in ('СОДЕРЖАНИЕ', 'ОГЛАВЛЕНИЕ') and not toc_seen:
            in_toc = True; toc_seen = True
            if up == 'ОГЛАВЛЕНИЕ': add('TOC1', 'заголовок «ОГЛАВЛЕНИЕ» вместо «СОДЕРЖАНИЕ»')
        elif in_toc and (up == 'ВВЕДЕНИЕ' or RE_STRUCT.match(text) and not re.search(r'\d\s*$', text) and not is_toc_style and p.get('_sdt_toc') != '1'):
            if up == 'ВВЕДЕНИЕ' or re.match(r'^(РЕФЕРАТ|ABSTRACT)$', up): in_toc = False
        in_toc_here = in_toc or is_toc_style or p.get('_sdt_toc') == '1'
        has_draw = bool(p.xpath('.//w:drawing|.//w:pict|.//w:object', namespaces=NS))
        pb = bool(p.xpath('.//w:br[@w:type="page"]|.//w:pageBreakBefore', namespaces=NS))
        pb = pb or bool(p.xpath('.//w:sectPr', namespaces=NS))
        pb_before = recent_pb or pb or bool(D.pprop(p, g_flag('pageBreakBefore')))
        # пустые абзацы подряд
        if not text and not has_draw and not pb and not p.xpath('.//w:sectPr', namespaces=NS):
            empties_run += 1
            if empties_run == 4: add('PG2', 'четыре и более пустых абзаца подряд (около «%s»)' % ptext(last_p).strip()[:30] if last_p is not None else 'три пустых абзаца подряд')
        else:
            empties_run = 0
        # гиперссылки вне оглавления
        if not in_toc_here:
            for h in p.iter(q('hyperlink')):
                a = h.get(q('anchor')) or ''
                if not a.startswith(('_Toc', '_Ref', '_GoBack')): hl += 1
            if p.xpath('.//w:rStyle[@w:val="Hyperlink" or @w:val="a8" or contains(@w:val,"yperlink")]', namespaces=NS) and not p.xpath('.//w:hyperlink', namespaces=NS): hl += 1
        if text and not in_toc_here: recent_pb = False
        elif pb: recent_pb = True
        if not text:
            last_p = p if has_draw else last_p
            if not has_draw: pass
            else: last_p = p
            continue
        if in_toc_here:
            toc_texts.append(text); toc_paras.append(p)
            if up not in ('СОДЕРЖАНИЕ', 'ОГЛАВЛЕНИЕ'):
                toc_fmt['n'] += 1
                line_ = D.pprop(p, g_sp('line')); bf_ = D.pprop(p, g_sp('before')); af_ = D.pprop(p, g_sp('after'))
                szs = [D.rprop(r_, p, g_sz) for r_ in p.findall('.//w:r', NS) if ''.join(t_.text or '' for t_ in r_.findall('w:t', NS)).strip()]
                bad_ = (szs and any(z_ != 28 for z_ in szs)) or not (line_ and abs(int(line_) - 360) <= 8) or (bf_ and int(bf_) > 0) or (af_ and int(af_) > 0)
                toc_fmt['bad'] += bool(bad_)
            last_p = p; continue
        if not seen_intro:
            if up == 'ВВЕДЕНИЕ': seen_intro = True
            else: pre_intro.append(text)
        if up == 'ABSTRACT' and not pb_before: add('RE1', 'ABSTRACT не с новой страницы')
        # --- типы абзацев
        is_chap = bool(RE_CHAP.match(text)); is_sub = bool(RE_SUB.match(text)) and len(text) < 200 and not text.endswith(('.', ';', ':')) or bool(re.match(r'^\d+(\.\d+)+\.\s', text) and len(text) < 160)
        is_struct = bool(RE_STRUCT.match(text)) and len(text) < 120
        is_vyv = bool(RE_VYV.match(text))
        sty_head = ('heading' in sname or 'заголовок' in sname) and len(text) < 200
        is_head = is_chap or is_sub or is_struct or is_vyv or sty_head
        is_cap_t = bool(re.match(r'^(Таблица|Продолжение\s+табл|Окончание\s+табл)', text, re.I)) and len(text) < 250
        is_cap_f = bool(re.match(r'^(Рисунок|Рис\.)\s*(№\s*)?[\dА-ЯA-Z]', text)) and len(text) < 250
        if pend_app:
            if is_cap_t or is_cap_f or re.match(r'^ПРИЛОЖЕНИЕ\s+\S', text, re.I): add('AP1', 'у приложения нет заголовка')
            pend_app = False
        is_list = D.pprop(p, lambda ppr: True if ppr.find('w:numPr', NS) is not None else None) or 'list' in sname or 'спис' in sname
        # перечисления (LS1, LS2): собираем последовательность маркеров, разбор по прогонам — после обхода
        if len(text) < 90 and RE_BIB.match(up) and not re.search(r'\d\s*$', text) and (text == up or is_head or len(text) < 45): in_bib = True
        elif is_head and re.match(r'^ПРИЛОЖЕНИЕ', up): in_bib = False
        mk = None if (is_head or is_cap_t or is_cap_f or not seen_intro) else list_marker(D, p, text, nums, absn)   # до введения — титул, задание, реферат: не перечисления
        seq.append((mk, text, in_bib))
        # пробелы / кавычки / дефисы
        cyr = len(re.findall('[А-Яа-я]', text)); lat = len(re.findall('[A-Za-z]', text))
        if cyr > lat:
            q_straight += len(re.findall(r'["“„]\s?[А-Яа-яЁё][^"“”„«»]{0,80}[А-Яа-яЁё.,!?]["”“]', text))
            dash_hyph += len(re.findall(r'\s-\s', text))
        dbl_sp += len(re.findall(r'(?<=\S)  +(?=\S)', text))
        if is_head:
            head_n += 1
            heads.append(text)
            jc = D.pprop(p, g_jc); fl = D.pprop(p, g_first)
            sup = D.pprop(p, g_flag('suppressAutoHyphens')); kn = D.pprop(p, g_flag('keepNext'))
            before = D.pprop(p, g_sp('before')); after = D.pprop(p, g_sp('after'))
            struct_like = is_struct or is_vyv or is_chap
            if PROFILE_H == 'gost' and not is_struct:
                # ГОСТ 7.32-2017, 6.2.3: заголовки разделов и подразделов — с абзацного отступа, полужирным, не по центру
                if jc == 'center':
                    add('H5', 'заголовок раздела/подраздела по центру (по ГОСТ 6.2.3 — с абзацного отступа): «%s»' % text[:40])
                elif fl in (None, 0, '0'):
                    add('H5', 'нет абзацного отступа в заголовке (ГОСТ 6.2.3): «%s»' % text[:40])
                runs_ = [r_ for r_ in p.findall('.//w:r', NS) if ''.join(t_.text or '' for t_ in r_.findall('w:t', NS)).strip()]
                if runs_ and not all(D.rprop(r_, p, g_b) for r_ in runs_):
                    add('H5', 'заголовок не полужирным (ГОСТ 6.2.3): «%s»' % text[:40])
            else:
                if jc != 'center':
                    add('H1' if is_struct else 'H5', 'заголовок структурного элемента не по центру (ГОСТ 6.2.1): «%s»' % text[:40])
                if fl not in (None, 0, '0') and fl != 0: add('H5', 'красная строка в заголовке: «%s»' % text[:40])
            if auto_h and not sup: add('H9', 'нет запрета переноса в заголовке: «%s»' % text[:40])
            if False and not kn: add('H8', 'заголовок без «не отрывать от следующего»: «%s»' % text[:40])
            if (before not in (None, '0')) or (after not in (None, '0')):
                if (before and int(before) > 0) or (after and int(after) > 0): spacing_heads += 1
            if re.match(r'^\d+(\.\d+)*\.\s', text) or (is_sub and text.endswith('.')) or (re.match(r'^\d+(\.\d+)+', text) and text.endswith('.')):
                add('H6', 'точка после номера или в конце заголовка: «%s»' % text[:40])
            if re.match(r'^(ГЛАВА|Глава)\s+[IVXLC]+\b', text): add('H7', 'римская цифра в номере главы: «%s»' % text[:40])
            if is_vyv and not text.split('.')[0].isupper(): add('H7', 'вывод по главе не прописными: «%s»' % text[:40])
            if is_vyv and re.search(r'\b[IVX]+\b', text): add('H7', 'римская цифра в выводах: «%s»' % text[:40])
            if re.match(r'^Приложение', text): add('AP1', 'заголовок приложения не прописными: «%s»' % text[:40])
            if re.match(r'^ПРИЛОЖЕНИЯ?\b', text, re.I): in_app = True
            m_app = re.match(r'^ПРИЛОЖЕНИЕ\s+(\S+)', text, re.I)
            if m_app:
                aid = m_app.group(1).strip('.:–-')
                apps.append(aid)
                a_ = aid.upper()
                ok_ = (re.fullmatch(r'[А-Я]', a_) and a_ not in 'ЗЙОЧЪЫЬ') or (re.fullmatch(r'[A-Z]', a_) and a_ not in 'IO')   # ГОСТ 7.32-2017, 6.17.4
                if not ok_: add('AP1', 'обозначение приложения не по ГОСТ 6.17.4: «%s»' % text[:30])
                pend_app = len(text) <= len(m_app.group(0)) + 2
            if is_struct and up in ('ВВЕДЕНИЕ', 'ЗАКЛЮЧЕНИЕ', 'СОДЕРЖАНИЕ', 'ОГЛАВЛЕНИЕ', 'РЕФЕРАТ', 'ABSTRACT') or re.match(r'^СПИСОК\s', up):
                if not pb_before: add('PG2', 'структурный заголовок не с новой страницы: «%s»' % text[:30])
            if PROFILE_H == 'gost' and is_chap and not pb_before:
                add('PG2', 'глава не с новой страницы (ГОСТ 6.2.1): «%s»' % text[:30])
            if is_struct and text != text.upper() and not re.match(r'^(Приложение)', text): add('H1', 'структурный заголовок не прописными: «%s»' % text[:40])
            last_p = p; continue
        if is_cap_t:
            jc = D.pprop(p, g_jc)
            if not re.match(r'^(Продолжение|Окончание)', text, re.I):
                if in_app: tab_app += 1
                else: tab_main += 1
            if in_app and re.match(r'^Таблица\s+\d', text): add('AP1', 'в приложении нумерация таблицы не буквенная: «%s»' % text[:30])
            if '№' in text[:15]: add('TB1', 'знак № в подписи таблицы: «%s»' % text[:40])
            if not re.match(r'^(Таблица|Продолжение таблицы|Окончание таблицы)\s+(\d+\.\d+|[А-ЯA-Z]\.\d+)\s+[–—-]', text) and not re.match(r'^Продолжение таблицы|^Окончание таблицы', text, re.I):
                add('TB1', 'подпись таблицы не по ГОСТ (нумерация по главам, тире): «%s»' % text[:40])
            if jc in ('center', 'right'): add('TB1', 'подпись таблицы не слева: «%s»' % text[:40])
            if re.match(r'^(Продолжение|Окончание)\s+табл\.', text, re.I) or re.match(r'^Продолжение\s+таблицы', text, re.I) and jc == 'right':
                add('TB2', 'продолжение таблицы: сокращение или выравнивание справа: «%s»' % text[:40])
            if text.rstrip().endswith('.') and not text.rstrip().endswith('..'): add('TB1', 'точка в конце подписи таблицы: «%s»' % text[:40])
            last_p = p; continue
        if is_cap_f:
            jc = D.pprop(p, g_jc)
            if in_app: fig_app += 1
            else: fig_main += 1
            if in_app and re.match(r'^(Рисунок|Рис\.)\s+\d', text): add('AP1', 'в приложении нумерация рисунка не буквенная: «%s»' % text[:30])
            if text.startswith('Рис.'): add('FG1', 'сокращение «Рис.»: «%s»' % text[:40])
            if '№' in text[:15]: add('FG1', 'знак № в подписи рисунка: «%s»' % text[:40])
            if not re.match(r'^(Рисунок|Рис\.)\s+(\d+\.\d+|[А-ЯA-Z]\.\d+)\s+[–—-]', text): add('FG1', 'подпись рисунка не по ГОСТ: «%s»' % text[:40])
            if jc != 'center': add('FG1', 'подпись рисунка не по центру: «%s»' % text[:40])
            last_p = p; continue
        # основной текст
        if len(text) >= 60 and not is_list and not has_draw:
            body_stats['n'] += 1
            jc = D.pprop(p, g_jc); fl = D.pprop(p, g_first); line = D.pprop(p, g_sp('line')); lr = D.pprop(p, g_sp('lineRule'))
            bf = D.pprop(p, g_sp('before')); af = D.pprop(p, g_sp('after'))
            if jc not in ('both', 'distribute'): body_stats['jc'] += 1
            if not (line and abs(int(line) - 360) <= 8 and lr in (None, 'auto')): body_stats['line'] += 1
            if not (isinstance(fl, int) and abs(fl - 709) <= 12): body_stats['ind'] += 1
            if (bf and int(bf) > 0) or (af and int(af) > 0): body_stats['spb'] += 1
            # шрифт и кегль по первым ранам с текстом
            bad_font = bad_sz = False
            for r in p.findall('.//w:r', NS):
                if not ''.join(t.text or '' for t in r.findall('w:t', NS)).strip(): continue
                f = D.rprop(r, p, g_font); sz = D.rprop(r, p, g_sz)
                if f and f.startswith('THEME:'): f = D.theme_minor
                if f != 'Times New Roman': bad_font = True
                if sz != 28: bad_sz = True
            if bad_font: body_stats['font'] += 1
            if bad_sz: body_stats['sz'] += 1
        last_p = p
    # --- LS1 / LS2: разбор прогонов перечислений
    run = []
    def flush():
        nonlocal run
        if run:
            bib_ = run[0][2]; m0 = run[0][0]
            if bib_:
                if len(run) >= 3 and m0['kind'].startswith('digit'): ls_check([(m, t) for m, t, b in run], add, bib=True)
            elif m0['kind'] != 'multi' and (m0['auto'] or len(run) >= 2):
                ls_check([(m, t) for m, t, b in run], add)
        run = []
    for mk, t_, b_ in seq:
        if mk is None: flush(); continue
        if run and (run[0][2] != b_ or (run[0][0]['kind'], run[0][0]['auto'], run[0][0].get('numid'), run[0][0].get('lvl')) != (mk['kind'], mk['auto'], mk.get('numid'), mk.get('lvl'))): flush()
        run.append((mk, t_, b_))
    flush()
    pre = ' '.join(pre_intro)
    def num(pats):
        for p_ in pats:
            m_ = re.search(p_, pre, re.I)
            if m_: return int(m_.group(1))
    a_fig = num([r'(\d+)\s*рис', r'(\d+)\s*figure']); a_tab = num([r'(\d+)\s*табл', r'(\d+)\s*table']); a_app = num([r'(\d+)\s*прил', r'(\d+)\s*appendi'])
    if re.search(r'\b0\s*(рис|табл|прил|figure|table|appendi)', pre, re.I): add('RE1', 'в реферате указан нуль')
    for lab, a_, m_, ap_ in (('рисунков', a_fig, fig_main, fig_app), ('таблиц', a_tab, tab_main, tab_app)):
        if a_ is not None and a_ != m_:
            if ap_ and a_ == m_ + ap_: add('NP2', 'в реферат включены %s приложений' % lab)
            else: add('RE1', 'в реферате %d %s, в тексте %d' % (a_, lab, m_))
    if a_app is not None and a_app != len(set(apps)): add('RE1', 'в реферате %d приложений, в тексте %d' % (a_app, len(set(apps))))
    toc_stat = toc_check(D, toc_paras, textw, add) if toc_paras else {}
    toc_compose(toc_paras, heads, add)
    if toc_texts:
        if apps and not any(re.search('ПРИЛОЖЕНИ', t_, re.I) for t_ in toc_texts): add('TOC1', 'приложения не вынесены в содержание')
        if not any(re.search('СПИСОК', t_, re.I) for t_ in toc_texts): add('TOC1', 'список источников не в содержании')
        if toc_fmt['n'] >= 3 and toc_fmt['bad'] / toc_fmt['n'] > .3: add('TOC1', 'содержание: кегль, интервал или отступы: %d из %d строк' % (toc_fmt['bad'], toc_fmt['n']))
    # итоги по основному тексту: порог 10 % абзацев
    n = max(body_stats['n'], 1)
    def share(k): return body_stats[k] / n
    if share('jc') > TXS: add('TX1', 'не по ширине: %d из %d абзацев' % (body_stats['jc'], n))
    if share('line') > TXS: add('TX1', 'межстрочный не 1,5: %d из %d' % (body_stats['line'], n))
    if share('ind') > TXS: add('TX1', 'отступ не 1,25 см: %d из %d' % (body_stats['ind'], n))
    if share('font') > TXS: add('TX1', 'шрифт не Times New Roman: %d из %d' % (body_stats['font'], n))
    if share('sz') > TXS: add('TX1', 'кегль не 14: %d из %d' % (body_stats['sz'], n))
    if body_stats['spb'] / n > .1 or spacing_heads >= 3:
        add('V1', 'интервалы до/после не равны 0: основной текст %d из %d, заголовки %d' % (body_stats['spb'], n, spacing_heads))
    if hl: add('HL1', 'гиперссылок вне оглавления: %d' % hl)
    if q_straight >= 2: add('Q1', 'не «ёлочки» вокруг русского текста: %d' % q_straight)
    if False and dash_hyph >= 3: add('D1', 'дефис вместо тире: %d' % dash_hyph)
    if dbl_sp >= 5: add('ST1', 'двойные пробелы: %d' % dbl_sp)
    return F, dict(toc=({k: v for k, v in toc_stat.items() if k != 'lv'} if toc_stat else None), toc_lv=({k: sorted(set(v)) for k, v in toc_stat['lv'].items()} if toc_stat else None), n_body=body_stats['n'], n_head=head_n, bs=body_stats, sph=spacing_heads)

if __name__ == '__main__':
    folder = sys.argv[1]; out = sys.argv[2]
    res = {}
    for f in sorted(glob.glob(os.path.join(folder, '*.docx'))):
        name = re.sub(r'^1_|\.docx$', '', os.path.basename(f))
        try:
            F, meta = check(f)
            res[name] = dict(findings=F, meta=meta)
            pass
        except Exception as e:
            res[name] = dict(error=repr(e)); print(name, 'ERR', repr(e), flush=True)
    json.dump(res, open(out, 'w', encoding='utf8'), ensure_ascii=False, indent=1)

    comp = {n: {k: len(v) for k, v in d.get('findings', {}).items()} for n, d in res.items() if 'findings' in d}
    json.dump(comp, open(out.replace('.json', '_counts.json'), 'w', encoding='utf8'), ensure_ascii=False)
    print(json.dumps(comp, ensure_ascii=False, separators=(',', ':')))
