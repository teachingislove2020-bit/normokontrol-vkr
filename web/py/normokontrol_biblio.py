"""Проверка списка использованных источников и ссылок на него.
Правила: L0, L1, REF0, REF1, REF2, LS2n, ORD1, DUP, BIB1–BIB6 (см. rules.json).
Порядок записей — алфавитный (решение 10), нумерация — по ГОСТ 7.32-2017, п. 6.16 (решение 11),
описание — по ГОСТ Р 7.0.100-2018 (решение 12)."""
import re
from lxml import etree
import normokontrol_core as C

NS = C.NS
q = C.q
DASH = r'\-–—−'
RE_APP = re.compile(r'^ПРИЛОЖЕНИЕ', re.I)
RE_HEAD_STYLE = re.compile(r'heading|заголовок', re.I)

# ---------------------------------------------------------------------------
# ссылки в тексте
RE_BR = re.compile(r'\[([^\[\]\n]{1,80})\]')
RE_RANGE_BR = re.compile(r'\[(\d{1,3})\]\s*[%s]\s*\[(\d{1,3})\]' % DASH)
RE_PAGE = re.compile(r'(?i)\s*[,;]?\s*(?:с|стр|p|pp)\.\s*[\d\s%s,]*' % DASH)
RE_OK_INNER = re.compile(r'^\d{1,3}(?:\s*[%s]\s*\d{1,3})?(?:\s*[,;]\s*\d{1,3}(?:\s*[%s]\s*\d{1,3})?)*$' % (DASH, DASH))


def parse_citations(text):
    """Список (номера, фрагмент) ссылок вида [1], [1, 3], [1–5], [1]–[4], [13, с. 112], [5; 17; 34]."""
    out = []
    used = []
    for m in RE_RANGE_BR.finditer(text):
        a, b = int(m.group(1)), int(m.group(2))
        if 0 < a <= b and b - a < 200:
            out.append((list(range(a, b + 1)), m.group(0), False))
            used.append(m.span())
    for m in RE_BR.finditer(text):
        if any(s <= m.start() < e or s < m.end() <= e for s, e in used):
            continue
        inner = m.group(1)
        inner = RE_PAGE.sub('', inner).strip().rstrip(',;').strip()
        if not inner or not RE_OK_INNER.match(inner):
            continue
        nums = []
        for part in re.split(r'\s*[,;]\s*', inner):
            r = re.split(r'\s*[%s]\s*' % DASH, part)
            if len(r) == 2:
                a, b = int(r[0]), int(r[1])
                if a <= b and b - a < 200:
                    nums.extend(range(a, b + 1))
                else:
                    nums.extend([a, b])
            else:
                nums.append(int(r[0]))
        # «[44, 155]»: вторая цифра чаще всего страница (методичка просит «с.»)
        out.append((nums, m.group(0), len(nums) == 2 and ',' in inner and ';' not in inner and not re.search(r'[%s]' % DASH, inner.split(',')[0])))
    return out


# ---------------------------------------------------------------------------
# обход документа
def iter_blocks(D):
    """Абзацы документа в порядке следования: (p, in_table). Страницы считаются по меткам Word."""
    body = D.doc.find('w:body', NS)
    res = []
    def walk(el, in_tbl):
        for ch in el:
            if ch.tag == q('p'):
                res.append((ch, in_tbl))
            elif ch.tag == q('tbl'):
                for p in ch.iter(q('p')):
                    res.append((p, True))
            elif ch.tag == q('sdt'):
                c = ch.find('w:sdtContent', NS)
                if c is not None:
                    walk(c, in_tbl)
    walk(body, False)
    return res


def find_bib(paras):
    """Индексы: заголовок списка, начало, конец (не включая). paras — [(p, text, in_tbl)]"""
    intro = None
    for i, (p, t, tb) in enumerate(paras):
        if t.upper() == 'ВВЕДЕНИЕ' and not tb:
            intro = i
            break
    cand = []
    for i, (p, t, tb) in enumerate(paras):
        if tb or not t or len(t) >= 90:
            continue
        up = t.upper()
        if C.RE_BIB.match(up) and not re.search(r'\d\s*$', t) and (t == up or len(t) < 45):
            cand.append(i)
    if intro is not None:
        cand = [i for i in cand if i > intro]
    if not cand:
        return None
    h = cand[-1]
    end = len(paras)
    for j in range(h + 1, len(paras)):
        t = paras[j][1]
        if t and not paras[j][2] and (RE_APP.match(t) and len(t) < 60):
            end = j
            break
    return h, h + 1, end


def collect(D):
    nums, absn = C.load_numbering(D)
    blocks = iter_blocks(D)
    paras = []
    page = 1
    pages = []
    for p, tb in blocks:
        page += len(p.findall('.//w:lastRenderedPageBreak', NS))
        paras.append((p, C.ptext(p).strip(), tb))
        pages.append(page)
    pages_known = bool(D.doc.xpath('.//w:lastRenderedPageBreak', namespaces=NS))
    return nums, absn, paras, pages, pages_known


def footnote_texts(D):
    out = []
    for name in ('word/footnotes.xml', 'word/endnotes.xml'):
        if name in D.z.namelist():
            x = etree.fromstring(D.z.read(name))
            for p in x.iter(q('p')):
                t = C.ptext(p).strip()
                if t:
                    out.append(t)
    return out


# ---------------------------------------------------------------------------
# записи списка
RE_NUM = re.compile(r'^\s*(\d{1,3})\s*([.)])?\s+(?=\S)')


def split_records(D, paras, nums, absn, start, end):
    """Записи: dict(num, text, par, marker, group, numbered). Подзаголовки списка пропускаются, но разделяют группы."""
    counters = {}
    rows = []
    for i in range(start, end):
        p, t, tb = paras[i]
        if not t:
            continue
        mk = C.list_marker(D, p, t, nums, absn)
        num = None
        text = t
        if mk and mk.get('auto') and not mk['kind'].startswith(('dash', 'bullet')):
            key = (mk['numid'], mk['lvl'])
            counters[key] = counters.get(key, mk.get('start', 1) - 1) + 1
            num = counters[key]
        elif mk and not mk.get('auto') and mk.get('ord') is not None and mk['kind'].startswith('digit'):
            num = mk['ord']
            text = re.sub(r'^\s*\d{1,3}\s*[.)]\s*', '', t)
        else:
            m = RE_NUM.match(t)
            if m:
                num = int(m.group(1))
                text = t[m.end():]
        rows.append((i, t, text, num, mk))
    numbered = sum(1 for r in rows if r[3] is not None)
    mode_num = numbered >= max(2, 0.5 * len(rows))
    recs = []
    group = 0
    prev = None
    for i, t, text, num, mk in rows:
        if mode_num:
            if num is None:
                looks_head = len(t) < 70 and not re.search(r'\d{4}', t) and not re.search(r'[/:]', t)
                if looks_head or prev is None:
                    group += 1
                    prev = None
                    continue
                prev['text'] += ' ' + t
                continue
            prev = dict(num=num, text=text.strip(), par=i, marker=mk, group=group, numbered=True)
            recs.append(prev)
        else:
            if len(t) < 40 and not re.search(r'\d{4}', t):
                group += 1
                prev = None
                continue
            if prev is not None and mk is None and not re.match(r'^[A-ZА-ЯЁ]', t):
                prev['text'] += ' ' + t
                continue
            prev = dict(num=len(recs) + 1, text=t.strip(), par=i, marker=mk, group=group, numbered=False)
            recs.append(prev)
    return recs


# ---------------------------------------------------------------------------
# порядок
def first_alpha(s):
    m = re.search(r'[A-Za-zА-Яа-яЁё]', s)
    return m.group(0) if m else ''


def is_cyr(s):
    return bool(re.match(r'[А-Яа-яЁё]', first_alpha(s)))


def sort_key(text):
    t = text.lower().replace('ё', 'е')
    t = re.sub(r'^[^0-9a-zа-я]+', '', t)
    toks = re.findall(r'[0-9]+|[a-zа-я]+', t)
    out = []
    for w in toks[:12]:
        out.append((0, '%08d' % int(w)) if w.isdigit() else (1, w))
    return out


def lis_keep(keys):
    """Индексы наибольшей неубывающей подпоследовательности."""
    n = len(keys)
    if not n:
        return set()
    best = [1] * n
    prev = [-1] * n
    for i in range(n):
        for j in range(i):
            if keys[j] <= keys[i] and best[j] + 1 > best[i]:
                best[i] = best[j] + 1
                prev[i] = j
    i = max(range(n), key=lambda k: best[k])
    keep = set()
    while i != -1:
        keep.add(i)
        i = prev[i]
    return keep


# ---------------------------------------------------------------------------
# описание
RE_YEAR = re.compile(r'(?<!\d)(1[89]\d{2}|20[0-4]\d)(?!\d)')
RE_URL = re.compile(r'https?://|www\.|URL\s*:', re.I)
RE_ACCESS = re.compile(r'\(\s*дата\s+обращения\s*:\s*\d{2}\.\d{2}\.\d{4}\s*\)', re.I)
RE_ACCESS_ANY = re.compile(r'дата\s+обращения|дата\s+публикации|дата\s+доступа|обращения', re.I)
RE_PAGES = re.compile(r'(?:^|[\s.–—-])(?:С|P|Р|p|pp|с)\.\s*\d+', )
RE_VOL = re.compile(r'\d+\s*(?:с|p|c|стр)\.?(?:\s|$|\.|,)', re.I)


def bib_checks(text):
    """Список (подправило, сообщение) для записи."""
    out = []
    low = text.lower()
    is_url = bool(RE_URL.search(text))
    raw = text
    text = re.sub(r'https?://\S+|www\.\S+', ' ', text)   # адреса не содержат разделителя «//» описания
    is_art = ' // ' in text or bool(re.search(r'\S//\S|\s//\S|\S//\s', text))
    is_std = bool(re.match(r'^(ГОСТ|ГОСТ\s*Р|СП|СНиП|ISO|ОК|ТУ)\b', text))
    if not RE_YEAR.search(text) and not is_url:
        out.append(('BIB1', 'в описании не найден год издания'))
    if '//' in text and ' // ' not in text:
        out.append(('BIB2', 'знак «//» отделяется пробелами с двух сторон'))
    if ' // ' in text and not RE_PAGES.search(text.split(' // ', 1)[1]) and not is_url:
        out.append(('BIB2', 'у статьи нет указания страниц («С. 25–29»)'))
    text_full = raw
    if is_url:
        if not re.search(r'URL\s*:', text_full):
            out.append(('BIB3', 'адрес в сети без «URL:»'))
        if not RE_ACCESS.search(text_full) and not re.search(r'(?i)дата\s+публикации\s*:', text_full):
            if RE_ACCESS_ANY.search(text_full):
                out.append(('BIB3', 'дата обращения оформлена не как «(дата обращения: ДД.ММ.ГГГГ)»'))
            else:
                out.append(('BIB3', 'нет даты обращения «(дата обращения: ДД.ММ.ГГГГ)»'))
    if re.search(r'\[Электронный ресурс\]', text, re.I):
        out.append(('BIB4', '«[Электронный ресурс]» не применяется в Р 7.0.100-2018 (указывают «Текст : электронный»)'))
    if re.search(r'Режим\s+доступа', text, re.I):
        out.append(('BIB4', '«Режим доступа» заменяется на «URL:» (Р 7.0.100-2018)'))
    if not is_url and not is_art and not is_std and ' // ' not in text and not RE_VOL.search(text) \
            and not re.search(r'(?i)\bдис\b|автореф|\bdiss|thesis|\bp\.\s*\d', text):
        out.append(('BIB5', 'для книги не указан объём («208 с.»)'))
    if re.search(r'\s[-]\s', text) and not re.search(r'\s[–—]\s', text):
        out.append(('BIB6', 'области описания разделены дефисом, нужно тире с пробелами'))
    return out


# ---------------------------------------------------------------------------
def check_biblio(src):
    """Возвращает список замечаний в том же формате, что слой A."""
    D = C.Doc(src)
    nums, absn, paras, pages, pages_known = collect(D)
    FL = []

    def add(rule, msg, quote='', head='', par=None, page=None, examples=None, sub=None):
        FL.append(dict(rule=rule, sub=sub, msg=msg, quote=quote, head=head, par=par,
                       page=page if pages_known else None, examples=list(examples or [])))

    bib = find_bib(paras)
    if not bib:
        add('L0', 'не найден раздел «Список использованных источников»')
        return FL
    h, start, end = bib
    HEAD = paras[h][1][:80]
    recs = split_records(D, paras, nums, absn, start, end)
    pg = lambda i: pages[i] if i is not None and i < len(pages) else None
    n = len(recs)

    # L1
    if 0 < n < 15:
        add('L1', 'в списке %d источников, нужно не менее 15 (методичка, п. 3.7)' % n, quote=paras[h][1], head=HEAD, par=h, page=pg(h))

    if recs and not recs[0]['numbered']:
        add('LS2n', 'записи списка не пронумерованы арабскими цифрами с точкой (ГОСТ 7.32-2017, п. 6.16)', quote=recs[0]['text'][:70], head=HEAD, par=recs[0]['par'], page=pg(recs[0]['par']))
    # нумерация: пропуски и повторы
    seen = {}
    prev_n = None
    msgs = []
    for k, r in enumerate(recs if recs and recs[0]['numbered'] else []):
        if r['num'] in seen:
            msgs.append((r, 'номер %d повторяется' % r['num']))
        elif prev_n is not None and r['num'] != prev_n + 1 and r['group'] == recs[k - 1]['group']:
            msgs.append((r, 'нарушена нумерация: после %d идёт %d' % (prev_n, r['num'])))
        elif k == 0 and r['num'] != 1:
            msgs.append((r, 'нумерация начинается с %d, а не с 1' % r['num']))
        seen[r['num']] = r
        prev_n = r['num']
    for r, m in msgs[:5]:
        add('LS2n', m + ' (ГОСТ 7.32-2017, п. 6.16)', quote=r['text'][:70], head=HEAD, par=r['par'], page=pg(r['par']))

    # дубли
    norm = {}
    for r in recs:
        k = re.sub(r'[^0-9a-zа-я]', '', r['text'].lower().replace('ё', 'е'))[:120]
        if len(k) > 25:
            norm.setdefault(k, []).append(r)
    for k, rs in norm.items():
        if len(rs) > 1:
            add('DUP', 'источник повторяется: записи %s' % ', '.join(str(r['num']) for r in rs), quote=rs[1]['text'][:70], head=HEAD, par=rs[1]['par'], page=pg(rs[1]['par']))

    # оформление записей
    agg = {}
    for r in recs:
        for sub, msg in bib_checks(r['text']):
            agg.setdefault(sub, []).append((r, msg))
    for sub, items in agg.items():
        if len(items) <= 3:
            for r, msg in items:
                add(sub, 'запись %d: %s' % (r['num'], msg), quote=r['text'][:70], head=HEAD, par=r['par'], page=pg(r['par']))
        else:
            add(sub, '%s — в %d записях из %d' % (items[0][1], len(items), n), quote='', head=HEAD,
                examples=[dict(quote='%d. %s' % (r['num'], r['text'][:60]), head=HEAD, page=pg(r['par'])) for r, m in items[:5]])

    nums_all = {r['num'] for r in recs}
    maxn = max(nums_all) if nums_all else 0
    if not recs:
        add('L1', 'в разделе «%s» нет ни одной записи (нужно не менее 15)' % HEAD, quote=paras[h][1], head=HEAD, par=h, page=pg(h))
        return FL
    # ссылки в тексте
    cited = {}
    cite_list = []
    head_now = ''
    for i, (p, t, tb) in enumerate(paras):
        if start - 1 <= i < end:
            continue
        if not t:
            continue
        if not tb and len(t) < 150 and (C.RE_STRUCT.match(t) or C.RE_CHAP.match(t) or C.RE_SUB.match(t) or RE_HEAD_STYLE.search(D.style_name(p))):
            head_now = t[:80]
        if '[' not in t:
            continue
        for nums_, frag, mp in parse_citations(t):
            cite_list.append((nums_, frag, i, head_now, t, mp))
            for k_, x in enumerate(nums_):
                if mp and k_ == 1 and x > maxn:
                    continue
                cited.setdefault(x, (i, head_now, frag))
    for t in footnote_texts(D):
        for nums_, frag, mp in parse_citations(t):
            for x in nums_:
                cited.setdefault(x, (None, 'Сноски', frag))
    if not cite_list and not cited:
        add('REF0', 'в тексте нет ссылок на источники в квадратных скобках, например [1] (ГОСТ 7.32-2017, п. 6.9)')
    else:
        bad = []
        seen_b = set()
        for nums_, frag, i, hd, t, mp in cite_list:
            for k_, x in enumerate(nums_):
                if mp and k_ == 1 and x > maxn:
                    continue
                if (x == 0 or x > maxn) and (x, frag) not in seen_b:
                    seen_b.add((x, frag))
                    bad.append((x, frag, i, hd, t))
        for x, frag, i, hd, t in bad[:10]:
            k = t.find(frag)
            ctx = t[max(0, k - 30):k + len(frag) + 10] if k >= 0 else frag
            add('REF1', 'ссылка %s: в списке нет источника с номером %d (всего %d)' % (frag, x, maxn), quote=ctx, head=hd, par=i, page=pg(i))
        if len(bad) > 10:
            add('REF1', 'ещё ссылок на несуществующие номера: %d' % (len(bad) - 10))
        # ORD1: порядок по первым ссылкам (ГОСТ 7.32-2017, п. 6.9.1: номер = порядок первого упоминания)
        first_seen = []
        seen_o = set()
        for nums_, frag, i, hd, t, mp in cite_list:
            for k_, x in enumerate(nums_):
                if mp and k_ == 1:
                    continue
                if 1 <= x <= maxn and x not in seen_o:
                    seen_o.add(x)
                    first_seen.append((x, frag, i, hd))
        out = []
        top = 0
        for x, frag, i, hd in first_seen:
            if x != top + 1:
                out.append((x, frag, i, hd, top + 1))
            top = max(top, x)
        if out:
            if len(out) > 4:
                add('ORD1', 'источники пронумерованы не по порядку первых ссылок: вне порядка %d из %d (вероятно, список составлен по алфавиту)' % (len(out), len(first_seen)),
                    quote=out[0][1], head=out[0][3], par=out[0][2], page=pg(out[0][2]),
                    examples=[dict(quote='ссылка %s, ожидался №%d' % (f, e), head=hd, page=pg(i)) for x, f, i, hd, e in out[:5]])
            else:
                for x, f, i, hd, e in out:
                    add('ORD1', 'первая ссылка на источник %d стоит там, где ожидался №%d' % (x, e), quote=f, head=hd, par=i, page=pg(i))
        if len(cite_list) >= 3:
            unc = [r for r in recs if r['num'] not in cited]
            if unc:
                add('REF2', 'на %d из %d источников нет ссылок в тексте: №%s' % (
                    len(unc), n, ', '.join(str(r['num']) for r in unc[:15]) + ('…' if len(unc) > 15 else '')),
                    quote='', head=HEAD,
                    examples=[dict(quote='%d. %s' % (r['num'], r['text'][:60]), head=HEAD, page=pg(r['par'])) for r in unc[:5]])
    return FL


def summary(src):
    D = C.Doc(src)
    nums, absn, paras, pages, pk = collect(D)
    bib = find_bib(paras)
    if not bib:
        return dict(found=False)
    recs = split_records(D, paras, nums, absn, bib[1], bib[2])
    return dict(found=True, n=len(recs), first=[r['text'][:60] for r in recs[:3]])
