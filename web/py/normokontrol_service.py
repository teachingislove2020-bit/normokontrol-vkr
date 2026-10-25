"""Точка входа для браузера: проверка .docx → отчёт (JSON-строка)."""
import io, json, re, traceback
import normokontrol_core as C
import normokontrol_biblio as B

VERSION = '0.9'
_RULES = None


def load_rules(text):
    global _RULES
    _RULES = json.loads(text)


def _meta(rule, sub):
    R = _RULES['rules'].get(rule, {})
    title = R.get('title', rule)
    if rule == 'TX1' and sub:
        title = _RULES['tx1_sub'].get(sub, title)
    return R, title


def run(data, cfg_json='{}'):
    """data — bytes .docx. Возвращает JSON-строку с отчётом."""
    cfg = json.loads(cfg_json or '{}')
    try:
        C.configure(profile='metod', toc='gost', headings='gost', no_title=True)
        F, meta = C.check(io.BytesIO(bytes(data)))
        findings = list(meta['findings'])
        pages_known = meta['pages_known']
        try:
            findings += B.check_biblio(io.BytesIO(bytes(data)))
        except Exception as e:  # модуль литературы не должен ронять отчёт
            findings.append(dict(rule='L0', sub=None, msg='Не удалось разобрать список источников: %r' % (e,), quote='', head='', page=None, par=None, examples=[]))
    except Exception as e:
        return json.dumps(dict(error='Не удалось прочитать файл как .docx: %s' % (e,)), ensure_ascii=False)

    disabled = set(_RULES.get('disabled', []))
    groups = {}
    order = []
    for f in findings:
        rule = f['rule']
        if rule in disabled:
            continue
        key = (rule, f.get('sub'))
        R, title = _meta(rule, f.get('sub'))
        g = groups.get(key)
        if g is None:
            g = dict(rule=rule, sub=f.get('sub'), title=title, level=R.get('level', 'check'), source=R.get('source', ''),
                     action=R.get('action', ''), count=0, items=[])
            groups[key] = g
            order.append(key)
        g['count'] += 1
        msg, quote, head = f['msg'], (f.get('quote') or ''), (f.get('head') or '')
        if quote and msg.endswith('«' + quote + '»'):
            msg = msg[:-(len(quote) + 2)].rstrip().rstrip(':').rstrip()
        elif quote:
            m_ = re.search(r':\s*«[^»]*»$', msg)
            if m_:
                msg = msg[:m_.start()]
        if quote and head and len(quote) >= 10 and head.startswith(quote[:40]):
            quote = ''
        f = dict(f, msg=msg, quote=quote)
        g['items'].append(dict(msg=f['msg'], quote=f.get('quote') or '', head=f.get('head') or '', page=f.get('page'),
                               examples=f.get('examples') or []))
    out = [groups[k] for k in order]
    # агрегатные замечания учитывают число нарушений в самом тексте сообщения; для сортировки используем число находок
    rank = {'fix': 0, 'check': 1}
    out.sort(key=lambda g: (rank.get(g['level'], 2), g['rule'], g['sub'] or ''))
    for g in out:
        g['more'] = max(0, len(g['items']) - 8)
        g['items'] = g['items'][:8]
    stats = dict(fix_groups=sum(1 for g in out if g['level'] == 'fix'), check_groups=sum(1 for g in out if g['level'] == 'check'),
                 fix_items=sum(g['count'] for g in out if g['level'] == 'fix'), check_items=sum(g['count'] for g in out if g['level'] == 'check'))
    return json.dumps(dict(version=VERSION, pages_known=pages_known, stats=stats, groups=out,
                           self_checks=_RULES.get('self_checks', []), body_paragraphs=meta.get('n_body')), ensure_ascii=False)
