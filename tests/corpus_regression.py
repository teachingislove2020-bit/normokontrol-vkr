"""Регрессия на корпусе (работы в репозиторий не кладутся): python tests/corpus_regression.py <папка_с_docx>

Якорь 1 — слой A не менялся: профиль metod, с титульным листом → 4055 замечаний на 46 работах (база v0.8).
Якорь 2 — сервис: профиль gost, без титульного листа, слой A + список источников → число из BASE_SERVICE (коммит f9aeaf9).
Кроме того, сервис без усечения отдаёт столько же замечаний, сколько находит слой A (за вычетом отключённых правил).
"""
import sys, os, io, glob, json, hashlib
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'web', 'py'))
import normokontrol_core as C, normokontrol_biblio as B, normokontrol_service as S

BASE_LAYER_A, BASE_SERVICE = 4055, 4359
SHA = {'normokontrol_core.py': '08b023b8b7c0dc1747b064faf7ae5ca9fd8b49eb56938bf2c2cfc26470974b08',
       'normokontrol_biblio.py': '0fccd811ea21cc07fb13cb1edeb782cbe9d97ae966fbc2b6dfacd61b1b93220f'}
folder = sys.argv[1]
files = sorted(glob.glob(os.path.join(folder, '*.docx')))
ok = True
for n, h in SHA.items():
    cur = hashlib.sha256(open(os.path.join(os.path.dirname(__file__), '..', 'web', 'py', n), 'rb').read()).hexdigest()
    print('%-26s %s' % (n, 'не изменён' if cur == h else 'ИЗМЕНЁН'))
    ok &= cur == h
C.configure(profile='metod', toc='gost', headings='gost', no_title=False)
a = sum(len(C.check(io.BytesIO(open(f, 'rb').read()))[1]['findings']) for f in files)
print('слой A (metod, с титульным): %d, ожидается %d, работ %d' % (a, BASE_LAYER_A, len(files)))
ok &= a == BASE_LAYER_A
with open(os.path.join(os.path.dirname(__file__), '..', 'web', 'rules.json'), encoding='utf8') as f:
    S.load_rules(f.read())
tot = items = 0
for f in files:
    data = open(f, 'rb').read()
    C.configure(profile='gost', toc='gost', headings='gost', no_title=True)
    fl = len(C.check(io.BytesIO(data))[1]['findings']) + len(B.check_biblio(io.BytesIO(data)))
    r = json.loads(S.run(data))
    got = sum(g['count'] for g in r['groups'])
    shown = sum(len(g['items']) for g in r['groups'])
    tot += fl; items += got
    assert got == shown, (f, got, shown)       # ничего не усечено: решение можно принять по каждому замечанию
print('сервис (gost, без титульного): найдено %d, в отчёте %d, ожидается %d' % (tot, items, BASE_SERVICE))
ok &= tot == BASE_SERVICE and items == tot
print('РЕГРЕССИЯ:', 'OK' if ok else 'РАСХОЖДЕНИЕ')
sys.exit(0 if ok else 1)
