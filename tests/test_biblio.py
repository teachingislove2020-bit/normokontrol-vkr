"""Тесты модуля списка литературы: золотой список, мутации, ссылки."""
import sys, os, random
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'web', 'py'))
sys.path.insert(0, os.path.dirname(__file__))
import normokontrol_core as C
import normokontrol_biblio as B
from make_docx import para, build

# Синтетические записи (не реальные издания), оформлены по ГОСТ Р 7.0.100-2018; русские — по алфавиту, затем иностранные
RU = [
 "Алексеев, А. А. Методика преподавания лексики / А. А. Алексеев. – Москва : Наука, 2015. – 240 с.",
 "Борисова, Е. И. Цифровые ресурсы микрообучения // Вестник образования. – 2019. – № 4. – С. 25–29.",
 "Васильев, П. Р. Иноязычная лексическая компетенция : монография / П. Р. Васильев. – Санкт-Петербург : Питер, 2018. – 312 с.",
 "ГОСТ 7.32-2017. Система стандартов по информации, библиотечному и издательскому делу. Отчет о научно-исследовательской работе. – Москва : Стандартинформ, 2017. – 32 с.",
 "Громова, О. Н. Лексика в контексте обучения // Иностранные языки в школе. – 2020. – № 2. – С. 3–10.",
 "Данилов, С. В. Теория и практика перевода. – Москва : Юрайт, 2017. – 410 с.",
 "Ермакова, Т. К. Микрообучение в вузе : дис. ... канд. пед. наук / Т. К. Ермакова. – Москва, 2016. – 180 с.",
 "Жуков, И. И. Цифровая педагогика / И. И. Жуков. – Санкт-Петербург : Лань, 2021. – 256 с.",
 "Захарова, Л. М. Оценка лексических навыков // Педагогика. – 2018. – № 7. – С. 44–51.",
 "Иванов, В. В. Основы лингводидактики. – Москва : Академия, 2014. – 288 с.",
 "Министерство просвещения Российской Федерации [Официальный сайт]. – URL: https://edu.gov.ru (дата обращения: 14.04.2023). – Текст : электронный.",
 "Петрова, Н. Н. Современные цифровые инструменты // Вопросы филологии. – 2022. – № 1. – С. 77–85. – URL: https://example.org/art1 (дата обращения: 02.03.2023).",
 "Сидоров, К. Л. Языковые компетенции студентов / К. Л. Сидоров. – Казань : Изд-во КФУ, 2019. – 198 с.",
 "Федорова, Ю. А. Методы контроля лексики. – Новосибирск : Сибирское соглашение, 2013. – 154 с.",
]
EN = [
 "Krashen, S. D. Principles and Practice in Second Language Acquisition / S. D. Krashen. – Oxford : Pergamon, 1982. – 202 p.",
 "Nation, I. S. P. Learning Vocabulary in Another Language. – Cambridge : Cambridge University Press, 2001. – 477 p.",
 "Schmitt, N. Vocabulary in Language Teaching. – Cambridge : Cambridge University Press, 2000. – 228 p.",
 "Thornbury, S. How to Teach Vocabulary. – Harlow : Longman, 2002. – 185 p.",
]
GOLD = RU + EN

def doc(records, cites=None, auto=True, head='СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ', extra_after=None):
    n = len(records)
    cites = cites if cites is not None else ' '.join('Утверждение [%d].' % i for i in range(1, n + 1))
    ps = [para('ВВЕДЕНИЕ', center=True), para('Текст работы. ' + cites), para(head, center=True)]
    for i, r in enumerate(records, 1):
        ps.append(para(r, auto_num=True) if auto else para('%d. %s' % (i, r)))
    if extra_after:
        ps += extra_after
    return build(ps)

def rules(src):
    return [(f['rule'], f['msg']) for f in B.check_biblio(src)]

def has(rs, rule): return any(r == rule for r, m in rs)

fails = []
def ok(cond, name):
    print(('PASS ' if cond else 'FAIL ') + name)
    if not cond: fails.append(name)

# 1. золотой список без замечаний
rs = rules(doc(GOLD))
ok(rs == [], 'gold: замечаний нет ' + str(rs))
# 2. ручная нумерация
rs = rules(doc(GOLD, auto=False))
ok(rs == [], 'gold manual: замечаний нет ' + str(rs))
# 3. меньше 15
rs = rules(doc(GOLD[:10]))
ok(has(rs, 'L1'), 'L1: 10 источников')
ok(not has(rules(doc(GOLD[:15])), 'L1'), 'L1: 15 источников — нет')
# 4. ссылка на несуществующий номер
rs = rules(doc(GOLD, cites=' '.join('[%d]' % i for i in range(1, 19)) + ' и [25] и [0]'))
ok(has(rs, 'REF1'), 'REF1: [25]')
# 5. источник без ссылок
rs = rules(doc(GOLD, cites=' '.join('[%d]' % i for i in range(1, 17)) + '.'))
ok(any(r == 'REF2' and '17' in m and '18' in m for r, m in rs), 'REF2: 17 и 18 не процитированы')
# 6. формы ссылок
tests = {'[1]': [1], '[1, 3]': [1, 3], '[1–5]': [1, 2, 3, 4, 5], '[13, с. 112]': [13], '[5; 17; 34]': [5, 17, 34],
         '[1]–[4]': [1, 2, 3, 4], '[9], раздел 5': [9], '[ˈkæt]': None, '[sic]': None, '[…]': None, '[2, С. 3-5]': [2], '[12, 14, с. 5]': [12, 14]}

for t, exp in tests.items():
    got = [n for ns, f, _mp in B.parse_citations('Текст ' + t + ' далее') for n in ns]
    ok((got == exp) if exp is not None else (got == []), 'cite %s -> %s' % (t, got))
# 6b. [номер, страница] без «с.»: страница не считается источником
rs = rules(doc(GOLD, cites=' '.join('[%d]' % i for i in range(1, 19)) + ' Цитата [3, 137].'))
ok(not has(rs, 'REF1'), 'REF1: [3, 137] — страница, не источник')
rs = rules(doc(GOLD, cites=' '.join('[%d]' % i for i in range(1, 19)) + ' Цитата [3; 137].'))
ok(has(rs, 'REF1'), 'REF1: [3; 137] — источник 137 не существует')
# 7. ссылки диапазоном закрывают источники
rs = rules(doc(GOLD, cites='Смотрите [1–18].'))
ok(not has(rs, 'REF2') and not has(rs, 'REF1'), 'диапазон [1–18] покрывает все')
# 8. нет ссылок вообще
rs = rules(doc(GOLD, cites='Текст без ссылок.'))
ok(has(rs, 'REF0'), 'REF0: нет ссылок')
# 9. порядок: перестановка одной записи
bad = GOLD[:]; bad[2], bad[7] = bad[7], bad[2]
rs = rules(doc(bad))
ok(has(rs, 'ORD1'), 'ORD1: перестановка')
# 10. по порядку ссылок (случайный)
random.seed(1); sh = GOLD[:]; random.shuffle(sh)
rs = rules(doc(sh))
ok(any(r == 'ORD1' and 'вероятно' in m for r, m in rs) or sum(r == 'ORD1' for r, m in rs) >= 3, 'ORD1: перемешанный список')
# 11. иностранные перед русскими
mix = EN[:2] + RU + EN[2:]
rs = rules(doc(mix))
ok(has(rs, 'ORD2'), 'ORD2: иностранные раньше русских')
# 12. дубль
dup = GOLD[:]; dup.insert(1, GOLD[0])
rs = rules(doc(sorted(dup, key=lambda s: B.sort_key(s) if B.is_cyr(s) else [(2, s)])))
ok(has(rs, 'DUP'), 'DUP: повтор записи')
# 13. пропуск номера (ручная нумерация)
ps = [para('ВВЕДЕНИЕ', center=True), para('Текст [1].'), para('СПИСОК ЛИТЕРАТУРЫ', center=True)]
for i, r in enumerate(GOLD, 1):
    if i == 5: continue
    ps.append(para('%d. %s' % (i, r)))
rs = rules(build(ps, numbering=False))
ok(has(rs, 'LS2n'), 'LS2n: пропущен номер 5')
# 14. оформление: нет года, нет дат обращения, Электронный ресурс, дефис
def with_record(i, new):
    r = GOLD[:]; r[i] = new; return r
rs = rules(doc(with_record(3, "Громова, О. Н. Лексика в контексте обучения // Иностранные языки в школе. – № 2. – С. 3–10.")))
ok(any(r == 'BIB1' for r, m in rs), 'BIB1: нет года')
rs = rules(doc(with_record(10, "Министерство просвещения Российской Федерации [Официальный сайт]. – URL: https://edu.gov.ru. – Текст : электронный.")))
ok(any(r == 'BIB3' for r, m in rs), 'BIB3: нет даты обращения')
rs = rules(doc(with_record(10, "Министерство просвещения [Электронный ресурс]. – Режим доступа: http://edu.gov.ru (дата обращения 14.04.2023).")))
ok(any(r == 'BIB4' for r, m in rs) and any(r == 'BIB3' for r, m in rs), 'BIB4/BIB3: старая форма электронного ресурса')
rs = rules(doc(with_record(4, "Данилов, С. В. Теория и практика перевода. – Москва : Юрайт, 2017.")))
ok(any(r == 'BIB5' for r, m in rs), 'BIB5: нет объёма')
rs = rules(doc(with_record(1, "Борисова, Е. И. Цифровые ресурсы микрообучения // Вестник образования. – 2019. – № 4.")))
ok(any(r == 'BIB2' for r, m in rs), 'BIB2: у статьи нет страниц')
rs = rules(doc(with_record(1, "Борисова, Е. И. Цифровые ресурсы микрообучения//Вестник образования. – 2019. – № 4. – С. 25–29.")))
ok(any(r == 'BIB2' for r, m in rs), 'BIB2: // без пробелов')
# 15. список без номеров
rs = rules(build([para('ВВЕДЕНИЕ', center=True), para('Текст [1].'), para('СПИСОК ЛИТЕРАТУРЫ', center=True)] + [para(r) for r in GOLD], numbering=False))
ok(has(rs, 'LS2n') and not has(rs, 'L1'), 'LS2n: записи без номеров, но 18 записей найдено ' + str([r for r, m in rs]))
# 16. список не найден
rs = rules(build([para('ВВЕДЕНИЕ', center=True), para('Текст.')]))
ok(has(rs, 'L0'), 'L0: списка нет')
# 17. подзаголовки в списке не считаются записями
ps = [para('ВВЕДЕНИЕ', center=True), para('Текст [1].'), para('СПИСОК ЛИТЕРАТУРЫ', center=True), para('Нормативные документы')]
ps += [para('%d. %s' % (i, r)) for i, r in enumerate(GOLD[9:10], 1)]
ps += [para('Научная литература')] + [para('%d. %s' % (i, r)) for i, r in enumerate(GOLD[:9], 2)]
s = B.summary(build(ps, numbering=False))
ok(s.get('n') == 10, 'подзаголовки пропущены: записей %s' % s.get('n'))
# 18. приложение закрывает список
ps = [para('ВВЕДЕНИЕ', center=True), para('Текст [1].'), para('СПИСОК ЛИТЕРАТУРЫ', center=True)] + [para('%d. %s' % (i, r)) for i, r in enumerate(GOLD, 1)] + [para('ПРИЛОЖЕНИЕ А'), para('Таблица без номера и какой-то текст в приложении')]
s = B.summary(build(ps, numbering=False))
ok(s.get('n') == 18, 'приложение не попало в список: %s' % s.get('n'))

print('\nFAILED:' if fails else '\nALL OK', fails)
sys.exit(1 if fails else 0)
