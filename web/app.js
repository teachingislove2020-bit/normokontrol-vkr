(function () {
  'use strict';
  const $ = (id) => document.getElementById(id);
  const fileInput = $('file'), runBtn = $('run'), statusEl = $('status');
  let file = null, ready = false, busy = false, last = null, filter = 'all';
  let state = {};            // id замечания → { d: 'fixed' | 'keep', clause, reason, touched }
  let byId = {};             // id → { g, it }
  let dirty = false;         // есть решения, которые не вошли в скачанный лист
  let sheetBusy = false;
  const openGroups = new Set();

  // Те же пороги, что в normokontrol_sheet.py (там они проверяются повторно)
  const MIN_CLAUSE = 3, MAX_CLAUSE = 120, MIN_REASON = 15, MAX_REASON = 1000;
  const norm = (s) => (s || '').replace(/\s+/g, ' ').trim();
  const len = (s) => Array.from(s).length;
  const clauseOk = (s) => { const c = norm(s); return len(c) >= MIN_CLAUSE && len(c) <= MAX_CLAUSE && /\d/.test(c); };
  const reasonOk = (s) => { const r = norm(s); return len(r) >= MIN_REASON && len(r) <= MAX_REASON; };

  const worker = new Worker('worker.js', { type: 'module' });
  worker.onmessage = (ev) => {
    const { type, payload } = ev.data;
    if (type === 'status') setStatus(payload);
    else if (type === 'ready') { ready = true; setStatus(file ? 'Движок готов. Нажмите «Проверить».' : 'Движок готов. Выберите файл.'); update(); }
    else if (type === 'fatal') { setStatus('Не удалось запустить проверку: ' + payload + '. Обновите страницу или попробуйте другой браузер (Chrome, Edge, Firefox, Safari последних версий).', true); }
    else if (type === 'result') { busy = false; update(); show(JSON.parse(payload)); }
    else if (type === 'error') { busy = false; update(); setStatus('Ошибка при проверке: ' + payload, true); }
    else if (type === 'sheet') onSheet(JSON.parse(payload));
  };
  worker.onerror = (e) => setStatus('Ошибка запуска: ' + (e.message || 'неизвестная'), true);

  function setStatus(t, err) { statusEl.textContent = t || ''; statusEl.className = 'status' + (err ? ' err' : ''); }
  function update() { runBtn.disabled = !(file && ready && !busy); }

  function pick(f) {
    if (!f) return;
    if (!/\.docx$/i.test(f.name)) {
      file = null;
      setStatus(/\.doc$/i.test(f.name) ? 'Формат .doc не поддерживается: откройте файл в Word и сохраните как .docx.' : 'Нужен файл .docx.', true);
      update(); return;
    }
    if (f.size > 120 * 1024 * 1024) { file = null; setStatus('Файл больше 120 МБ. Уменьшите картинки в документе и повторите.', true); update(); return; }
    file = f;
    $('drop-title').textContent = f.name;
    $('drop-sub').textContent = (f.size / 1048576).toFixed(1) + ' МБ';
    setStatus(ready ? 'Нажмите «Проверить».' : statusEl.textContent);
    update();
  }
  fileInput.addEventListener('change', () => pick(fileInput.files[0]));
  const drop = $('drop');
  ['dragenter', 'dragover'].forEach(n => drop.addEventListener(n, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(n => drop.addEventListener(n, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => { if (e.dataTransfer.files[0]) pick(e.dataTransfer.files[0]); });

  runBtn.addEventListener('click', async () => {
    if (!file || !ready || busy) return;
    if (dirty && !window.confirm('Принятые решения не вошли в скачанный лист и будут потеряны. Проверить файл заново?')) return;
    busy = true; update(); setStatus('Читаю файл…');
    $('report').hidden = true;
    try {
      const buf = await file.arrayBuffer();
      worker.postMessage({ type: 'check', data: buf, cfg: {} }, [buf]);
    } catch (e) { busy = false; update(); setStatus('Не удалось прочитать файл: ' + e.message, true); }
  });

  function el(tag, cls, text) { const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; }
  const LEVEL = { fix: 'Исправьте', check: 'Проверьте' };
  const offset = () => Math.max(0, parseInt($('offset').value, 10) || 0);

  // ---------- решения ----------
  function status(id) {                          // 'none' | 'bad' | 'ok'
    const s = state[id];
    if (!s || !s.d) return 'none';
    if (s.d === 'fixed') return 'ok';
    return clauseOk(s.clause) && reasonOk(s.reason) ? 'ok' : 'bad';
  }
  function counts() {
    let none = 0, bad = 0, ok = 0, total = 0;
    last.groups.forEach(g => g.items.forEach(it => { total++; const s = status(it.id); if (s === 'ok') ok++; else if (s === 'bad') bad++; else none++; }));
    return { total, ok, bad, none };
  }
  function groupDone(g) { return g.items.reduce((n, it) => n + (status(it.id) === 'ok' ? 1 : 0), 0); }

  function setDecision(id, patch) {
    state[id] = Object.assign({}, state[id], patch);
    dirty = true;
    refresh();
  }

  function refresh() {
    if (!last) return;
    const c = counts();
    $('sb-count').textContent = 'Решения приняты: ' + c.ok + ' из ' + c.total;
    const parts = [];
    if (c.none) parts.push('не выбрано решение: ' + c.none);
    if (c.bad) parts.push('не заполнены пункт или причина: ' + c.bad);
    const hint = c.total === 0 ? 'замечаний нет, лист можно скачать' : parts.length ? parts.join(' · ') : 'лист можно скачать';
    $('sb-hint').textContent = hint.charAt(0).toUpperCase() + hint.slice(1);
    $('sheet-dl').disabled = sheetBusy || c.none > 0 || c.bad > 0;
    $('goto').hidden = !(c.none || c.bad);
    last.groups.forEach(g => {
      const e = document.querySelector('[data-gdone="' + g.gid + '"]');
      if (e) { const d = groupDone(g); e.textContent = d === g.items.length ? 'решено' : d + ' из ' + g.items.length; e.className = 'gdone' + (d === g.items.length ? ' all' : ''); }
    });
  }

  // ---------- вывод ----------
  function where(item, pagesKnown) {
    const parts = [];
    if (pagesKnown && item.page != null) parts.push('стр. ' + (item.page + offset()));
    if (item.head) parts.push('раздел «' + item.head + '»');
    return parts.join(' · ');
  }

  function decideBlock(g, it) {
    const id = it.id;
    const box = el('div', 'decide');
    const row = el('div', 'dbtns');
    const bFix = el('button', 'dbtn', 'Исправил(а)'); bFix.type = 'button';
    const bKeep = el('button', 'dbtn', g.level === 'fix' ? 'Не согласен(на), оставляю' : 'Оставляю'); bKeep.type = 'button';
    row.appendChild(bFix); row.appendChild(bKeep);
    const form = el('div', 'dform');
    const l1 = el('label', null, 'Пункт ГОСТ или методички');
    const clause = el('input'); clause.type = 'text'; clause.maxLength = MAX_CLAUSE; clause.autocomplete = 'off';
    clause.placeholder = 'например: ГОСТ 7.32-2017, п. 6.2.3';
    l1.appendChild(clause);
    const e1 = el('div', 'derr');
    const l2 = el('label', null, 'Почему оставляете');
    const reason = el('textarea'); reason.rows = 3; reason.maxLength = MAX_REASON;
    reason.placeholder = 'Кратко: почему требование здесь не применяется или чем оно заменено. Без ФИО и контактов.';
    l2.appendChild(reason);
    const e2 = el('div', 'derr');
    form.appendChild(l1); form.appendChild(e1); form.appendChild(l2); form.appendChild(e2);
    box.appendChild(row); box.appendChild(form);

    function sync() {
      const s = state[id] || {};
      bFix.setAttribute('aria-pressed', s.d === 'fixed' ? 'true' : 'false');
      bKeep.setAttribute('aria-pressed', s.d === 'keep' ? 'true' : 'false');
      form.hidden = s.d !== 'keep';
      if (s.d === 'keep') {
        if (document.activeElement !== clause) clause.value = s.clause || '';
        if (document.activeElement !== reason) reason.value = s.reason || '';
        e1.textContent = s.touched && !clauseOk(s.clause) ? 'Укажите пункт с номером, например «п. 6.2.3».' : '';
        e2.textContent = s.touched && !reasonOk(s.reason) ? 'Опишите причину не короче ' + MIN_REASON + ' знаков.' : '';
      }
      const st = status(id);
      box.parentNode && box.parentNode.classList.toggle('done', st === 'ok');
      box.parentNode && box.parentNode.classList.toggle('bad', st === 'bad');
    }
    bFix.addEventListener('click', () => { setDecision(id, { d: 'fixed' }); sync(); });
    bKeep.addEventListener('click', () => { setDecision(id, { d: 'keep', clause: (state[id] || {}).clause || '', reason: (state[id] || {}).reason || '', touched: false }); sync(); clause.focus(); });
    clause.addEventListener('input', () => { setDecision(id, { clause: clause.value }); sync(); });
    reason.addEventListener('input', () => { setDecision(id, { reason: reason.value }); sync(); });
    [clause, reason].forEach(x => x.addEventListener('blur', () => { setDecision(id, { touched: true }); sync(); }));
    // внешнее состояние (групповые действия) обновляется перерисовкой группы, поэтому sync вызывается после вставки в li
    box.sync = sync;
    return box;
  }

  function renderItem(g, it, pagesKnown) {
    const li = el('li'); li.dataset.id = it.id;
    li.appendChild(el('div', 'msg', it.msg));
    const w = where(it, pagesKnown);
    if (w) li.appendChild(el('div', 'where', w));
    if (it.quote) li.appendChild(el('div', 'quote', '«' + it.quote + '»'));
    if (it.examples && it.examples.length) {
      const ul = el('ul', 'ex');
      it.examples.forEach(x => {
        const l = el('li');
        const w2 = where(x, pagesKnown);
        l.appendChild(el('span', 'quote', '«' + x.quote + '»'));
        if (w2) l.appendChild(el('span', 'where', ' — ' + w2));
        ul.appendChild(l);
      });
      li.appendChild(el('div', 'where', 'Примеры:'));
      li.appendChild(ul);
    }
    const d = decideBlock(g, it);
    li.appendChild(d);
    d.sync();
    return li;
  }

  function bulkBar(g) {
    const bar = el('div', 'bulk');
    const all = el('button', 'ghost', 'Отметить все «исправил(а)»'); all.type = 'button';
    all.title = 'Применяется к замечаниям этого вида, по которым решение ещё не выбрано';
    all.addEventListener('click', () => {
      g.items.forEach(it => { if (status(it.id) === 'none') state[it.id] = { d: 'fixed' }; });
      dirty = true; buildBody(g); refresh();
    });
    const reset = el('button', 'ghost', 'Сбросить решения в этом виде'); reset.type = 'button';
    reset.addEventListener('click', () => { g.items.forEach(it => { delete state[it.id]; }); buildBody(g); refresh(); });
    bar.appendChild(all); bar.appendChild(reset);

    const det = el('details', 'bulkkeep');
    det.appendChild(el('summary', null, 'Оставить все с одной причиной'));
    const l1 = el('label', null, 'Пункт ГОСТ или методички'); const c = el('input'); c.type = 'text'; c.maxLength = MAX_CLAUSE; l1.appendChild(c);
    const l2 = el('label', null, 'Почему оставляете'); const r = el('textarea'); r.rows = 2; r.maxLength = MAX_REASON; l2.appendChild(r);
    const err = el('div', 'derr');
    const go = el('button', null, 'Применить к замечаниям без решения'); go.type = 'button';
    go.addEventListener('click', () => {
      if (!clauseOk(c.value) || !reasonOk(r.value)) { err.textContent = 'Укажите пункт с номером и причину не короче ' + MIN_REASON + ' знаков.'; return; }
      g.items.forEach(it => { if (status(it.id) === 'none') state[it.id] = { d: 'keep', clause: c.value, reason: r.value, touched: true }; });
      dirty = true; buildBody(g); refresh();
    });
    det.appendChild(l1); det.appendChild(l2); det.appendChild(err); det.appendChild(go);
    bar.appendChild(det);
    return bar;
  }

  function buildBody(g) {
    const b = document.querySelector('[data-body="' + g.gid + '"]');
    if (!b) return;
    b.textContent = '';
    if (g.source) { const p = el('p'); p.appendChild(el('strong', null, 'Требование: ')); p.appendChild(document.createTextNode(g.source)); b.appendChild(p); }
    if (g.action) { const p = el('p'); p.appendChild(el('strong', null, 'Что сделать: ')); p.appendChild(document.createTextNode(g.action)); b.appendChild(p); }
    b.appendChild(bulkBar(g));
    const ul = el('ul', 'items');
    g.items.forEach(it => { if (filter === 'todo' && status(it.id) === 'ok') return; ul.appendChild(renderItem(g, it, last.pages_known)); });
    if (!ul.children.length) b.appendChild(el('p', 'empty', 'Все замечания этого вида решены.'));
    else b.appendChild(ul);
    b.dataset.built = '1';
  }

  function renderGroups(r) {
    const box = $('groups'); box.textContent = '';
    r.groups.forEach(g => {
      if ((filter === 'fix' || filter === 'check') && g.level !== filter) return;
      if (filter === 'todo' && groupDone(g) === g.items.length) return;
      const d = el('details', 'group ' + g.level);
      d.open = openGroups.has(g.gid);
      const s = el('summary');
      s.appendChild(el('span', 'badge ' + g.level, LEVEL[g.level] || g.level));
      s.appendChild(el('span', 'gtitle', g.title));
      s.appendChild(el('span', 'gid', g.rule + (g.sub ? '·' + g.sub : '')));
      const dn = el('span', 'gdone'); dn.dataset.gdone = g.gid; s.appendChild(dn);
      s.appendChild(el('span', 'gcount', String(g.count)));
      d.appendChild(s);
      const b = el('div', 'gbody'); b.dataset.body = g.gid;
      d.appendChild(b);
      d.addEventListener('toggle', () => {
        if (d.open) { openGroups.add(g.gid); if (!b.dataset.built) buildBody(g); } else openGroups.delete(g.gid);
      });
      box.appendChild(d);
      if (d.open) buildBody(g);
    });
    if (!box.children.length) box.appendChild(el('p', 'empty', filter === 'all' ? 'Замечаний не найдено.' : filter === 'todo' ? 'Решения приняты по всем замечаниям.' : 'В этой категории замечаний нет.'));
    refresh();
  }

  function renderAllBodies() {            // для печати и сохранения отчёта: раскрыть и построить все группы
    document.querySelectorAll('#groups details').forEach(d => { d.open = true; });
    last.groups.forEach(g => { const b = document.querySelector('[data-body="' + g.gid + '"]'); if (b && !b.dataset.built) buildBody(g); });
  }

  function show(r) {
    if (r.error) { setStatus(r.error, true); return; }
    last = r; state = {}; byId = {}; dirty = false; openGroups.clear(); filter = 'all';
    document.querySelectorAll('.tab').forEach(x => x.classList.toggle('active', x.dataset.f === 'all'));
    r.groups.forEach(g => { if (g.level === 'fix') openGroups.add(g.gid); g.items.forEach(it => { byId[it.id] = { g, it }; }); });
    setStatus(''); $('sheet-status').textContent = '';
    const s = r.stats;
    const sum = $('summary'); sum.textContent = '';
    const total = s.fix_items + s.check_items;
    if (!total) {
      sum.appendChild(el('h2', null, 'Замечаний не найдено'));
      sum.appendChild(el('p', null, 'Это не отменяет нормоконтроль: проверьте разделы «Проверьте сами» ниже. Лист самопроверки можно скачать и так.'));
    } else {
      sum.appendChild(el('h2', null, 'Результат проверки'));
      const chips = el('div', 'chips');
      chips.appendChild(el('span', 'chip fix', 'Исправьте: ' + s.fix_items + ' (видов: ' + s.fix_groups + ')'));
      chips.appendChild(el('span', 'chip check', 'Проверьте: ' + s.check_items + ' (видов: ' + s.check_groups + ')'));
      sum.appendChild(chips);
      sum.appendChild(el('p', 'hint', 'По каждому замечанию выберите «Исправил(а)» или «Оставляю». Для «Оставляю» нужны пункт ГОСТ или методички и причина. Когда решения приняты по всем замечаниям, внизу страницы станет доступна кнопка «Скачать лист».'));
      if (!r.pages_known) sum.appendChild(el('p', 'hint', 'В файле нет меток страниц Word (он сохранён не в Word), поэтому вместо страниц показаны разделы и фрагменты текста.'));
    }
    const sl = $('self-list'); sl.textContent = '';
    r.self_checks.forEach(c => { const li = el('li'); li.appendChild(el('strong', null, c.title + '. ')); li.appendChild(document.createTextNode(c.text)); sl.appendChild(li); });
    renderGroups(r);
    $('report').hidden = false;
    $('report').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  document.querySelectorAll('.tab').forEach(t => t.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(x => x.classList.toggle('active', x === t));
    filter = t.dataset.f; if (last) renderGroups(last);
  }));
  $('offset').addEventListener('input', () => { if (last) renderGroups(last); });

  // ---------- лист самопроверки ----------
  $('goto').addEventListener('click', () => {
    if (!last) return;
    let target = null;
    for (const g of last.groups) { const it = g.items.find(x => status(x.id) !== 'ok'); if (it) { target = { g, it }; break; } }
    if (!target) return;
    filter = 'all';
    document.querySelectorAll('.tab').forEach(x => x.classList.toggle('active', x.dataset.f === 'all'));
    openGroups.add(target.g.gid);
    renderGroups(last);
    const li = document.querySelector('li[data-id="' + target.it.id + '"]');
    if (li) { li.scrollIntoView({ behavior: 'smooth', block: 'center' }); const b = li.querySelector('.dbtn'); if (b) b.focus({ preventScroll: true }); }
  });

  function today() {
    const d = new Date(), p = (n) => String(n).padStart(2, '0');
    return { ru: p(d.getDate()) + '.' + p(d.getMonth() + 1) + '.' + d.getFullYear(), iso: d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) };
  }

  $('sheet-dl').addEventListener('click', () => {
    if (!last || sheetBusy) return;
    const c = counts();
    if (c.none || c.bad) return;
    sheetBusy = true; refresh();
    $('sheet-status').textContent = 'Формирую лист…'; $('sheet-status').className = 'status';
    const t = today();
    const decisions = {};
    Object.keys(state).forEach(id => { const s = state[id]; if (s && s.d) decisions[id] = { d: s.d, clause: s.clause || '', reason: s.reason || '' }; });
    const payload = {
      version: last.version, date_ru: t.ru, date_iso: t.iso, offset: offset(), pages_known: !!last.pages_known,
      include_quotes: !$('noquotes').checked, self_checks: last.self_checks.map(x => x.title), decisions,
      groups: last.groups.map(g => ({ rule: g.rule, sub: g.sub, title: g.title, level: g.level, source: g.source,
        items: g.items.map(it => ({ id: it.id, msg: it.msg, quote: it.quote, head: it.head, page: it.page, examples: it.examples })) }))
    };
    worker.postMessage({ type: 'sheet', payload: JSON.stringify(payload) });
  });

  function onSheet(r) {
    sheetBusy = false;
    const st = $('sheet-status');
    if (!r.ok) {
      st.textContent = r.error || 'Лист не сформирован: ' + (r.undecided && r.undecided.length ? 'не по всем замечаниям выбрано решение. ' : '') + (r.bad_keep && r.bad_keep.length ? 'у части замечаний нет пункта или причины.' : '');
      st.className = 'status err'; refresh(); return;
    }
    const bin = atob(r.b64), bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }));
    a.download = 'list-samoproverki-vkr.docx';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    dirty = false;
    st.textContent = 'Лист скачан: list-samoproverki-vkr.docx. Приложите его к работе.'; st.className = 'status';
    refresh();
  }

  window.addEventListener('beforeunload', (e) => { if (dirty) { e.preventDefault(); e.returnValue = ''; } });

  // ---------- печать и сохранение отчёта ----------
  $('print').addEventListener('click', () => window.print());
  window.addEventListener('beforeprint', () => { if (last) renderAllBodies(); });

  $('save').addEventListener('click', async () => {
    if (!last) return;
    let css = '';
    try { css = await (await fetch('style.css')).text(); } catch (_) {}
    renderAllBodies();
    const clone = $('groups').cloneNode(true);
    clone.querySelectorAll('.decide,.bulk').forEach(x => x.remove());
    const body = $('summary').outerHTML + clone.outerHTML + $('self').outerHTML;
    const html = '<!doctype html><html lang="ru"><head><meta charset="utf-8"><title>Отчёт самопроверки ВКР</title><style>' + css + '</style></head><body><main><h1>Отчёт самопроверки ВКР</h1><p class="hint">' + new Date().toLocaleString('ru-RU') + ' · версия ' + last.version + '</p>' + body + '</main></body></html>';
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([html], { type: 'text/html;charset=utf-8' }));
    a.download = 'otchet-samoproverka-vkr.html';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  });
})();
