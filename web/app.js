(function () {
  'use strict';
  const $ = (id) => document.getElementById(id);
  const fileInput = $('file'), runBtn = $('run'), statusEl = $('status');
  let file = null, ready = false, busy = false, last = null, filter = 'all';

  const worker = new Worker('worker.js', { type: 'module' });
  worker.onmessage = (ev) => {
    const { type, payload } = ev.data;
    if (type === 'status') setStatus(payload);
    else if (type === 'ready') { ready = true; setStatus(file ? 'Движок готов. Нажмите «Проверить».' : 'Движок готов. Выберите файл.'); update(); }
    else if (type === 'fatal') { setStatus('Не удалось запустить проверку: ' + payload + '. Обновите страницу или попробуйте другой браузер (Chrome, Edge, Firefox, Safari последних версий).', true); }
    else if (type === 'result') { busy = false; update(); show(JSON.parse(payload)); }
    else if (type === 'error') { busy = false; update(); setStatus('Ошибка при проверке: ' + payload, true); }
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
    busy = true; update(); setStatus('Читаю файл…');
    $('report').hidden = true;
    try {
      const buf = await file.arrayBuffer();
      worker.postMessage({ type: 'check', data: buf, cfg: {} }, [buf]);
    } catch (e) { busy = false; update(); setStatus('Не удалось прочитать файл: ' + e.message, true); }
  });

  function el(tag, cls, text) { const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; }
  const LEVEL = { fix: 'Исправьте', check: 'Проверьте' };

  function where(item, offset, pagesKnown) {
    const parts = [];
    if (pagesKnown && item.page != null) parts.push('стр. ' + (item.page + offset));
    if (item.head) parts.push('раздел «' + item.head + '»');
    return parts.join(' · ');
  }

  function renderItem(it, offset, pagesKnown) {
    const li = el('li');
    li.appendChild(el('div', 'msg', it.msg));
    const w = where(it, offset, pagesKnown);
    if (w) li.appendChild(el('div', 'where', w));
    if (it.quote) li.appendChild(el('div', 'quote', '«' + it.quote + '»'));
    if (it.examples && it.examples.length) {
      const ul = el('ul', 'ex');
      it.examples.forEach(x => {
        const l = el('li');
        const w2 = where(x, offset, pagesKnown);
        l.appendChild(el('span', 'quote', '«' + x.quote + '»'));
        if (w2) l.appendChild(el('span', 'where', ' — ' + w2));
        ul.appendChild(l);
      });
      li.appendChild(el('div', 'where', 'Примеры:'));
      li.appendChild(ul);
    }
    return li;
  }

  function renderGroups(r) {
    const offset = Math.max(0, parseInt($('offset').value, 10) || 0);
    const box = $('groups'); box.textContent = '';
    r.groups.forEach(g => {
      if (filter !== 'all' && g.level !== filter) return;
      const d = el('details', 'group ' + g.level);
      if (g.level === 'fix') d.open = true;
      const s = el('summary');
      s.appendChild(el('span', 'badge ' + g.level, LEVEL[g.level] || g.level));
      s.appendChild(el('span', 'gtitle', g.title));
      s.appendChild(el('span', 'gid', g.rule + (g.sub ? '·' + g.sub : '')));
      s.appendChild(el('span', 'gcount', String(g.count)));
      d.appendChild(s);
      const b = el('div', 'gbody');
      if (g.source) { const p = el('p'); p.appendChild(el('strong', null, 'Требование: ')); p.appendChild(document.createTextNode(g.source)); b.appendChild(p); }
      if (g.action) { const p = el('p'); p.appendChild(el('strong', null, 'Что сделать: ')); p.appendChild(document.createTextNode(g.action)); b.appendChild(p); }
      const ul = el('ul', 'items');
      g.items.forEach(it => ul.appendChild(renderItem(it, offset, r.pages_known)));
      b.appendChild(ul);
      if (g.more > 0) b.appendChild(el('p', 'more', 'и ещё ' + g.more + ' замечаний этого вида'));
      d.appendChild(b);
      box.appendChild(d);
    });
    if (!box.children.length) box.appendChild(el('p', 'empty', filter === 'all' ? 'Замечаний не найдено.' : 'В этой категории замечаний нет.'));
  }

  function show(r) {
    if (r.error) { setStatus(r.error, true); return; }
    last = r;
    setStatus('');
    const s = r.stats;
    const sum = $('summary'); sum.textContent = '';
    const total = s.fix_items + s.check_items;
    if (!total) {
      sum.appendChild(el('h2', null, 'Замечаний не найдено'));
      sum.appendChild(el('p', null, 'Это не отменяет нормоконтроль: проверьте разделы «Проверьте сами» ниже.'));
    } else {
      sum.appendChild(el('h2', null, 'Результат проверки'));
      const chips = el('div', 'chips');
      chips.appendChild(el('span', 'chip fix', 'Исправьте: ' + s.fix_items + ' (видов: ' + s.fix_groups + ')'));
      chips.appendChild(el('span', 'chip check', 'Проверьте: ' + s.check_items + ' (видов: ' + s.check_groups + ')'));
      sum.appendChild(chips);
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

  $('print').addEventListener('click', () => window.print());
  window.addEventListener('beforeprint', () => document.querySelectorAll('details').forEach(d => d.open = true));

  $('save').addEventListener('click', async () => {
    if (!last) return;
    let css = '';
    try { css = await (await fetch('style.css')).text(); } catch (_) {}
    const prevFilter = filter; filter = 'all'; renderGroups(last);
    document.querySelectorAll('#groups details').forEach(d => d.setAttribute('open', ''));
    const body = $('summary').outerHTML + $('groups').outerHTML + $('self').outerHTML;
    filter = prevFilter; renderGroups(last);
    const name = file ? file.name : 'документ';
    const html = '<!doctype html><html lang="ru"><head><meta charset="utf-8"><title>Отчёт самопроверки ВКР</title><style>' + css + '</style></head><body><main><h1>Отчёт самопроверки ВКР</h1><p class="hint">Файл: ' + name.replace(/[<>&]/g, '') + ' · ' + new Date().toLocaleString('ru-RU') + ' · версия ' + last.version + '</p>' + body + '</main></body></html>';
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([html], { type: 'text/html;charset=utf-8' }));
    a.download = 'otchet-samoproverka-vkr.html';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  });
})();
