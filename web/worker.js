/* Web Worker: Pyodide + lxml + модули проверки. Все файлы берутся с этого же сайта. */
import { loadPyodide } from './pyodide/pyodide.mjs';

const BASE = new URL('./', self.location.href).href;
const V = '0.9';
let py = null;
const ready = (async () => {
  post('status', 'Загружаю движок проверки (около 16 МБ, один раз)…');
  py = await loadPyodide({ indexURL: BASE + 'pyodide/' });
  post('status', 'Подключаю разбор документов…');
  await py.loadPackage('lxml');
  py.FS.mkdirTree('/app');
  for (const f of ['normokontrol_core.py', 'normokontrol_biblio.py', 'normokontrol_service.py']) {
    const r = await fetch(BASE + 'py/' + f + '?v=' + V);
    if (!r.ok) throw new Error('Не загрузился ' + f);
    py.FS.writeFile('/app/' + f, await r.text());
  }
  const rules = await (await fetch(BASE + 'rules.json?v=' + V)).text();
  py.FS.writeFile('/app/rules.json', rules);
  py.runPython("import sys; sys.path.insert(0, '/app')\nimport normokontrol_service as S\nS.load_rules(open('/app/rules.json', encoding='utf8').read())");
  post('ready', null);
})();
ready.catch(e => post('fatal', String(e && e.message || e)));

function post(type, payload) { self.postMessage({ type, payload }); }

self.onmessage = async (ev) => {
  const m = ev.data;
  if (m.type !== 'check') return;
  try {
    await ready;
    post('status', 'Проверяю документ…');
    py.FS.writeFile('/in.docx', new Uint8Array(m.data));
    const res = py.runPython("import json\nS.run(open('/in.docx', 'rb').read(), " + JSON.stringify(JSON.stringify(m.cfg || {})) + ")");
    try { py.FS.unlink('/in.docx'); } catch (_) {}
    post('result', res);
  } catch (e) {
    post('error', String(e && e.message || e));
  }
};
