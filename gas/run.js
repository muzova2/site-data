// 24.11.2026: сбор данных для страниц bulkacapital.ru на GitHub.
// Мосбиржа перешла на сертификат Минцифры, и Google к ней больше не подключается. Поэтому здесь запускается
// тот же код Google Apps Script (gas/*.gs — копия «Календарь дивидендов → Код.gs»), а сервисы Google заменены:
//   таблица   — листы в памяти: ручные и входящие листы читаются из веб-приложения Google (?d=sheet),
//               накопительные хранятся в папке state (ветка state репозитория);
//   запросы   — curl (параллельно, синхронно для кода);
//   триггеры  — «догрузить через минуту» выполняется сразу же, в этом же запуске;
//   готовый JSON — сразу в data/<ключ>.json.
// Запуск: node gas/run.js bonds coupons dividends curve indices shares funds fx futures options dohod
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm'), cp = require('child_process'), crypto = require('crypto'), os = require('os');

const EXEC = process.env.GAS_EXEC || 'https://script.google.com/macros/s/AKfycbyQammVyKFaW1iym9aCJV7lxRWvePm8T4xx5wkstlF56SHBSVNlfsji4BX7AXXq_YAG/exec';
const STATE = process.env.STATE_DIR || 'state';
const DATA = process.env.DATA_DIR || 'data';
const GAS_DIR = path.join(__dirname);
// листы, которые ведёт человек или сборщик дивидендов (он присылает данные в Google): каждый раз свежие из Google
const INPUT = ['Секторы вручную', 'Дивиденды (сырые)', 'Дивиденды: компании', 'Показатели акций', 'Рейтинги', 'Доходность: недвижимость'];
// листы, которые код накапливает сам: берём из state, при первом запуске — из Google
const KEEP = ['Компании', 'Выплаты', 'Облигации', 'Облигации: справочник', 'Кривая: история', 'Фонды: справочник', 'Фонды: СЧА с сайтов УК'];
const PROPS_FROM_GOOGLE = ['FUND_UPDATED'];   // его ставит приём показателей акций в Google

// ---------- HTTP через curl ----------
function curlMany(reqs) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'gasf-')), cfg = [];
  reqs.forEach((r, i) => {
    if (i) cfg.push('next');
    cfg.push('url = ' + JSON.stringify(r.url));
    cfg.push('output = ' + JSON.stringify(path.join(dir, String(i))));
    const h = Object.assign({ 'User-Agent': 'Mozilla/5.0' }, r.headers || {});
    for (const k of Object.keys(h)) cfg.push('header = ' + JSON.stringify(k + ': ' + h[k]));
    if (r.followRedirects !== false) cfg.push('location');
    // после «next» curl сбрасывает параметры запроса, поэтому они повторяются у каждого
    cfg.push('compressed', 'max-time = 90', 'retry = 2', 'write-out = "%{response_code} %{filename_effective}\\n"');
  });
  fs.writeFileSync(path.join(dir, 'cfg'), cfg.join('\n') + '\n');
  let out = '';
  try {
    out = cp.execFileSync('curl', ['-s', '--parallel', '--parallel-max', '12', '-K', path.join(dir, 'cfg')], { maxBuffer: 1 << 26, encoding: 'utf8' });
  } catch (e) { out = String(e.stdout || ''); }
  const code = {};
  out.split(/\r?\n/).forEach(l => { const m = l.match(/^(\d+) (.*)$/); if (m) code[path.basename(m[2])] = Number(m[1]); });
  const res = reqs.map((r, i) => {
    const f = path.join(dir, String(i));
    return new HTTPResponse(code[String(i)] || 0, fs.existsSync(f) ? fs.readFileSync(f) : Buffer.alloc(0));
  });
  fs.rmSync(dir, { recursive: true, force: true });
  if (process.env.GAS_DEBUG) console.log('  curl', reqs.length, 'запросов, коды:', res.map(r => r.code).join(','), reqs[0].url.slice(0, 120));
  return res;
}
class HTTPResponse {
  constructor(code, buf) { this.code = code; this.buf = buf; }
  getResponseCode() { return this.code; }
  getContentText(cs) { return new TextDecoder(cs || 'utf-8').decode(this.buf); }
  getBlob() { const b = this.buf; return { getBytes: () => Array.from(b), getDataAsString: cs => new TextDecoder(cs || 'utf-8').decode(b) }; }
}
function toReq(u, o) { return typeof u === 'string' ? Object.assign({ url: u }, o || {}) : u; }
// запросы, не получившие ответа (0, 429, 5xx), повторяем ещё дважды с паузой
function curlRetry(reqs) {
  const res = curlMany(reqs);
  for (let attempt = 0; attempt < 2; attempt++) {
    const bad = res.map((r, i) => (r.code === 0 || r.code === 429 || r.code >= 500 ? i : -1)).filter(i => i >= 0);
    if (!bad.length) break;
    Utilities.sleep(3000 * (attempt + 1));
    const again = curlMany(bad.map(i => reqs[i]));
    bad.forEach((i, k) => { res[i] = again[k]; });
  }
  return res;
}
const UrlFetchApp = {
  fetch(url, opts) {
    const r = curlRetry([toReq(url, opts)])[0];
    if (!(opts && opts.muteHttpExceptions) && !(r.code >= 200 && r.code < 300)) throw new Error('Request failed for ' + url + ' returned code ' + r.code);
    return r;
  },
  fetchAll(reqs) { return curlRetry(reqs.map(r => toReq(r))); },
};

// ---------- таблица ----------
const revive = v => (v && typeof v === 'object' && v.$d ? new Date(v.$d) : v);
const pack = v => (v instanceof Date ? { $d: v.toISOString() } : v);
class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr: nr || 1, nc: nc || 1 }); }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const row = this.sh.rows[this.r - 1 + i] || [], o = [];
      for (let j = 0; j < this.nc; j++) { const v = row[this.c - 1 + j]; o.push(v === undefined || v === null ? '' : v); }
      out.push(o);
    }
    return out;
  }
  getValue() { return this.getValues()[0][0]; }
  setValues(v) {
    for (let i = 0; i < v.length; i++) {
      const ri = this.r - 1 + i;
      while (this.sh.rows.length <= ri) this.sh.rows.push([]);
      const row = this.sh.rows[ri];
      for (let j = 0; j < v[i].length; j++) row[this.c - 1 + j] = v[i][j];
    }
    this.sh.dirty = true;
    return this;
  }
  setValue(x) { return this.setValues([[x]]); }
  clearContent() { return this.setValues(Array.from({ length: this.nr }, () => Array(this.nc).fill(''))); }
  clear() { return this.clearContent(); }
  setNumberFormat() { return this; }
  setFontWeight() { return this; }
}
class Sheet {
  constructor(name, rows) { this.name = name; this.rows = rows || []; this.dirty = false; }
  getName() { return this.name; }
  getLastRow() { let n = this.rows.length; while (n > 0 && (this.rows[n - 1] || []).every(v => v === '' || v === null || v === undefined)) n--; return n; }
  getLastColumn() { let m = 0; for (const r of this.rows) for (let c = r.length; c > m; c--) if (r[c - 1] !== '' && r[c - 1] != null) { m = c; break; } return m; }
  getMaxColumns() { return Math.max(26, ...this.rows.map(r => r.length)); }
  getMaxRows() { return Math.max(1000, this.rows.length); }
  getRange(r, c, nr, nc) { return new Range(this, r, c, nr, nc); }
  getDataRange() { return this.getRange(1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
  appendRow(v) { this.rows.length = this.getLastRow(); this.rows.push(v.slice()); this.dirty = true; return this; }
  setFrozenRows() { return this; }
}
const sheets = {};
let googleInn = null;   // ИНН, которые были в листе «Рейтинги» Google на начало запуска
function googleSheet(name) {
  let err;
  for (let i = 0; i < 3; i++) {
    try {
      const r = curlMany([{ url: EXEC + '?d=sheet&name=' + encodeURIComponent(name) }])[0];
      if (r.code !== 200) throw new Error('код ' + r.code);
      const j = JSON.parse(r.getContentText());
      if (!j.rows) throw new Error(j.error || 'нет данных');
      return j.rows.map(row => row.map(revive));
    } catch (e) { err = e; Utilities.sleep(5000); }
  }
  throw err;
}
const stateFile = name => path.join(STATE, 'sheets', name.replace(/[\\/:*?"<>|]/g, '_') + '.json');
function loadSheet(name) {
  const f = stateFile(name);
  const fromState = () => (fs.existsSync(f) ? JSON.parse(fs.readFileSync(f, 'utf8')).map(r => r.map(revive)) : null);
  let rows = null, src = '';
  if (INPUT.includes(name)) {
    try { rows = googleSheet(name); src = 'Google'; if (name === 'Рейтинги') googleInn = new Set(rows.slice(1).map(r => String(r[0]).trim())); }
    catch (e) { rows = fromState(); src = 'state (Google не ответил: ' + e.message + ')'; }
  } else {
    rows = fromState(); src = 'state';
    if (!rows) { try { rows = googleSheet(name); src = 'Google (первый запуск)'; } catch (e) { rows = null; } }
  }
  console.log('  лист «' + name + '»: ' + (rows ? rows.length + ' строк, ' + src : 'нет'));
  return rows;
}
const SpreadsheetApp = {
  getActive() { return SS; }, getActiveSpreadsheet() { return SS; },
};
const SS = {
  getSheetByName(name) {
    if (!(name in sheets)) { const rows = loadSheet(name); sheets[name] = rows ? new Sheet(name, rows) : null; }
    return sheets[name];
  },
  insertSheet(name) { sheets[name] = new Sheet(name, []); sheets[name].dirty = true; return sheets[name]; },
};

// ---------- свойства, кэш, триггеры, прочее ----------
const propsFile = path.join(STATE, 'props.json');
let props = fs.existsSync(propsFile) ? JSON.parse(fs.readFileSync(propsFile, 'utf8')) : null;
let gProps = {};
try { gProps = JSON.parse(curlMany([{ url: EXEC + '?d=sheet&name=' + encodeURIComponent('#props') }])[0].getContentText()).props || {}; } catch (e) { console.log('свойства из Google не получены:', e.message); }
if (!props) props = Object.assign({}, gProps);
PROPS_FROM_GOOGLE.forEach(k => { if (gProps[k]) props[k] = gProps[k]; });
const PropertiesService = {
  getScriptProperties() {
    return {
      getProperty: k => (k === 'TINVEST_TOKEN' ? process.env.TINVEST_TOKEN || null : props[k] === undefined ? null : props[k]),
      setProperty(k, v) { if (k !== 'TINVEST_TOKEN') props[k] = String(v); return this; },
      deleteProperty(k) { delete props[k]; return this; },
      getProperties: () => Object.assign({}, props),
    };
  },
};
const cache = { get: () => null, getAll: () => ({}), put() {}, putAll() {}, remove() {}, removeAll() {} };
const CacheService = { getScriptCache: () => cache };
const pending = [];
function builder(fn) {
  const b = { timeBased: () => b, everyDays: () => b, atHour: () => b, nearMinute: () => b, inTimezone: () => b, everyHours: () => b,
    after() { b.once = true; return b; }, create() { if (b.once && !pending.includes(fn)) pending.push(fn); return { getHandlerFunction: () => fn }; } };
  return b;
}
const ScriptApp = { getProjectTriggers: () => [], deleteTrigger() {}, newTrigger: builder };
const pad = n => String(n).padStart(2, '0');
const Utilities = {
  formatDate(d, tz, fmt) {   // в коде только Europe/Moscow (UTC+3, без перехода на летнее время)
    const m = new Date(d.getTime() + 3 * 3600e3);
    return fmt.replace(/yyyy|MM|dd|HH|mm|ss/g, t => ({ yyyy: m.getUTCFullYear(), MM: pad(m.getUTCMonth() + 1), dd: pad(m.getUTCDate()), HH: pad(m.getUTCHours()), mm: pad(m.getUTCMinutes()), ss: pad(m.getUTCSeconds()) })[t]);
  },
  sleep(ms) { Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms); },
  DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' },
  computeDigest(alg, text) { return Array.from(crypto.createHash(alg).update(text, 'utf8').digest()).map(b => (b > 127 ? b - 256 : b)); },
};
const ContentService = { MimeType: { JSON: 'json' }, createTextOutput(t) { const o = { t, setMimeType: () => o, getContent: () => t }; return o; } };

// ---------- код Google ----------
const ctx = vm.createContext({ UrlFetchApp, SpreadsheetApp, PropertiesService, CacheService, ScriptApp, Utilities, ContentService, console, Date, Math, JSON });
const src = ['Code.gs', 'Bonds.gs', 'Market.gs', 'Dohod.gs'].map(f => fs.readFileSync(path.join(GAS_DIR, f), 'utf8')).join('\n');
vm.runInContext(src, ctx, { filename: 'gas.js' });
vm.runInContext('TIME_LIMIT_MS = 40 * 60 * 1000;', ctx);   // у Google 6 минут на запуск, здесь ограничения нет
const saved = {};
// готовые наборы страниц — в data/, служебные (например, dohod-base) — в state/json/
const PAGE_KEYS = ['bonds', 'coupons', 'dividends', 'curve', 'indices', 'shares', 'funds', 'fx', 'futures', 'options', 'dohod'];
const jsonFile = key => (PAGE_KEYS.includes(key) ? path.join(DATA, key + '.json') : path.join(STATE, 'json', key + '.json'));
ctx.saveJson_ = function (key, obj) {
  const s = JSON.stringify(obj), f = jsonFile(key);
  fs.mkdirSync(path.dirname(f), { recursive: true });
  fs.writeFileSync(f, s);
  saved[key] = s.length;
  return s.length;
};
ctx.loadJson_ = function (key) { const f = jsonFile(key); return fs.existsSync(f) ? fs.readFileSync(f, 'utf8') : ''; };
ctx.marketJson_ = function (key) { return ctx.loadJson_(key) || '{"updated":"","error":"данные ещё не собраны"}'; };
const journal = [];
ctx.log_ = function (msg) { const line = Utilities.formatDate(new Date(), '', 'dd.MM.yyyy HH:mm') + ' ' + msg; journal.push(line); console.log('  журнал:', msg); };

const JOBS = {
  bonds: { fn: 'refreshBonds', out: () => { if (props.BONDS_UPDATED === ctx.ymd_(new Date())) write('bonds', ctx.bondsJson_()); else console.log('  облигации: дата обновления не сегодняшняя, файл не пишем'); } },
  coupons: { fn: 'refreshCoupons' },
  dividends: { fn: 'refresh', out: () => { if (props.UPDATED === ctx.ymd_(new Date())) write('dividends', ctx.json_()); else console.log('  дивиденды: дата обновления не сегодняшняя, файл не пишем'); } },
  curve: { fn: 'refreshCurve' }, indices: { fn: 'refreshIndices' }, shares: { fn: 'refreshShares' }, funds: { fn: 'refreshFunds' },
  fx: { fn: 'refreshFx' }, futures: { fn: 'refreshFutures' }, options: { fn: 'refreshOptions' }, dohod: { fn: 'refreshDohod' },
};
function write(key, s) { fs.mkdirSync(DATA, { recursive: true }); fs.writeFileSync(path.join(DATA, key + '.json'), s); saved[key] = s.length; }

const want = process.argv.slice(2).length ? process.argv.slice(2) : Object.keys(JOBS);
const report = [];
for (const k of want) {
  const job = JOBS[k];
  if (!job) { console.log('нет такого набора:', k); continue; }
  const t0 = Date.now();
  console.log('== ' + k + ' (' + job.fn + ')');
  let err = '';
  try {
    ctx[job.fn]();
    for (let i = 0; pending.length && i < 30; i++) { const fn = pending.shift(); console.log('  догрузка: ' + fn); ctx[fn](); }
    pending.length = 0;
    if (job.out) job.out();
  } catch (e) { err = String(e && e.stack || e); console.log('  ОШИБКА: ' + err); }
  const sec = Math.round((Date.now() - t0) / 1000);
  report.push(k + ': ' + (err ? 'ошибка' : saved[k] ? 'готово, ' + Math.round(saved[k] / 1024) + ' КБ' : 'файл не обновлён') + ', ' + sec + ' с');
}

// ---------- новые эмитенты → лист «Рейтинги» в Google ----------
// Раньше ratings_() дописывала их прямо в таблицу; теперь отправляем в веб-приложение (doPost, kind ratings_add),
// оно дописывает только тех, кого в листе ещё нет. Подпись — SHA-256 токена Т-Инвестиций, как у сборщика дивидендов.
(function sendNewIssuers() {
  const sh = sheets['Рейтинги'];
  if (!sh || !googleInn) return;
  const fresh = sh.rows.slice(1).filter(r => r[0] !== '' && r[0] != null && !googleInn.has(String(r[0]).trim())).map(r => [String(r[0]).trim(), String(r[1] || '')]);
  if (!fresh.length) { console.log('Рейтинги: новых эмитентов нет'); return; }
  if (!process.env.TINVEST_TOKEN) { console.log('Рейтинги: новых эмитентов ' + fresh.length + ', но нет TINVEST_TOKEN — не отправлены'); return; }
  const body = JSON.stringify({ kind: 'ratings_add', sig: crypto.createHash('sha256').update(process.env.TINVEST_TOKEN, 'utf8').digest('hex'), rows: fresh });
  const f = path.join(os.tmpdir(), 'ratings_add.json');
  fs.writeFileSync(f, body);
  let out = '';
  try { out = cp.execFileSync('curl', ['-sL', '--max-time', '120', '-H', 'Content-Type: application/json', '--data-binary', '@' + f, EXEC], { encoding: 'utf8' }); }
  catch (e) { out = 'ошибка curl: ' + e.message; }
  fs.rmSync(f, { force: true });
  console.log('Рейтинги: отправлено новых эмитентов ' + fresh.length + ' (' + fresh.map(r => r[1] || r[0]).slice(0, 5).join('; ') + (fresh.length > 5 ? '…' : '') + '), ответ: ' + out.slice(0, 200));
})();

// ---------- сохранить состояние ----------
fs.mkdirSync(path.join(STATE, 'sheets'), { recursive: true });
for (const name of Object.keys(sheets)) {
  const sh = sheets[name];
  if (!sh || !(KEEP.includes(name) || INPUT.includes(name))) continue;
  fs.writeFileSync(stateFile(name), JSON.stringify(sh.rows.slice(0, sh.getLastRow()).map(r => r.map(pack))));
}
fs.writeFileSync(propsFile, JSON.stringify(props, null, 1));
const jf = path.join(STATE, 'journal.txt');
const old = fs.existsSync(jf) ? fs.readFileSync(jf, 'utf8').split('\n').filter(Boolean) : [];
fs.writeFileSync(jf, old.concat(journal).slice(-2000).join('\n') + '\n');
console.log('\nИтог:\n' + report.join('\n'));
if (process.env.GITHUB_STEP_SUMMARY) fs.appendFileSync(process.env.GITHUB_STEP_SUMMARY, '### Сбор на GitHub\n' + report.map(r => '- ' + r).join('\n') + '\n');
