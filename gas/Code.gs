/**
 * Календарь дивидендов: данные для страницы bulkacapital.ru/kalendar-dividendov
 *
 * Раз в сутки забирает список акций, цены и секторы с Мосбиржи (ISS). Дивиденды раз в неделю присылает
 * сборщик с компьютера (Т-Инвестиции не пускают серверы Google), всё складывается в таблицу
 * и отдаёт странице готовый JSON (веб-приложение, doGet).
 *
 * Первый запуск: функция setup() (создаёт листы, ставит ежедневный триггер, запускает обновление).
 * Цены на даты прошлых выплат подгружаются порциями: если за один запуск не успели,
 * следующий запуск продолжит с того же места.
 */
var ISS = 'https://iss.moex.com/iss';
var YEARS_BACK = 13;          // сколько лет истории показываем
var T1_FROM = '2023-07-31';   // с этой даты акции торгуются в режиме Т+1, раньше Т+2
var TIME_LIMIT_MS = 4.5 * 60 * 1000;
var BATCH = 40;

var SH_CO = 'Компании', SH_PAY = 'Выплаты', SH_MAN = 'Секторы вручную', SH_LOG = 'Журнал';
var CO_HEAD = ['Тикер', 'Название', 'Сектор', 'Цена, ₽', 'Цена на начало года, ₽', 'Год этой цены'];
var PAY_HEAD = ['Тикер', 'Закрытие реестра', 'Дивиденд, ₽', 'Последний день покупки', 'Цена в этот день, ₽'];
var MAN_HEAD = ['Тикер', 'Сектор', 'Название'];
var SH_RAW = 'Дивиденды (сырые)', SH_RAWCO = 'Дивиденды: компании';
var RAW_HEAD = ['Тикер', 'Закрытие реестра', 'Дивиденд, ₽', 'Последний день покупки', 'Объявлено', 'Тип'];
var RAWCO_HEAD = ['Тикер', 'Название', 'Сектор (Т-Инвестиции)'];

// Отраслевые индексы Мосбиржи: по ним определяем сектор компании.
var SECTOR_INDEXES = [
  ['MOEXOG', 'Нефть и газ'], ['MOEXFN', 'Финансы'], ['MOEXMM', 'Металлы и добыча'],
  ['MOEXEU', 'Электроэнергетика'], ['MOEXTL', 'Телекоммуникации'], ['MOEXCN', 'Потребительский сектор'],
  ['MOEXCH', 'Химия и удобрения'], ['MOEXIT', 'Информационные технологии'],
  ['MOEXRE', 'Строительство и недвижимость'], ['MOEXTN', 'Транспорт']
];
var OTHER_SECTOR = 'Другое';
// Сектор из Т-Инвестиций: берём его, когда он однозначный. Для промышленности и сырья точнее индекс Мосбиржи.
var T_SECTOR = {
  consumer: 'Потребительский сектор', energy: 'Нефть и газ', financial: 'Финансы', health_care: 'Здравоохранение',
  it: 'Информационные технологии', real_estate: 'Строительство и недвижимость', telecom: 'Телекоммуникации',
  utilities: 'Электроэнергетика'
};
var T_SECTOR_WEAK = { industrials: 'Промышленность', materials: 'Металлы и добыча' };
var FIX_SECTOR = {
  KAZTP: 'Химия и удобрения', VLHZ: 'Химия и удобрения', SGZH: 'Промышленность', EELT: 'Промышленность',
  WTCM: 'Строительство и недвижимость', WTCMP: 'Строительство и недвижимость'
};
var FIX_NAME = {
  OGKB: 'ОГК-2', X5: 'ИКС 5', PIKK: 'ПИК', MAGN: 'ММК', TRMK: 'ТМК', BELU: 'НоваБев Групп', UNAC: 'ОАК',
  RKKE: 'РКК Энергия', CHKZ: 'ЧКПЗ', OMZZP: 'ОМЗ', AQUA: 'Инарктика'
};
function sector_(t, tsec, moex) {
  return FIX_SECTOR[t] || T_SECTOR[tsec] || moex || T_SECTOR_WEAK[tsec] || OTHER_SECTOR;
}
/** Короткое название; привилегированные акции помечаем «(прив.)». */
function name_(t, name, listed) {
  var n = String(name || '');
  var pref = /привилегированн|-ап$/i.test(n) || (t.length === 5 && /P$/.test(t) && !!listed[t.slice(0, 4)]);
  n = FIX_NAME[t] || n.replace(/\s*[-(]?\s*(акции\s+)?привилегированные\)?(\s+акции)?\s*$/i, '').replace(/-ап$/i, '').trim();
  return pref ? n + ' (прив.)' : n;
}

// ---------- запуск ----------
function setup() {
  var ss = SpreadsheetApp.getActive();
  sheet_(SH_CO, CO_HEAD); sheet_(SH_PAY, PAY_HEAD); sheet_(SH_MAN, MAN_HEAD); sheet_(SH_LOG, ['Время', 'Событие']);
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'refresh') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('refresh').timeBased().everyDays(1).atHour(20).inTimezone('Europe/Moscow').create();
  refresh();
}

function refresh() {
  var started = Date.now();
  var today = ymd_(new Date());
  var minDate = (Number(today.slice(0, 4)) - YEARS_BACK + 1) + '-01-01';
  var baseYear = Number(today.slice(0, 4)) - 2;

  // 1. Список акций и последние цены
  var j = fetchJson_(ISS + '/engines/stock/markets/shares/boards/TQBR/securities.json?iss.meta=off' +
    '&iss.only=securities,marketdata&securities.columns=SECID,SHORTNAME,PREVPRICE,PREVLEGALCLOSEPRICE' +
    '&marketdata.columns=SECID,LAST');
  if (!j) { log_('Мосбиржа не ответила, обновление пропущено'); return; }
  var last = {};
  rows_(j.marketdata).forEach(function (r) { last[r.SECID] = r.LAST; });
  var secs = rows_(j.securities).map(function (r) {
    return { t: r.SECID, name: r.SHORTNAME, price: pos_(last[r.SECID], r.PREVLEGALCLOSEPRICE, r.PREVPRICE) };
  });

  // 2. Что уже лежит в таблице
  var oldCo = {}, oldPay = {}, manual = {};
  read_(SH_CO).forEach(function (r) { oldCo[r[0]] = r; });
  read_(SH_PAY).forEach(function (r) { oldPay[r[0] + '|' + day_(r[1])] = r; });
  read_(SH_MAN).forEach(function (r) { if (r[0]) manual[String(r[0]).trim().toUpperCase()] = r; });

  // 3. Секторы по отраслевым индексам
  var sector = {};
  var idx = fetchAll_(SECTOR_INDEXES.map(function (s) {
    return ISS + '/statistics/engines/stock/markets/index/analytics/' + s[0] + '/tickers.json?iss.meta=off';
  }));
  idx.forEach(function (res, i) {
    rows_(res && res.tickers).forEach(function (r) { var t = r.ticker || r.TICKER; if (t) sector[t] = SECTOR_INDEXES[i][1]; });
  });

  // 4. Дивиденды. Т-Инвестиции не пускают серверы Google, поэтому раз в неделю их присылает сборщик
  //    с компьютера Юлии (см. doPost), а здесь читаем готовое с листа «Дивиденды (сырые)».
  var raw = read_(SH_RAW);
  if (!raw.length) { log_('На листе «' + SH_RAW + '» нет дивидендов, обновление пропущено'); return; }
  var tName = {}, tSec = {};
  read_(SH_RAWCO).forEach(function (r) { tName[r[0]] = r[1]; tSec[r[0]] = r[2]; });
  var listed = {};
  secs.forEach(function (s) { listed[s.t] = 1; });
  var pays = [], seen = {}, failed = 0;   // pays: {t, reg, div, buy, price, tbuy}
  raw.forEach(function (r) {
    var t = String(r[0]), reg = day_(r[1]), val = Number(r[2]);
    if (!listed[t] || !(val > 0) || !reg || reg < minDate || seen[t + '|' + reg]) return;
    seen[t + '|' + reg] = 1;
    var o = oldPay[t + '|' + reg];
    pays.push({ t: t, reg: reg, div: val, buy: o ? day_(o[3]) : '', price: o ? o[4] : '', tbuy: day_(r[3]) });
  });

  // 5. Цена и последний день покупки для прошедших выплат (порциями, пока есть время)
  var need = pays.filter(function (p) { return p.reg < today && p.price === ''; });
  var done = 0;
  for (var a = 0; a < need.length && Date.now() - started < TIME_LIMIT_MS; a += BATCH) {
    var part = need.slice(a, a + BATCH);
    var hist = fetchAll_(part.map(function (p) {
      return ISS + '/history/engines/stock/markets/shares/securities/' + p.t + '.json?iss.meta=off&limit=100' +
        '&history.columns=BOARDID,TRADEDATE,CLOSE,LEGALCLOSEPRICE,VALUE&from=' + addDays_(p.reg, -14) + '&till=' + p.reg;
    }));
    part.forEach(function (p, i) {
      if (!hist[i]) return;                       // попробуем в следующий раз
      var r = buyDay_(daily_(hist[i]), p.reg);
      p.buy = r.buy || weekdayBuy_(p.reg);
      p.price = r.price || '-';                   // «-» значит: искали, цены нет
      done++;
    });
    Utilities.sleep(300);
  }
  pays.forEach(function (p) { if (p.reg >= today || !p.buy) p.buy = p.tbuy || weekdayBuy_(p.reg); });

  // 6. Цена на начало позапрошлого года (для калькулятора «за 2 года»)
  var payers = {};
  pays.forEach(function (p) { payers[p.t] = 1; });
  var companies = secs.filter(function (s) { return payers[s.t]; });
  var needP0 = companies.filter(function (s) { var o = oldCo[s.t]; return !(o && Number(o[5]) === baseYear && o[4] !== ''); });
  var p0 = {};
  for (var b = 0; b < needP0.length && Date.now() - started < TIME_LIMIT_MS; b += BATCH) {
    var part0 = needP0.slice(b, b + BATCH);
    var h0 = fetchAll_(part0.map(function (s) {
      return ISS + '/history/engines/stock/markets/shares/securities/' + s.t + '.json?iss.meta=off&limit=100' +
        '&history.columns=BOARDID,TRADEDATE,CLOSE,LEGALCLOSEPRICE,VALUE&from=' + baseYear + '-01-01&till=' + baseYear + '-01-20';
    }));
    part0.forEach(function (s, i) {
      if (!h0[i]) return;
      var d = daily_(h0[i]);
      p0[s.t] = d.length ? d[0].close : '-';
    });
    Utilities.sleep(300);
  }

  // 7. Запись в таблицу
  var coRows = companies.map(function (s) {
    var o = oldCo[s.t], m = manual[s.t];
    var start = p0[s.t] !== undefined ? p0[s.t] : (o && Number(o[5]) === baseYear ? o[4] : '');
    return [s.t, (m && m[2]) || name_(s.t, tName[s.t] || s.name, listed), (m && m[1]) || sector_(s.t, tSec[s.t], sector[s.t]), s.price || (o ? o[3] : ''), start, start === '' ? '' : baseYear];
  });
  pays.sort(function (x, y) { return x.t < y.t ? -1 : x.t > y.t ? 1 : (x.reg < y.reg ? -1 : 1); });
  write_(SH_CO, CO_HEAD, coRows);
  write_(SH_PAY, PAY_HEAD, pays.map(function (p) { return [p.t, p.reg, p.div, p.buy, p.price]; }));
  PropertiesService.getScriptProperties().setProperty('UPDATED', today);
  dropCache_();
  var left = need.length - done;
  log_('Обновлено: компаний ' + coRows.length + ', выплат ' + pays.length + ', цен найдено ' + done +
    (left > 0 ? ', осталось найти ' + left : '') + (failed ? ', без ответа биржи ' + failed : ''));
  if (left > 0) again_();
}

/** Если за один запуск не успели найти все цены, продолжаем через минуту. */
function again_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'refreshMore') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('refreshMore').timeBased().after(60 * 1000).create();
}
function refreshMore() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'refreshMore') ScriptApp.deleteTrigger(t);
  });
  refresh();
}

// ---------- веб-приложение ----------
/**
 * Приём дивидендов от сборщика на компьютере (dividendy/local/collect_dividends.py).
 * Подпись запроса: SHA-256 от токена Т-Инвестиций, который лежит в свойствах скрипта. Сам токен не передаётся.
 */
function doPost(e) {
  var out;
  try {
    var body = JSON.parse(e.postData.contents);
    var token = PropertiesService.getScriptProperties().getProperty('TINVEST_TOKEN') || '';
    if (!token || body.sig !== sha256_(token)) throw new Error('подпись не совпала');
    if (body.kind === 'fund') return postFund_(body);     // финансовые показатели акций для страницы «Анализ акций»
    if (!body.items || !body.items.length) throw new Error('пустой список выплат');
    write_(SH_RAW, RAW_HEAD, body.items.map(function (r) { return [r[0], r[1], r[2], r[3], r[4], r[5]]; }));
    write_(SH_RAWCO, RAWCO_HEAD, (body.companies || []).map(function (r) { return [r[0], r[1], r[2]]; }));
    log_('Дивиденды получены с компьютера: выплат ' + body.items.length + ', компаний ' + (body.companies || []).length);
    again_();   // календарь пересчитается через минуту, чтобы не держать соединение
    out = { ok: true, items: body.items.length };
  } catch (x) { out = { ok: false, error: String(x.message || x) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function postFund_(body) {
  if (!body.items || body.items.length < 50) throw new Error('слишком мало бумаг с показателями');
  write_(SH_FUND, ['Тикер', 'Показатели (JSON)'], body.items.map(function (r) { return [r.t, '#' + JSON.stringify(r)]; }));
  PropertiesService.getScriptProperties().setProperty('FUND_UPDATED', String(body.updated || ''));
  log_('Показатели акций получены с компьютера: бумаг ' + body.items.length);
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'refreshSharesMore') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('refreshSharesMore').timeBased().after(60 * 1000).create();
  return ContentService.createTextOutput(JSON.stringify({ ok: true, items: body.items.length })).setMimeType(ContentService.MimeType.JSON);
}
function refreshSharesMore() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'refreshSharesMore') ScriptApp.deleteTrigger(t); });
  refreshShares();
}
function sha256_(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}
function doGet(e) {
  var d = e && e.parameter && e.parameter.d;   // .../exec?d=bonds отдаёт облигации, без параметра календарь дивидендов
  var market = { shares: 1, funds: 1, indices: 1, curve: 1, coupons: 1, fx: 1, futures: 1, options: 1, dohod: 1 };
  var text = d === 'sheet' ? sheetDump_(e.parameter.name) : d === 'bonds' ? bondsJson_() : market[d] ? marketJson_(d) : json_();   // 24.11.2026: d=sheet — листы для сбора на GitHub
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

function json_() {
  var cache = CacheService.getScriptCache();
  var n = Number(cache.get('n') || 0), parts = [];
  if (n) {
    var keys = []; for (var i = 0; i < n; i++) keys.push('p' + i);
    var got = cache.getAll(keys);
    for (var k = 0; k < n; k++) parts.push(got['p' + k]);
    if (parts.every(function (x) { return typeof x === 'string'; })) return parts.join('');
  }
  var today = ymd_(new Date());
  var byT = {};
  read_(SH_PAY).forEach(function (r) {
    var price = Number(r[4]);
    var k = splitK_(r[0], day_(r[1]));     // выплаты и цены до дробления приводим к нынешней акции
    (byT[r[0]] = byT[r[0]] || []).push([day_(r[1]), k === 1 ? Number(r[2]) : Math.round(Number(r[2]) * k * 1e6) / 1e6,
      price > 0 ? Math.round(price * k * 100) / 100 : null, day_(r[3])]);
  });
  var out = {
    updated: PropertiesService.getScriptProperties().getProperty('UPDATED') || today,
    companies: read_(SH_CO).filter(function (r) { return byT[r[0]]; }).map(function (r) {
      var k0 = splitK_(r[0], r[5] + '-01-20');
      return { t: r[0], name: r[1], sector: r[2], price: Number(r[3]) || null, p0: Number(r[4]) > 0 ? Math.round(Number(r[4]) * k0 * 100) / 100 : null, pays: byT[r[0]] };
    })
  };
  var s = JSON.stringify(out), put = {}, size = 90000, cnt = Math.ceil(s.length / size);
  for (var c = 0; c < cnt; c++) put['p' + c] = s.substr(c * size, size);
  put.n = String(cnt);
  try { cache.putAll(put, 3600); } catch (e) {}
  return s;
}
function dropCache_() { CacheService.getScriptCache().remove('n'); }
/** Во сколько раз уменьшить сумму на акцию для даты до дробления (KNOWN_SPLITS задан в части про акции). */
function splitK_(t, date) {
  var k = 1;
  (KNOWN_SPLITS[t] || []).forEach(function (s) { if (date < s[0]) k *= s[1]; });
  return k;
}

// ---------- расчёт последнего дня покупки ----------
/** Из ответа истории торгов: по одной строке на день, основной режим торгов. */
function daily_(res) {
  var rank = { TQBR: 3, EQBR: 2, EQNE: 2, EQBS: 1, EQNL: 1 };
  var best = {};
  rows_(res && res.history).forEach(function (r) {
    var close = pos_(r.LEGALCLOSEPRICE, r.CLOSE);
    if (!close) return;
    var score = (rank[r.BOARDID] || 0) * 1e15 + (Number(r.VALUE) || 0);
    if (!best[r.TRADEDATE] || score > best[r.TRADEDATE].score) best[r.TRADEDATE] = { date: r.TRADEDATE, close: close, score: score };
  });
  return Object.keys(best).sort().map(function (d) { return best[d]; });
}
/** Последний день покупки: за 1 торговый день до дня расчётов (за 2 дня до 31.07.2023). */
function buyDay_(days, reg) {
  var lag = reg < T1_FROM ? 2 : 1;
  var i = days.length - 1;
  while (i >= 0 && days[i].date > reg) i--;
  var b = days[i - lag];
  return b ? { buy: b.date, price: b.close } : {};
}
/** Запасной расчёт по календарю без праздников: для будущих выплат. */
function weekdayBuy_(reg) {
  var d = parse_(reg);
  function back() { d.setUTCDate(d.getUTCDate() - 1); }
  function weekend() { var w = d.getUTCDay(); return w === 0 || w === 6; }
  while (weekend()) back();
  var lag = reg < T1_FROM ? 2 : 1;
  for (var i = 0; i < lag; i++) { back(); while (weekend()) back(); }
  return d.toISOString().slice(0, 10);
}

// ---------- вспомогательное ----------
function fetchJson_(url) { return fetchAll_([url])[0]; }
function fetchAll_(urls) {
  var out = [];
  for (var i = 0; i < urls.length; i += BATCH) {
    var res;
    try {
      res = UrlFetchApp.fetchAll(urls.slice(i, i + BATCH).map(function (u) { return { url: u, muteHttpExceptions: true }; }));
    } catch (e) { res = urls.slice(i, i + BATCH).map(function () { return null; }); }
    res.forEach(function (r) {
      try { out.push(r && r.getResponseCode() === 200 ? JSON.parse(r.getContentText()) : null); } catch (x) { out.push(null); }
    });
  }
  return out;
}
function rows_(block) {
  if (!block || !block.columns) return [];
  return block.data.map(function (r) { var o = {}; block.columns.forEach(function (c, i) { o[c] = r[i]; }); return o; });
}
function pos_() {
  for (var i = 0; i < arguments.length; i++) { var v = arguments[i]; if (typeof v === 'number' && v > 0) return v; }
  return null;
}
function ymd_(d) { return Utilities.formatDate(d, 'Europe/Moscow', 'yyyy-MM-dd'); }
function day_(v) { return v instanceof Date ? ymd_(v) : String(v || '').slice(0, 10); }
function parse_(s) { return new Date(s + 'T00:00:00Z'); }
function addDays_(s, n) { var d = parse_(s); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }
function sheet_(name, head) {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) { sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
function read_(name) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues().filter(function (r) { return r[0] !== ''; });
}
function write_(name, head, rows) {
  var sh = sheet_(name, head);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getMaxColumns()).clearContent();
  if (!rows.length) return;
  sh.getRange(2, 1, rows.length, head.length).setNumberFormat('@').setValues(rows.map(function (r) {
    return r.map(function (v) { return v === null || v === undefined ? '' : String(v); });
  }));
}
function log_(msg) {
  var sh = sheet_(SH_LOG, ['Время', 'Событие']);
  sh.appendRow([Utilities.formatDate(new Date(), 'Europe/Moscow', 'dd.MM.yyyy HH:mm'), msg]);
  console.log(msg);
}

// 24.11.2026: Мосбиржа перешла на сертификат Минцифры, Google к ней больше не подключается. Сбор перенесён на GitHub (muzova2/site-data),
// он читает здесь листы с ручными и входящими данными. Только чтение и только листы из списка; свойства — только служебные даты.
var DUMP_SHEETS = ['Компании', 'Выплаты', 'Секторы вручную', 'Дивиденды (сырые)', 'Дивиденды: компании', 'Показатели акций', 'Облигации',
  'Облигации: справочник', 'Рейтинги', 'Кривая: история', 'Фонды: справочник', 'Фонды: СЧА с сайтов УК', 'Доходность: недвижимость'];
var DUMP_PROPS = ['UPDATED', 'FUND_UPDATED', 'BONDS_UPDATED', 'LIQ_LAST'];
function sheetDump_(name) {
  if (name === '#props') {
    var p = PropertiesService.getScriptProperties(), o = {};
    DUMP_PROPS.forEach(function (k) { o[k] = p.getProperty(k); });
    return JSON.stringify({ props: o });
  }
  if (DUMP_SHEETS.indexOf(name) < 0) return JSON.stringify({ error: 'лист не в списке' });
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) return JSON.stringify({ name: name, rows: [] });
  var v = sh.getDataRange().getValues().map(function (r) { return r.map(function (x) { return x instanceof Date ? { $d: x.toISOString() } : x; }); });
  return JSON.stringify({ name: name, rows: v });
}
