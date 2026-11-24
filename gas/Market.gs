/**
 * Данные для страниц «Анализ акций», «Скринер фондов», «Индексы», «Кривая ОФЗ и ставки», «Календарь купонов».
 *
 * Каждая страница получает готовый JSON: .../exec?d=shares | funds | indices | curve | coupons.
 * Готовый текст лежит на листе «Данные для страниц» кусками (в ячейку помещается до 50 тысяч знаков),
 * править его руками не нужно. Обновление раз в сутки, своим триггером на каждую страницу (setupMarket).
 * Историю цен скрипт не хранит: каждый день берёт у биржи дневные свечи за последние 14 месяцев
 * и недельные за 5 лет, этого хватает для всех показателей.
 */
var SH_OUT = 'Данные для страниц', SH_FUND = 'Показатели акций', SH_FREF = 'Фонды: справочник', SH_CURVE = 'Кривая: история';
var OUT_CHUNK = 45000;
var KNOWN_SPLITS = { BELU: [['2024-08-22', 1 / 8]] };   // НоваБев: по 7 новых акций на каждую, биржа старые цены не поправила
var NICE = [3, 4, 5, 6, 7, 8, 10, 20, 25, 40, 50, 100, 200, 500, 1000, 5000];
var T_SECTOR_ALL = {
  consumer: 'Потребительский сектор', energy: 'Нефть и газ', financial: 'Финансы', health_care: 'Здравоохранение',
  it: 'Информационные технологии', real_estate: 'Строительство и недвижимость', telecom: 'Телекоммуникации',
  utilities: 'Электроэнергетика', industrials: 'Промышленность', materials: 'Металлы и добыча'
};

// ---------- запуск ----------
function setupMarket() {
  var plan = [['refreshCoupons', 19, 40], ['refreshCurve', 20, 20], ['refreshIndices', 20, 40], ['refreshShares', 21, 0], ['refreshFunds', 21, 20], ['refreshFx', 21, 40], ['refreshFutures', 22, 0], ['refreshOptions', 22, 20]];
  var names = plan.map(function (p) { return p[0]; });
  ScriptApp.getProjectTriggers().forEach(function (t) { if (names.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); });
  plan.forEach(function (p) {
    ScriptApp.newTrigger(p[0]).timeBased().everyDays(1).atHour(p[1]).nearMinute(p[2]).inTimezone('Europe/Moscow').create();
  });
  log_('Триггеры страниц поставлены: ' + names.join(', '));
}
function refreshMarketAll() { refreshCoupons(); refreshIndices(); refreshCurve(); }

// ---------- хранение готового JSON ----------
function saveJson_(key, obj) {
  var s = JSON.stringify(obj), sh = sheet_(SH_OUT, ['Ключ', 'Часть', 'Текст']);
  var all = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues() : [];
  var keep = all.filter(function (r) { return r[0] !== '' && r[0] !== key; });
  for (var i = 0, n = 0; i < s.length; i += OUT_CHUNK, n++) keep.push([key, n, '#' + s.substr(i, OUT_CHUNK)]);   // «#» защищает кусок от разбора как формулы
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).clearContent();
  sh.getRange(2, 1, keep.length, 3).setNumberFormat('@').setValues(keep.map(function (r) { return [String(r[0]), String(r[1]), String(r[2])]; }));
  return s.length;
}
function loadJson_(key) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SH_OUT);
  if (!sh || sh.getLastRow() < 2) return '';
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
    .filter(function (r) { return r[0] === key; })
    .sort(function (a, b) { return Number(a[1]) - Number(b[1]); })
    .map(function (r) { return String(r[2]).slice(1); }).join('');
}
function marketJson_(key) { return loadJson_(key) || '{"updated":"","error":"данные ещё не собраны"}'; }

// ---------- общее: свечи и показатели ----------
/** Свечи биржи по списку бумаг: {код: [[дата, закрытие, оборот], ...]}. */
function candles_(market, ids, interval, from, board) {
  var urls = ids.map(function (id) {
    return ISS + '/engines/' + (market.indexOf('/') < 0 ? 'stock/markets/' + market : market) + '/' + (board ? 'boards/' + board + '/' : '') + 'securities/' + encodeURIComponent(id) +
      '/candles.json?iss.meta=off&interval=' + interval + '&from=' + from + '&candles.columns=close,value,begin';
  });
  var res = fetchAll_(urls), out = {};
  ids.forEach(function (id, i) {
    var d = res[i] && res[i].candles ? res[i].candles.data : null;
    out[id] = d ? d.filter(function (r) { return r[0]; }).map(function (r) { return [String(r[2]).slice(0, 10), r[0], r[1] || 0]; }) : null;
  });
  return out;
}
/** Дробления, видимые в ценах как скачок в «круглое» число раз. exact: брать любой скачок как есть (для биржевых фондов). */
function splits_(series, exact) {
  var out = [];
  for (var i = 1; i < series.length; i++) {
    var r = series[i][1] / series[i - 1][1];
    if (r > 0.36 && r < 2.7) continue;
    var best = null, err = 9;
    NICE.forEach(function (n) { var e = Math.abs((r < 1 ? r * n : r / n) - 1); if (e < err) { err = e; best = n; } });
    if (err <= 0.15) out.push([series[i][0], r < 1 ? 1 / best : best]);
    else if (exact) out.push([series[i][0], r]);
  }
  return out;
}
function adjust_(series, splits) {
  if (!splits.length) return series;
  return series.map(function (p) {
    var m = 1;
    splits.forEach(function (s) { if (p[0] < s[0]) m *= s[1]; });
    return [p[0], p[1] * m, p[2]];
  });
}
function round_(v, d) { var k = Math.pow(10, d); return Math.round(v * k) / k; }
/** Изменения цены за периоды и технические показатели. daily: около 14 месяцев дневных цен, weekly: 5 лет недельных. */
function indicators_(daily, weekly, price, today) {
  var res = { w1: null, m1: null, m3: null, m6: null, ytd: null, y1: null, y3: null, y5: null, dhi: null, vol: null, ma50: null, ma200: null, rsi: null, hi: null, lo: null };
  if (!price) return res;
  function at(series, iso) {
    if (!series || !series.length || series[0][0] > iso) return null;
    var v = null;
    for (var i = 0; i < series.length && series[i][0] <= iso; i++) v = series[i][1];
    return v;
  }
  function chg(then) { return then ? round_((price / then - 1) * 100, 1) : null; }
  if (daily && daily.length > 1) {
    [['w1', 7], ['m1', 30], ['m3', 91], ['m6', 182], ['y1', 365]].forEach(function (p) { res[p[0]] = chg(at(daily, addDays_(today, -p[1]))); });
    res.ytd = chg(at(daily, (Number(today.slice(0, 4)) - 1) + '-12-31'));
    var yearAgo = addDays_(today, -365), year = daily.filter(function (p) { return p[0] >= yearAgo; }).map(function (p) { return p[1]; });
    var closes = daily.map(function (p) { return p[1]; });
    if (year.length) {
      res.hi = round_(Math.max.apply(null, year.concat([price])), 4);
      res.lo = round_(Math.min.apply(null, year.concat([price])), 4);
    }
    if (year.length >= 60) {
      res.dhi = round_((price / res.hi - 1) * 100, 1);
      var rets = [], sum = 0;
      for (var i = 1; i < year.length; i++) if (year[i] > 0 && year[i - 1] > 0) rets.push(Math.log(year[i] / year[i - 1]));
      rets.forEach(function (x) { sum += x; });
      var mean = sum / rets.length, sq = 0;
      rets.forEach(function (x) { sq += (x - mean) * (x - mean); });
      res.vol = round_(Math.sqrt(sq / rets.length) * Math.sqrt(252) * 100, 1);
    }
    [['ma50', 50], ['ma200', 200]].forEach(function (p) {
      if (closes.length < p[1]) return;
      var s = 0; closes.slice(-p[1]).forEach(function (x) { s += x; });
      res[p[0]] = round_((price / (s / p[1]) - 1) * 100, 1);
    });
    if (closes.length >= 60) {
      var tail = closes.slice(-250), gain = 0, loss = 0, d;
      for (var a = 1; a < 15; a++) { d = tail[a] - tail[a - 1]; gain += Math.max(d, 0) / 14; loss += Math.max(-d, 0) / 14; }
      for (var b = 15; b < tail.length; b++) { d = tail[b] - tail[b - 1]; gain = (gain * 13 + Math.max(d, 0)) / 14; loss = (loss * 13 + Math.max(-d, 0)) / 14; }
      res.rsi = loss ? Math.round(100 - 100 / (1 + gain / loss)) : 100;
    }
  }
  if (weekly && weekly.length > 1) {
    res.y3 = chg(at(weekly, addDays_(today, -365 * 3)));
    res.y5 = chg(at(weekly, addDays_(today, -365 * 5)));
    if (res.y1 === null) res.y1 = chg(at(weekly, addDays_(today, -365)));
  }
  return res;
}
function nz_(v, digits) {      // ноль в данных брокера означает «нет данных»
  if (v === null || v === undefined || v === '' || Number(v) === 0 || isNaN(Number(v))) return null;
  return round_(Number(v), digits === undefined ? 2 : digits);
}
function pick_(row, keys) { return keys.map(function (k) { return row[k] === undefined ? null : row[k]; }); }
function pages_(path, block) {       // все страницы постраничного ответа биржи
  var out = [], start = 0;
  for (;;) {
    var j = fetchJson_(ISS + path + (path.indexOf('?') < 0 ? '?' : '&') + 'iss.meta=off&limit=100&start=' + start);
    var part = rows_(j && j[block]);
    out = out.concat(part);
    start += 100;
    if (part.length < 100 || start > 20000) return out;
  }
}

// ---------- акции ----------
var SHARE_F = ['t', 'name', 'sector', 'isin', 'list', 'lot', 'price', 'd1', 'w1', 'm1', 'm3', 'm6', 'ytd', 'y1', 'y3', 'y5', 'dhi',
  'cap', 'ev', 'turn', 'ff', 'beta', 'vol', 'rsi', 'ma50', 'ma200', 'hi', 'lo',
  'pe', 'pb', 'ps', 'evebitda', 'pfcf', 'roe', 'roa', 'margin', 'debt', 'de', 'nde', 'rev', 'ebitda', 'ni', 'eps', 'fcf',
  'dy', 'dy5', 'payout', 'dps', 'g1', 'g3', 'g5', 'div', 'sp'];
function refreshShares() {
  var today = ymd_(new Date());
  var fund = {};
  read_(SH_FUND).forEach(function (r) { try { fund[r[0]] = JSON.parse(String(r[1]).slice(1)); } catch (e) {} });
  if (!Object.keys(fund).length) { log_('Акции: на листе «' + SH_FUND + '» нет показателей, обновление пропущено'); return; }
  var j = fetchJson_(ISS + '/engines/stock/markets/shares/boards/TQBR/securities.json?iss.meta=off&iss.only=securities,marketdata' +
    '&securities.columns=SECID,SHORTNAME,ISIN,LOTSIZE,LISTLEVEL,PREVPRICE,PREVLEGALCLOSEPRICE&marketdata.columns=SECID,LAST,LCLOSEPRICE,ISSUECAPITALIZATION');
  if (!j) { log_('Акции: Мосбиржа не ответила, обновление пропущено'); return; }
  var md = {}, secs = {}, listed = {};
  rows_(j.marketdata).forEach(function (r) { md[r.SECID] = r; });
  rows_(j.securities).forEach(function (r) { listed[r.SECID] = 1; if (fund[r.SECID]) secs[r.SECID] = r; });
  var cal = {};
  read_(SH_CO).forEach(function (r) { cal[r[0]] = r; });
  var ids = Object.keys(secs).sort();
  var daily = candles_('shares', ids, 24, addDays_(today, -430)), weekly = candles_('shares', ids, 7, addDays_(today, -365 * 5 - 40));
  var rows = [], noHist = 0;
  ids.forEach(function (t) {
    var s = secs[t], m = md[t] || {}, f = fund[t], c = cal[t];
    var price = pos_(m.LCLOSEPRICE, m.LAST, s.PREVLEGALCLOSEPRICE, s.PREVPRICE);
    if (!price) return;
    var sp = KNOWN_SPLITS[t] || [];
    if (!daily[t]) noHist++;
    var ni = Number(f.netIncomeTtm) || 0, fcf = Number(f.freeCashFlowTtm) || 0, rev = Number(f.revenueTtm) || 0;
    var row = {
      t: t, name: c ? c[1] : name_(t, f.name || s.SHORTNAME, listed), sector: c ? c[2] : (T_SECTOR_ALL[f.sector] || OTHER_SECTOR),
      isin: s.ISIN, list: s.LISTLEVEL, lot: s.LOTSIZE, price: price,
      d1: pos_(m.LCLOSEPRICE, m.LAST) && s.PREVPRICE ? round_((price / s.PREVPRICE - 1) * 100, 1) : null,
      cap: nz_((pos_(m.ISSUECAPITALIZATION) || Number(f.marketCapitalization) || 0) / 1e9), ev: nz_((Number(f.totalEnterpriseValueMrq) || 0) / 1e9),
      turn: nz_((Number(f.averageDailyVolumeLast4Weeks) || 0) * price / 1e6), ff: nz_((Number(f.freeFloat) || 0) * 100, 0), beta: nz_(f.beta),
      pe: ni > 0 ? nz_(f.peRatioTtm) : null, pb: nz_(f.priceToBookTtm), ps: rev > 0 ? nz_(f.priceToSalesTtm) : null,
      evebitda: nz_(f.evToEbitdaMrq), pfcf: fcf > 0 ? nz_(f.priceToFreeCashFlowTtm) : null,
      roe: nz_(f.roe, 1), roa: nz_(f.roa, 1), margin: nz_(f.netMarginMrq, 1),
      debt: nz_((Number(f.totalDebtMrq) || 0) / 1e9), de: nz_(f.totalDebtToEquityMrq, 0), nde: nz_(f.netDebtToEbitda),
      rev: nz_(Math.max(rev, 0) / 1e9), ebitda: nz_((Number(f.ebitdaTtm) || 0) / 1e9), ni: nz_(ni / 1e9), eps: nz_(f.epsTtm), fcf: nz_(fcf / 1e9),
      dy: nz_(f.dividendYieldDailyTtm, 1), dy5: nz_(f.fiveYearsAverageDividendYield, 1), payout: nz_(f.dividendPayoutRatioFy, 0),
      dps: nz_(f.dividendRateTtm, 4), g1: nz_(f.oneYearAnnualRevenueGrowthRate, 1), g3: nz_(f.threeYearAnnualRevenueGrowthRate, 1),
      g5: nz_(f.revenueChangeFiveYears, 1), div: c ? 1 : 0, sp: sp.length ? sp : null
    };
    var ind = indicators_(adjust_(daily[t] || [], sp), adjust_(weekly[t] || [], sp), price, today);
    Object.keys(ind).forEach(function (k) { row[k] = ind[k]; });
    rows.push(pick_(row, SHARE_F));
  });
  if (rows.length < 50) { log_('Акции: собрано подозрительно мало бумаг (' + rows.length + '), прежние данные оставлены'); return; }
  var size = saveJson_('shares', { updated: today, fund: PropertiesService.getScriptProperties().getProperty('FUND_UPDATED') || '', fields: SHARE_F, rows: rows });
  log_('Акции: бумаг ' + rows.length + (noHist ? ', без истории цен ' + noHist : '') + ', объём ' + Math.round(size / 1024) + ' КБ');
}

// ---------- фонды ----------
var FUND_F = ['t', 'name', 'sector', 'kind', 'uk', 'qual', 'isin', 'list', 'lot', 'age', 'price', 'd1', 'w1', 'm1', 'm3', 'm6', 'ytd', 'y1',
  'y3', 'y5', 'dhi', 'turn', 'aum', 'nav', 'prem', 'vol', 'rsi', 'ma50', 'ma200', 'hi', 'lo', 'sp', 'bd'];
var FUND_KIND = { J: 'Биржевой', '9': 'Открытый', A: 'Интервальный', B: 'Закрытый' };
// Класс активов по словам в названии: порядок важен, первое совпадение побеждает.
var FUND_CLASSES = [
  ['Золото и металлы', /золот|gold|платин|паллад|platinum|серебр|silver|драгоцен/],
  ['Денежный рынок', /денеж|ликвидн|money|liquid|сберегат|накопит|cumulative|ежедневный процент|daily rate|сбереж/],
  ['Облигации', /обл|офз|bond|ofz|флоатер|floater|купон|coupon|долг|казначейск|суборд|вдо /],
  ['Недвижимость', /недвиж|аренд|склад|real estate|рентал|rental|логист|девелоп|строй|строит|жил|офис|торгов|парус|земл|лофт|промышленные парки|атриум/],
  ['Смешанные', /вечный|all-weather|всепогод|портфел|portfolio|моя цель|консерват|осторожн|умный|сбаланс|balanced|инфляц|inflation|лежебок|смеш/],
  ['Альтернативные стратегии', /хедж|hedge/],
  ['Акции', /акци|индекс|index|equit|stock|shares|голуб|blue chip|imoex|див|ит лидер|ipo|айпио|первичных размещ|технолог|интеллект|квант|quant|рост|growth|халяль|esg|устойч|ответствен/]
];
// Ручные поправки там, где название вводит в заблуждение (решение Юлии 24.10.2026): по составу это облигационные фонды.
var FUND_CLASS_FIX = { SAFE: 'Облигации', INFL: 'Облигации', SBRS: 'Облигации' };
function fundClass_(name, lat, t) {
  if (FUND_CLASS_FIX[t]) return FUND_CLASS_FIX[t];
  var text = (name + ' ' + lat).toLowerCase();
  for (var i = 0; i < FUND_CLASSES.length; i++) if (FUND_CLASSES[i][1].test(text)) return FUND_CLASSES[i][0];
  return 'Другое';
}
function fundName_(n) {
  n = String(n || '').replace(/\s+/g, ' ').trim();
  return n.replace(/^(БПИФ|ОПИФ|ИПИФ|ЗПИФ|ПИФ)\s*(РФИ\s*)?/, '').replace(/^[\s\-–]+|[\s\-–]+$/g, '') || n;
}
function shortUk_(s) {
  var n = String(s || '').replace(/[«»]/g, '"');
  n = n.replace(/(общество с ограниченной ответственностью|публичное акционерное общество|акционерное общество|закрытое)/ig, '');
  n = n.replace(/(?<![а-яё])(управляющая компания|ук)(?![а-яё])/ig, '').replace(/"/g, ' ');
  return n.replace(/\s+/g, ' ').replace(/^[\s\-]+|[\s\-]+$/g, '') || 'не указана';
}
// ---------- СЧА и стоимость пая закрытых фондов недвижимости с сайтов УК (29.10.2026) ----------
// Мосбиржа для ЗПИФ расчётную стоимость не публикует. Берём со страниц фондов «ПАРУС» и «Современные Фонды Недвижимости»;
// последние удачные значения хранятся на листе SH_ZPIF (если сайт УК не ответил, остаются прежние).
var SH_ZPIF = 'Фонды: СЧА с сайтов УК';
var ZPIF_SRC = [
  ['parus', 'https://parus.properties/funds/ozon/'], ['parus', 'https://parus.properties/funds/sberlogistika/'],
  ['parus', 'https://parus.properties/funds/nordway/'], ['parus', 'https://parus.properties/funds/logistic/'],
  ['parus', 'https://parus.properties/funds/dvintsev/'], ['parus', 'https://parus.properties/funds/krasnoyarsk/'],
  ['parus', 'https://parus.properties/funds/nino/'], ['parus', 'https://parus.properties/funds/zolya/'],
  ['parus', 'https://parus.properties/funds/mv/'], ['parus', 'https://parus.properties/funds/triumph/'],
  ['sfn', 'https://www.sfn-am.ru/funds/zpifn-arendnyy-biznes-7'], ['sfn', 'https://www.sfn-am.ru/funds/fond-sovremennyi-arendnyi-biznes-3'],
  ['sfn', 'https://www.sfn-am.ru/funds/zpifn-sovremennaya-kommercheskaya-nedvizhimost']
];
function zpifNum_(s) { return Number(String(s).replace(/[\s ]/g, '').replace(',', '.')); }
function zpifLines_(html) {
  return String(html).replace(/<script[\s\S]*?<\/script>|<style[\s\S]*?<\/style>/g, '').replace(/<[^>]+>/g, '\n')
    .replace(/&nbsp;/g, ' ').split('\n').map(function (x) { return x.replace(/\s+/g, ' ').trim(); }).filter(String);
}
function zpifParse_(kind, html) {
  var L = zpifLines_(html), isin = (String(html).match(/RU[0-9A-Z]{10}/) || [])[0], r = { isin: null, nav: null, aum: null, date: null };
  var m = String(html).match(/ISIN[^R]{0,80}(RU[0-9A-Z]{10})/); r.isin = m ? m[1] : isin;
  for (var i = 0; i < L.length; i++) {
    if (kind === 'parus') {
      if (r.aum === null && /^КАПИТАЛ ИНВЕСТОРОВ$/i.test(L[i]) && /^[\d ,]+ (млрд|млн)/.test(L[i + 1] || ''))
        r.aum = zpifNum_(L[i + 1].split(' м')[0]) * (/млрд/.test(L[i + 1]) ? 1 : 0.001);
      if (r.nav === null && /^\d{1,2} [а-я]{3} \d{4}$/.test(L[i]) && /^[\d ]+$/.test(L[i + 1] || '') && L.slice(Math.max(0, i - 30), i).join(' ').indexOf('стоимость пая') >= 0)
        r.nav = zpifNum_(L[i + 1]);
    } else {
      var d = L[i].match(/^Расчетная стоимость инвестиционного пая на (\d\d)\.(\d\d)\.(\d{4})/);
      if (d && r.nav === null) { r.nav = zpifNum_((L[i - 1] || '').replace(/Руб\.?/, '')); r.date = d[3] + '-' + d[2] + '-' + d[1]; }
      if (r.aum === null && /^Стоимость чистых активов фонда на/.test(L[i])) r.aum = round_(zpifNum_((L[i - 1] || '').replace(/Руб\.?/, '')) / 1e9, 2);
    }
  }
  if (r.aum !== null) r.aum = round_(r.aum, 2);
  return r;
}
// возвращает { ISIN: {nav, aum, date} } и обновляет лист
function zpifNav_(today) {
  var old = {};
  read_(SH_ZPIF).forEach(function (r) { old[r[0]] = { nav: Number(r[1]) || null, aum: Number(r[2]) || null, date: String(r[3] || ''), src: r[4], upd: r[5] }; });
  var res;
  try { res = UrlFetchApp.fetchAll(ZPIF_SRC.map(function (x) { return { url: x[1], muteHttpExceptions: true, followRedirects: true, headers: { 'User-Agent': 'Mozilla/5.0' } }; })); }
  catch (e) { res = []; log_('Фонды: сайты УК не ответили (' + e + '), СЧА ЗПИФ прежние'); }
  var ok = 0;
  res.forEach(function (r, i) {
    if (!r || r.getResponseCode() !== 200) return;
    var p = zpifParse_(ZPIF_SRC[i][0], r.getContentText());
    if (!p.isin || !p.nav) return;
    old[p.isin] = { nav: p.nav, aum: p.aum, date: p.date || today, src: ZPIF_SRC[i][1], upd: today }; ok++;
  });
  write_(SH_ZPIF, ['ISIN', 'Стоимость пая по данным УК, ₽', 'СЧА, млрд ₽', 'На дату', 'Источник', 'Собрано'],
    Object.keys(old).sort().map(function (k) { var o = old[k]; return [k, o.nav, o.aum, o.date, o.src, o.upd]; }));
  log_('Фонды: СЧА ЗПИФ с сайтов УК обновлено у ' + ok + ' из ' + ZPIF_SRC.length);
  return old;
}
function refreshFunds() {
  var today = ymd_(new Date());
  var j = fetchJson_(ISS + '/engines/stock/markets/shares/boards/TQBR/securities.json?iss.meta=off&iss.only=securities,marketdata' +
    '&securities.columns=SECID,SECNAME,LATNAME,ISIN,LOTSIZE,LISTLEVEL,PREVPRICE,PREVLEGALCLOSEPRICE,SECTYPE&marketdata.columns=SECID,LAST,LCLOSEPRICE');
  if (!j) { log_('Фонды: Мосбиржа не ответила, обновление пропущено'); return; }
  var secs = {}, md = {};
  rows_(j.securities).forEach(function (r) { if (FUND_KIND[r.SECTYPE]) secs[r.SECID] = r; });
  rows_(j.marketdata).forEach(function (r) { md[r.SECID] = r; });
  var ids = Object.keys(secs).sort();
  // расчётная стоимость пая: индексы режима INAV
  var nav = fetchJson_(ISS + '/engines/stock/markets/index/boards/INAV/securities.json?iss.meta=off&iss.only=securities,marketdata' +
    '&securities.columns=SECID,CURRENCYID&marketdata.columns=SECID,CURRENTVALUE,CAPITALIZATION,TRADEDATE');
  var inav = {}, cur = {};
  rows_(nav && nav.securities).forEach(function (r) { cur[r.SECID] = r.CURRENCYID; });
  rows_(nav && nav.marketdata).forEach(function (r) { r.cur = cur[r.SECID]; inav[r.SECID] = r; });
  // управляющие компании и записи биржи о дроблениях
  var uk = {};
  pages_('/securities.json?engine=stock&market=shares&is_trading=1&group_by=group&group_by_filter=stock_ppif', 'securities')
    .forEach(function (r) { uk[r.secid] = r.emitent_title; });
  // справочник: доступность и год выпуска паёв (читаем один раз на фонд)
  var ref = {};
  read_(SH_FREF).forEach(function (r) { ref[r[0]] = r; });
  var need = ids.filter(function (t) { return !ref[t]; });
  fetchAll_(need.map(function (t) { return ISS + '/securities/' + encodeURIComponent(t) + '.json?iss.meta=off&iss.only=description'; })).forEach(function (res, i) {
    if (!res || !res.description) return;
    var d = {}; res.description.data.forEach(function (r) { d[r[0]] = r[2]; });
    ref[need[i]] = [need[i], String(d.ISQUALIFIEDINVESTORS) === '1' ? 'да' : 'нет', String(d.ISSUEDATE || '').slice(0, 10), 'нет'];
  });
  var from = addDays_(today, -430), fromW = addDays_(today, -365 * 5 - 40);
  var daily = candles_('shares', ids, 24, from), weekly = candles_('shares', ids, 7, fromW);
  // У фондов, которые торгуются ещё и в юанях, общий запрос отдаёт юаневую историю: для них берём рублёвый режим TQBR.
  var other = ids.filter(function (t) {
    var d = daily[t], p = pos_((md[t] || {}).LCLOSEPRICE, (md[t] || {}).LAST, secs[t].PREVPRICE);
    return (ref[t] && ref[t][3] === 'да') || (d && d.length && p && Math.abs(d[d.length - 1][1] / p - 1) > 0.3);
  });
  if (other.length) {
    var d2 = candles_('shares', other, 24, from, 'TQBR'), w2 = candles_('shares', other, 7, fromW, 'TQBR');
    other.forEach(function (t) { daily[t] = d2[t]; weekly[t] = w2[t]; if (ref[t]) ref[t][3] = 'да'; });
  }
  write_(SH_FREF, ['Тикер', 'Только квал', 'Дата выпуска паёв', 'История из режима TQBR'], Object.keys(ref).sort().map(function (t) { return ref[t]; }));
  var zp = zpifNav_(today);
  var fresh = addDays_(today, -10), rows = [];
  ids.forEach(function (t) {
    var s = secs[t], m = md[t] || {}, r = ref[t] || [t, 'нет', '', 'нет'];
    var price = pos_(m.LCLOSEPRICE, m.LAST, s.PREVLEGALCLOSEPRICE, s.PREVPRICE);
    if (!price) return;
    var exact = s.SECTYPE === 'J';
    var dsp = splits_(daily[t] || [], exact), wsp = splits_(weekly[t] || [], exact);
    // расчётная стоимость: индекс «тикер + буква» в рублях, свежий, ближайший к цене; расхождение больше 15% не показываем
    var best = null;
    Object.keys(inav).forEach(function (k) {
      var x = inav[k];
      if (k.indexOf(t) !== 0 || k.length > t.length + 2 || x.cur !== 'RUB' || !x.CURRENTVALUE || String(x.TRADEDATE || '') < fresh) return;
      if (!best || Math.abs(x.CURRENTVALUE / price - 1) < Math.abs(best.CURRENTVALUE / price - 1)) best = x;
    });
    if (best && Math.abs(best.CURRENTVALUE / price - 1) > 0.15) best = null;
    var month = (daily[t] || []).slice(-30), turn = 0;
    month.forEach(function (p) { turn += p[2]; });
    var row = {
      t: t, name: fundName_(s.SECNAME), sector: fundClass_(s.SECNAME, s.LATNAME || '', t), kind: FUND_KIND[s.SECTYPE], uk: shortUk_(uk[t]),
      qual: r[1] === 'да' ? 'только для квалифицированных' : 'для всех', isin: s.ISIN, list: s.LISTLEVEL, lot: s.LOTSIZE,
      age: r[2] ? Number(String(r[2]).slice(0, 4)) : null, price: price,
      d1: pos_(m.LCLOSEPRICE, m.LAST) && s.PREVPRICE ? round_((price / s.PREVPRICE - 1) * 100, 1) : null,
      turn: month.length && turn > 0 ? round_(turn / month.length / 1e6, 2) : null,
      aum: best && best.CAPITALIZATION ? round_(best.CAPITALIZATION / 1e9, 2) : null,
      nav: best ? best.CURRENTVALUE : null, prem: best ? round_((price / best.CURRENTVALUE - 1) * 100, 1) : null,
      sp: wsp.length ? wsp : null, bd: r[3] === 'да' ? 'TQBR' : null
    };
    // у ЗПИФ без индекса INAV — СЧА и стоимость пая с сайта УК
    var z = !best && zp[s.ISIN];
    if (z && z.nav) { row.nav = z.nav; row.aum = z.aum; row.prem = round_((price / z.nav - 1) * 100, 1); }
    var ind = indicators_(adjust_(daily[t] || [], dsp), adjust_(weekly[t] || [], wsp), price, today);
    Object.keys(ind).forEach(function (k) { row[k] = ind[k]; });
    rows.push(pick_(row, FUND_F));
  });
  if (rows.length < 50) { log_('Фонды: собрано подозрительно мало бумаг (' + rows.length + '), прежние данные оставлены'); return; }
  var size = saveJson_('funds', { updated: today, fields: FUND_F, rows: rows });
  log_('Фонды: бумаг ' + rows.length + ', история из TQBR у ' + other.length + ', объём ' + Math.round(size / 1024) + ' КБ');
}

// ---------- индексы ----------
var INDEX_F = ['t', 'name', 'sector', 'full', 'ord', 'price', 'd1', 'w1', 'm1', 'm3', 'm6', 'ytd', 'y1', 'y3', 'y5', 'dhi', 'n', 'cap',
  'vol', 'rsi', 'ma50', 'ma200', 'hi', 'lo', 'comp'];
var IG1 = 'Акции: рынок в целом', IG2 = 'Акции: отрасли', IG3 = 'Облигации', IG4 = 'Золото, недвижимость, ставки';
// код, короткое название, группа, показывать ли состав по бумагам
var INDEXES = [
  ['IMOEX', 'Индекс МосБиржи', IG1, 1], ['MCFTR', 'Индекс МосБиржи полной доходности', IG1, 0], ['RTSI', 'Индекс РТС (в долларах)', IG1, 1],
  ['MOEXBC', 'Голубые фишки', IG1, 1], ['MOEX10', 'МосБиржи 10', IG1, 1], ['MRBC', 'МосБиржи 15', IG1, 1], ['MOEXBMI', 'Широкий рынок', IG1, 1],
  ['MCXSM', 'Средняя и малая капитализация', IG1, 1], ['MOEXINN', 'Инновации', IG1, 1], ['MVBI', 'Создание стоимости', IG1, 1],
  ['IRDIV', 'Дивидендные акции (индекс УК «ДОХОДЪ»)', IG1, 0],
  ['MOEXOG', 'Нефть и газ', IG2, 1], ['MOEXFN', 'Финансы', IG2, 1], ['MOEXMM', 'Металлы и добыча', IG2, 1], ['MOEXEU', 'Электроэнергетика', IG2, 1],
  ['MOEXCN', 'Потребительский сектор', IG2, 1], ['MOEXCH', 'Химия и нефтехимия', IG2, 1], ['MOEXIT', 'Информационные технологии', IG2, 1],
  ['MOEXRE', 'Недвижимость', IG2, 1], ['MOEXTN', 'Транспорт', IG2, 1], ['MOEXTL', 'Телекоммуникации', IG2, 1],
  ['RGBITR', 'ОФЗ, полная доходность', IG3, 0], ['RGBI', 'ОФЗ, только цена', IG3, 0],
  ['RUCBTRNS', 'Корпоративные облигации, полная доходность', IG3, 0], ['RUCBCPNS', 'Корпоративные облигации, только цена', IG3, 0],
  ['RUMBTRNS', 'Муниципальные облигации, полная доходность', IG3, 0], ['RUFLGBITR', 'ОФЗ с плавающим купоном', IG3, 0],
  ['RUFLCBTR', 'Корпоративные с плавающим купоном', IG3, 0], ['RUCBHYTR', 'Высокодоходные облигации', IG3, 0], ['RUCNYTR', 'Облигации в юанях', IG3, 0],
  ['RUGOLD', 'Золото', IG4, 0], ['MREDC', 'Недвижимость Москвы (Домклик)', IG4, 0], ['RUSFAR', 'Ставка денежного рынка RUSFAR', IG4, 0],
  ['RVI', 'Индекс волатильности RVI', IG4, 0]
];
function refreshIndices() {
  var today = ymd_(new Date()), ids = INDEXES.map(function (x) { return x[0]; });
  var j = fetchJson_(ISS + '/engines/stock/markets/index/securities.json?iss.meta=off&iss.only=securities,marketdata&securities=' + ids.join(',') +
    '&securities.columns=SECID,NAME&marketdata.columns=SECID,CURRENTVALUE,LASTCHANGEPRC,CAPITALIZATION,TRADEDATE');
  if (!j) { log_('Индексы: Мосбиржа не ответила, обновление пропущено'); return; }
  var secs = {}, md = {};
  rows_(j.securities).forEach(function (r) { secs[r.SECID] = r; });
  rows_(j.marketdata).forEach(function (r) { md[r.SECID] = r; });
  var known = {};
  rows_((fetchJson_(ISS + '/statistics/engines/stock/markets/index/analytics.json?iss.meta=off') || {}).indices).forEach(function (r) { known[r.indexid] = 1; });
  var withComp = ids.filter(function (t) { return known[t] && secs[t]; }), comp = {};
  function compPage(list, start) {
    return fetchAll_(list.map(function (t) {
      return ISS + '/statistics/engines/stock/markets/index/analytics/' + t + '.json?iss.meta=off&limit=100&start=' + start;
    })).map(function (res) { return rows_(res && res.analytics); });
  }
  var first = compPage(withComp, 0), more = [];
  withComp.forEach(function (t, i) { comp[t] = first[i]; if (first[i].length === 100) more.push(t); });
  compPage(more, 100).forEach(function (part, i) { comp[more[i]] = comp[more[i]].concat(part); });
  var daily = candles_('index', ids, 24, addDays_(today, -430)), weekly = candles_('index', ids, 7, addDays_(today, -365 * 5 - 40));
  var stale = addDays_(today, -15), rows = [], skipped = [];
  INDEXES.forEach(function (x, order) {
    var t = x[0], s = secs[t], m = md[t] || {}, price = m.CURRENTVALUE;
    if (!s || !price || String(m.TRADEDATE || '') < stale) { skipped.push(t); return; }
    var seen = {}, list = (comp[t] || []).filter(function (r) { if (seen[r.ticker]) return false; seen[r.ticker] = 1; return true; })
      .sort(function (a, b) { return (b.weight || 0) - (a.weight || 0); }).map(function (r) { return [r.ticker, r.shortnames, r.weight]; });
    var row = {
      t: t, name: x[1], sector: x[2], full: s.NAME, ord: order, price: price,
      d1: typeof m.LASTCHANGEPRC === 'number' ? round_(m.LASTCHANGEPRC, 1) : null,
      n: list.length || null, cap: m.CAPITALIZATION ? round_(m.CAPITALIZATION / 1e9, 1) : null, comp: x[3] && list.length ? list : null
    };
    var ind = indicators_(daily[t] || [], weekly[t] || [], price, today);
    Object.keys(ind).forEach(function (k) { row[k] = ind[k]; });
    if (row.hi !== null) { row.hi = round_(row.hi, 2); row.lo = round_(row.lo, 2); }
    rows.push(pick_(row, INDEX_F));
  });
  if (rows.length < 10) { log_('Индексы: собрано подозрительно мало (' + rows.length + '), прежние данные оставлены'); return; }
  var size = saveJson_('indices', { updated: today, fields: INDEX_F, rows: rows });
  log_('Индексы: ' + rows.length + (skipped.length ? ', без свежих данных ' + skipped.join(', ') : '') + ', объём ' + Math.round(size / 1024) + ' КБ');
}

// ---------- кривая ОФЗ и ставки ----------
var TENORS = [0.25, 0.5, 0.75, 1, 2, 3, 5, 7, 10, 15, 20];
var CURVE_FIRST = '2014-01-10';      // раньше биржа кривую не отдаёт
/** Кривая на ближайший торговый день не позже каждой из дат: {дата запроса: [дата кривой, значения по срокам]}. */
function curves_(dates) {
  var out = {}, pending = dates.map(function (d) { return { want: d, ask: d }; });
  for (var round = 0; round < 9 && pending.length; round++) {
    var res = fetchAll_(pending.map(function (p) { return ISS + '/engines/stock/zcyc.json?iss.meta=off&iss.only=yearyields&date=' + p.ask; }));
    var next = [];
    pending.forEach(function (p, i) {
      var data = res[i] && res[i].yearyields ? res[i].yearyields.data : [];
      if (data.length && data[0][0] === p.ask) {
        var by = {}; data.forEach(function (r) { by[r[2]] = round_(r[3], 2); });
        out[p.want] = [p.ask, TENORS.map(function (t) { return by[t] === undefined ? null : by[t]; })];
      } else { p.ask = addDays_(p.ask, -1); next.push(p); }
    });
    pending = next;
  }
  return out;
}
function refreshCurve() {
  var started = Date.now(), today = ymd_(new Date());
  // 1. история по пятницам: дозагружаем недостающие недели
  var have = {}, hist = read_(SH_CURVE).map(function (r) { return r.map(function (v, i) { return i < 2 ? day_(v) : (v === '' ? null : Number(v)); }); });
  hist.forEach(function (r) { have[r[0]] = 1; });
  var fr = parse_(today);
  fr.setUTCDate(fr.getUTCDate() - ((fr.getUTCDay() + 2) % 7 || 7));      // последняя прошедшая пятница
  var missing = [];
  for (var d = fr.toISOString().slice(0, 10); d >= CURVE_FIRST; d = addDays_(d, -7)) if (!have[d]) missing.push(d);
  for (var a = 0; a < missing.length && Date.now() - started < TIME_LIMIT_MS; a += 80) {
    var part = missing.slice(a, a + 80), got = curves_(part);
    // недели без торгов (длинные праздники, март 2022) записываем пустой строкой, чтобы не спрашивать их снова
    part.forEach(function (want) { hist.push(got[want] ? [want, got[want][0]].concat(got[want][1]) : [want, 'нет данных'].concat(TENORS.map(function () { return null; }))); });
  }
  hist.sort(function (x, y) { return x[0] < y[0] ? -1 : 1; });
  write_(SH_CURVE, ['Пятница', 'Дата кривой'].concat(TENORS.map(function (t) { return 'Срок ' + t; })), hist);
  var left = missing.filter(function (m) { return !hist.some(function (r) { return r[0] === m; }); }).length;
  var good = hist.filter(function (r) { return /^\d{4}-\d\d-\d\d$/.test(String(r[1])); });
  // 2. кривая сегодня и в прошлом
  var marks = [['now', 'Сегодня', 0], ['m1', 'Месяц назад', 30], ['m6', 'Полгода назад', 182], ['y1', 'Год назад', 365], ['y3', '3 года назад', 1095], ['y5', '5 лет назад', 1825]];
  var cv = curves_(marks.map(function (m) { return addDays_(today, -m[2]); })), curves = [];
  marks.forEach(function (m) {
    var c = cv[addDays_(today, -m[2])];
    if (c) curves.push({ k: m[0], label: m[1], date: c[0], vals: c[1] });
  });
  if (!curves.length || curves[0].k !== 'now') { log_('Кривая: биржа не отдала сегодняшнюю кривую, прежние данные оставлены'); return; }
  // 3. выпуски, по которым биржа строит кривую
  var sj = fetchJson_(ISS + '/engines/stock/zcyc.json?iss.meta=off&iss.only=securities&date=' + curves[0].date), bonds = [];
  rows_(sj && sj.securities).forEach(function (s) {
    var y = s.trdyield || s.clcyield;
    if (!s.expdate || !y) return;
    var years = (parse_(s.expdate) - parse_(curves[0].date)) / 864e5 / 365.25;
    if (years >= 0.2 && years <= 20.5) bonds.push([s.shortname || s.secid, round_(years, 2), round_(y, 2)]);
  });
  // 4. RUSFAR недельными значениями и ключевая ставка Банка России (точки изменения)
  var rc = candles_('index', ['RUSFAR'], 7, '2014-01-01').RUSFAR || [];
  var prev = {}; try { prev = JSON.parse(loadJson_('curve') || '{}'); } catch (e) {}
  var key = prev.key || [], keyLast = prev.key_last || '';
  try {
    var now = new Date();
    var page = UrlFetchApp.fetch('https://www.cbr.ru/hd_base/KeyRate/?UniDbQuery.Posted=True&UniDbQuery.From=01.01.2014&UniDbQuery.To=' +
      Utilities.formatDate(now, 'Europe/Moscow', 'dd.MM.yyyy'), { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } }).getContentText();
    var re = /<td>(\d\d)\.(\d\d)\.(\d{4})<\/td>\s*<td>([\d,]+)<\/td>/g, m, dailyKey = [];
    while ((m = re.exec(page))) dailyKey.push([m[3] + '-' + m[2] + '-' + m[1], Number(m[4].replace(',', '.'))]);
    dailyKey.sort(function (x, y) { return x[0] < y[0] ? -1 : 1; });
    if (dailyKey.length > 100) {
      key = dailyKey.filter(function (r, i) { return i === 0 || r[1] !== dailyKey[i - 1][1]; });
      keyLast = dailyKey[dailyKey.length - 1][0];
    } else log_('Кривая: не удалось прочитать ключевую ставку с сайта Банка России, оставлена прежняя');
  } catch (e) { log_('Кривая: сайт Банка России не ответил, ключевая ставка оставлена прежняя'); }
  if (!key.length) { log_('Кривая: ключевой ставки нет совсем, данные не сохранены'); return; }
  var size = saveJson_('curve', {
    updated: curves[0].date, tenors: TENORS, curves: curves, bonds: bonds,
    hist: { dates: good.map(function (r) { return r[1]; }), y1: good.map(function (r) { return r[5]; }), y5: good.map(function (r) { return r[8]; }), y10: good.map(function (r) { return r[10]; }) },
    rusfar: rc.map(function (r) { return [r[0], r[1]]; }), key: key, key_last: keyLast
  });
  log_('Кривая: недель истории ' + good.length + (left ? ', осталось загрузить ' + left : '') + ', выпусков ' + bonds.length + ', объём ' + Math.round(size / 1024) + ' КБ');
  if (left) againMarket_('refreshCurve');
}
function againMarket_(fn) {
  ScriptApp.newTrigger(fn + 'More').timeBased().after(60 * 1000).create();
}
function refreshCurveMore() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'refreshCurveMore') ScriptApp.deleteTrigger(t); });
  refreshCurve();
}

// ---------- календарь купонов ----------
function shortIssuer_(s) {
  return String(s || '').replace(/\\"/g, '"').replace(/Министерство финансов Российской Федерации/i, 'Минфин России')
    .replace(/(публичное|открытое|закрытое|непубличное)?\s*акционерное общество/ig, '').replace(/общество с ограниченной ответственностью/ig, '')
    .replace(/[«»"()]/g, ' ').replace(/\s+/g, ' ').trim();
}
var COUPON_DAYS = 90;
function refreshCoupons() {
  var today = ymd_(new Date()), till = addDays_(today, COUPON_DAYS);
  var base = ISS + '/statistics/engines/stock/markets/bonds/bondization.json?iss.meta=off&from=' + today + '&till=' + till + '&limit=100';
  function all(block) {
    var head = fetchJson_(base + '&iss.only=' + block + ',' + block + '.cursor&start=0');
    if (!head || !head[block]) return null;
    var total = head[block + '.cursor'].data[0][1], out = rows_(head[block]), urls = [];
    for (var s = 100; s < total; s += 100) urls.push(base + '&iss.only=' + block + '&start=' + s);
    var ok = true;
    fetchAll_(urls).forEach(function (res) { if (res && res[block]) out = out.concat(rows_(res[block])); else ok = false; });
    return ok ? out : null;
  }
  var cp = all('coupons'), am = all('amortizations'), of = all('offers');
  if (!cp || !am || !of) { log_('Купоны: Мосбиржа ответила не полностью, прежние данные оставлены'); return; }
  // облигации из скринера
  var bond = {};
  read_(SH_B).forEach(function (r) { var o = {}; B_KEYS.forEach(function (k, i) { o[k] = r[i] instanceof Date ? ymd_(r[i]) : r[i]; }); bond[o.id] = o; });
  if (Object.keys(bond).length < 100) { log_('Купоны: лист облигаций пуст, обновление пропущено'); return; }
  var B = [], at = {}, UNIT = { RUR: 'RUB', SUR: 'RUB' };
  function ref(id) {
    var o = bond[id];
    // структурные ноты, дефолты и бумаги без торгов в календарь не берём
    if (!o || o.def === 'да' || o.liq === '' || o.ctype === 'Структурная') return -1;
    if (at[id] === undefined) {
      var rt = String(o.rating || '').split('|'), fl = String(o.ctype || '').indexOf('Плавающий') === 0;
      at[id] = B.length;
      B.push([id, o.name, shortIssuer_(o.issuer).slice(0, 48), o.type || '', rt[0] || '', rt.length > 1 && rt[1] !== '' ? Number(rt[1]) : -1,
        o.qual === 'да' ? 1 : 0, Number(o.liq), !fl && Number(o.y) > 0 ? round_(Number(o.y), 1) : null, o.ctype || '',
        o.qual !== 'да' && Number(o.liq) >= 20 ? 1 : 0]);
    }
    return at[id];
  }
  var C = [], A = [], O = [];
  cp.forEach(function (r) { var i = ref(r.secid); if (i >= 0) C.push([r.coupondate, i, r.value, r.valueprc, UNIT[r.faceunit] || r.faceunit || 'RUB']); });
  am.forEach(function (r) { var i = ref(r.secid); if (i >= 0 && r.value) A.push([r.amortdate, i, r.value, r.valueprc, UNIT[r.faceunit] || r.faceunit || 'RUB', r.data_source === 'maturity' ? 1 : 0]); });
  of.forEach(function (r) { var i = ref(r.secid); if (i >= 0) O.push([r.offerdate, i, r.price, r.offerdatestart || '', r.offerdateend || '', r.offertype || '']); });
  Object.keys(bond).forEach(function (id) { var o = bond[id]; if (o.qual !== 'да' && Number(o.liq) >= 20) ref(id); });   // все бумаги для расчёта «Мои купоны»
  [C, A, O].forEach(function (list) { list.sort(function (x, y) { return x[0] < y[0] ? -1 : x[0] > y[0] ? 1 : B[y[1]][7] - B[x[1]][7]; }); });
  var size = saveJson_('coupons', { updated: today, till: till, bonds: B, coupons: C, amort: A, offers: O });
  log_('Купоны: облигаций ' + B.length + ', купонов ' + C.length + ', погашений ' + A.length + ', оферт ' + O.length + ', объём ' + Math.round(size / 1024) + ' КБ');
}

// ---------- курсы валют, золото, криптовалюта ----------
// код, идентификатор Банка России, название, за сколько единиц показываем курс
var FX_CBR = [
  ['USD', 'R01235', 'Доллар США', 1], ['EUR', 'R01239', 'Евро', 1], ['CNY', 'R01375', 'Китайский юань', 1],
  ['GBP', 'R01035', 'Фунт стерлингов', 1], ['CHF', 'R01775', 'Швейцарский франк', 1], ['JPY', 'R01820', 'Японская иена, за 100', 100],
  ['AED', 'R01230', 'Дирхам ОАЭ', 1], ['TRY', 'R01700J', 'Турецкая лира', 1], ['KZT', 'R01335', 'Казахстанский тенге, за 100', 100],
  ['BYN', 'R01090B', 'Белорусский рубль', 1], ['HKD', 'R01200', 'Гонконгский доллар', 1]
];
var FX_METALS = [['1', 'AU', 'Золото, за грамм'], ['2', 'AG', 'Серебро, за грамм'], ['3', 'PT', 'Платина, за грамм'], ['4', 'PD', 'Палладий, за грамм']];
var FX_MOEX = [['CNYRUB_TOM', 'Юань на бирже'], ['GLDRUB_TOM', 'Золото на бирже, за грамм'], ['SLVRUB_TOM', 'Серебро на бирже, за грамм']];
var FX_CRYPTO = [['MOEXBTC', 'Биткоин, индекс МосБиржи'], ['MOEXETH', 'Эфир, индекс МосБиржи']];
var FX_BASE = '2014-01-01';
function cbrText_(url) {
  var r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
  return r.getResponseCode() === 200 ? r.getContentText('windows-1251') : '';
}
function ru2iso_(d) { return d.slice(6) + '-' + d.slice(3, 5) + '-' + d.slice(0, 2); }
/** Прореживание: последние 190 дней по дням, до 5 лет по неделям, старше по месяцам. Формат точки: [дней от FX_BASE, значение]. */
function fxThin_(series, today) {
  var dayFrom = addDays_(today, -190), weekFrom = addDays_(today, -365 * 5 - 10), out = [], lastKey = '';
  series.forEach(function (p) {
    var key;
    if (p[0] >= dayFrom) key = p[0];
    else if (p[0] >= weekFrom) { var d = parse_(p[0]); d.setUTCDate(d.getUTCDate() - ((d.getUTCDay() + 6) % 7)); key = 'w' + d.toISOString().slice(0, 10); }
    else key = 'm' + p[0].slice(0, 7);
    var point = [Math.round((parse_(p[0]) - parse_(FX_BASE)) / 864e5), p[1]];
    if (key === lastKey) out[out.length - 1] = point; else out.push(point);     // в неделе и месяце оставляем последнее значение
    lastKey = key;
  });
  return out;
}
function fxItem_(id, name, group, unit, src, series, today) {
  if (!series || series.length < 2) return null;
  var last = series[series.length - 1], price = last[1];
  function at(iso) { if (series[0][0] > iso) return null; var v = null; for (var i = 0; i < series.length && series[i][0] <= iso; i++) v = series[i][1]; return v; }
  function chg(then) { return then ? round_((price / then - 1) * 100, 1) : null; }
  var o = { id: id, name: name, group: group, unit: unit, src: src, price: price, date: last[0], d1: chg(series[series.length - 2][1]) };
  [['w1', 7], ['m1', 30], ['m3', 91], ['m6', 182], ['y1', 365], ['y3', 1095], ['y5', 1825]].forEach(function (p) { o[p[0]] = chg(at(addDays_(last[0], -p[1]))); });
  o.ytd = chg(at((Number(last[0].slice(0, 4)) - 1) + '-12-31'));
  var yearAgo = addDays_(last[0], -365), year = series.filter(function (p) { return p[0] >= yearAgo; }).map(function (p) { return p[1]; });
  o.hi = Math.max.apply(null, year); o.lo = Math.min.apply(null, year);
  o.hist = fxThin_(series, today);
  return o;
}
function fxCandles_(engine, market, id, interval, from) {
  var out = [], start = 0;
  for (;;) {
    var j = fetchJson_(ISS + '/engines/' + engine + '/markets/' + market + '/securities/' + id + '/candles.json?iss.meta=off&interval=' + interval +
      '&from=' + from + '&start=' + start + '&candles.columns=close,begin');
    var part = j && j.candles ? j.candles.data : [];
    part.forEach(function (r) { if (r[0]) out.push([String(r[1]).slice(0, 10), r[0]]); });
    start += 500;
    if (part.length < 500 || start > 5000) return out;
  }
}
/** Недельные свечи с 2014 года плюс дневные за последний год: одна лента по возрастанию дат. */
function fxMoexSeries_(engine, market, id, today) {
  var daily = fxCandles_(engine, market, id, 24, addDays_(today, -400)), weekly = fxCandles_(engine, market, id, 7, FX_BASE);
  var first = daily.length ? daily[0][0] : '9999';
  return weekly.filter(function (p) { return p[0] < addDays_(first, -6); }).concat(daily);
}
function refreshFx() {
  var today = ymd_(new Date()), till = Utilities.formatDate(new Date(Date.now() + 864e5), 'Europe/Moscow', 'dd/MM/yyyy'), items = [], fails = [];   // по завтра: ЦБ публикует курс на завтра около 15:30 МСК
  FX_CBR.forEach(function (c) {
    var xml = '', series = [], re = /<Record Date="(\d\d\.\d\d\.\d{4})"[^>]*>\s*<Nominal>(\d+)<\/Nominal>\s*<Value>([\d,]+)<\/Value>(?:\s*<VunitRate>([\d,]+)<\/VunitRate>)?/g, m;
    try { xml = cbrText_('https://www.cbr.ru/scripts/XML_dynamic.asp?date_req1=01/01/2014&date_req2=' + till + '&VAL_NM_RQ=' + c[1]); } catch (e) {}
    while ((m = re.exec(xml))) {
      var unit = m[4] ? Number(m[4].replace(',', '.')) : Number(m[3].replace(',', '.')) / Number(m[2]);
      series.push([ru2iso_(m[1]), round_(unit * c[3], 4)]);
    }
    var it = fxItem_(c[0], c[2], 'Валюты: курс Банка России', '₽', 'официальный курс Банка России', series, today);
    if (it) items.push(it); else fails.push(c[0]);
  });
  var mx = '';
  try { mx = cbrText_('https://www.cbr.ru/scripts/xml_metall.asp?date_req1=01/01/2014&date_req2=' + till); } catch (e) {}
  var byCode = {}, mre = /<Record Date="(\d\d\.\d\d\.\d{4})" Code="(\d)">\s*<Buy>([\d,]+)<\/Buy>/g, mm;
  while ((mm = mre.exec(mx))) (byCode[mm[2]] = byCode[mm[2]] || []).push([ru2iso_(mm[1]), Number(mm[3].replace(',', '.'))]);
  FX_METALS.forEach(function (c) {
    var s = (byCode[c[0]] || []).sort(function (a, b) { return a[0] < b[0] ? -1 : 1; });
    var it = fxItem_(c[1], c[2], 'Драгоценные металлы: цена Банка России', '₽', 'учётная цена Банка России', s, today);
    if (it) items.push(it); else fails.push(c[1]);
  });
  FX_MOEX.forEach(function (c) {
    var it = fxItem_(c[0], c[1], 'Биржевые торги', '₽', 'торги Московской биржи', fxMoexSeries_('currency', 'selt', c[0], today), today);
    if (it) items.push(it); else fails.push(c[0]);
  });
  FX_CRYPTO.forEach(function (c) {
    var it = fxItem_(c[0], c[1], 'Криптовалюта', '$', 'индекс Московской биржи, в долларах', fxMoexSeries_('stock', 'index', c[0], today), today);
    if (it) items.push(it); else fails.push(c[0]);
  });
  if (items.length < 10) { log_('Курсы: собрано слишком мало (' + items.length + '), прежние данные оставлены'); return; }
  var size = saveJson_('fx', { updated: today, base: FX_BASE, items: items });
  log_('Курсы: показателей ' + items.length + (fails.length ? ', не получены ' + fails.join(', ') : '') + ', объём ' + Math.round(size / 1024) + ' КБ');
}

// ---------- фьючерсы ----------
var FUT_F = ['t', 'name', 'sector', 'asset', 'aname', 'kind', 'near', 'exp', 'days', 'price', 'd1', 'w1', 'm1', 'm3', 'dhi', 'val', 'im', 'lev', 'imp',
  'stepv', 'step', 'lot', 'oi', 'turn', 'fee', 'vol', 'rsi', 'ma50', 'hi', 'lo'];
// Биржа не отдаёт вид базового актива, поэтому группы заданы списками кодов; всё остальное считается акциями.
var FUT_GROUPS = {
  'Индексы': 'MIX MXI MMI RTS RTSM IMOEX MOEXCNY RGBI RGBIF RVI HOME IPO OGI FNI CNI NASD SPYF SP500F QQQF DJ30 R2000 DAX STOX HANG NIKK EM INDIA CHINA BRAZIL KOREA SAUDI AFRICA ARGT SOXQ TLT',
  'Валюты': 'Si Eu ED CNY UCNY USDRUBTOM EURRUBTOM CNYRUBTOM AED AMD BYN HKD INR KZT TRY UCAD UCHF UINR UJPY UKZT UTRY AUDU GBPU ECAD EGBP EJPY USDM EURM',
  'Товары': 'BR BRM WTI NG NGM TTF GOLD GOLDM GL GLDRUBTOM SILV SILVM SL SLVRUBTOM PLT PLTM PLD PLDM COPPER ALUM NICKEL ZINC WHEAT SUGAR SUGR COCOA COFFEE ORANGE AI92 AI95 DTL',
  'Ставки': 'RUONIA KEYRATE',
  'Криптовалюта': 'BTC ETH SOL XRP TRX IBIT ETHA ETHAA BTCUSDF ETHUSDF SOLUSDF XRPUSDF TRXUSDF'
};
var FUT_NAMES = {
  Si: 'доллар к рублю', Eu: 'евро к рублю', CNY: 'юань к рублю', ED: 'евро к доллару', UCNY: 'доллар к юаню', USDRUBTOM: 'доллар к рублю', EURRUBTOM: 'евро к рублю',
  CNYRUBTOM: 'юань к рублю', MIX: 'индекс МосБиржи', MXI: 'индекс МосБиржи, мини', IMOEX: 'индекс МосБиржи', RTS: 'индекс РТС', RTSM: 'индекс РТС, мини',
  RGBI: 'индекс гособлигаций', NASD: 'индекс Nasdaq 100', SPYF: 'индекс S&P 500', BR: 'нефть Brent', BRM: 'нефть Brent, мини', WTI: 'нефть WTI',
  NG: 'природный газ', NGM: 'природный газ, мини', GOLD: 'золото', GOLDM: 'золото, мини', GL: 'золото в рублях', GLDRUBTOM: 'золото в рублях',
  SILV: 'серебро', SILVM: 'серебро, мини', PLT: 'платина', PLD: 'палладий', WHEAT: 'пшеница', SUGAR: 'сахар', RUONIA: 'ставка RUONIA', KEYRATE: 'ключевая ставка',
  BTC: 'биткоин', ETH: 'эфир', SOL: 'солана', IBIT: 'фонд биткоина IBIT', ETHA: 'фонд эфира ETHA', SBRF: 'Сбербанк', SBPR: 'Сбербанк (прив.)', GAZR: 'Газпром',
  LKOH: 'Лукойл', ROSN: 'Роснефть', GMKN: 'Норильский никель', VTBR: 'ВТБ', YDEX: 'Яндекс', T: 'Т-Технологии', MGNT: 'Магнит', TATN: 'Татнефть'
};
function refreshFutures() {
  var today = ymd_(new Date());
  var j = fetchJson_(ISS + '/engines/futures/markets/forts/securities.json?iss.meta=off&iss.only=securities,marketdata' +
    '&securities.columns=SECID,SHORTNAME,ASSETCODE,LASTTRADEDATE,PREVSETTLEPRICE,PREVPRICE,MINSTEP,STEPPRICE,INITIALMARGIN,LOTVOLUME,PREVOPENPOSITION,BUYSELLFEE' +
    '&marketdata.columns=SECID,LAST,SETTLEPRICE,OPENPOSITION,VALTODAY');
  if (!j) { log_('Фьючерсы: Мосбиржа не ответила, обновление пропущено'); return; }
  var group = {};
  Object.keys(FUT_GROUPS).forEach(function (g) { FUT_GROUPS[g].split(' ').forEach(function (c) { group[c] = g; }); });
  var md = {}, secs = rows_(j.securities);
  rows_(j.marketdata).forEach(function (r) { md[r.SECID] = r; });
  // ближайший контракт по каждому базовому активу
  var nearest = {};
  secs.forEach(function (s) {
    if (s.LASTTRADEDATE >= today && s.LASTTRADEDATE < '2090' && (!nearest[s.ASSETCODE] || s.LASTTRADEDATE < nearest[s.ASSETCODE])) nearest[s.ASSETCODE] = s.LASTTRADEDATE;
  });
  var live = secs.filter(function (s) { var m = md[s.SECID] || {}; return (m.OPENPOSITION || s.PREVOPENPOSITION) > 0; });
  var daily = candles_('futures/markets/forts', live.map(function (s) { return s.SECID; }), 24, addDays_(today, -430));
  var rows = [];
  secs.forEach(function (s) {
    var m = md[s.SECID] || {}, price = pos_(m.SETTLEPRICE, m.LAST, s.PREVSETTLEPRICE, s.PREVPRICE), im = pos_(s.INITIALMARGIN);
    if (!price || !im || !s.MINSTEP || !s.STEPPRICE) return;
    var perp = s.LASTTRADEDATE >= '2090', val = price / s.MINSTEP * s.STEPPRICE, d = daily[s.SECID] || [];
    var month = d.slice(-30), turn = 0;
    month.forEach(function (p) { turn += p[2]; });
    var row = {
      t: s.SECID, name: s.SHORTNAME, sector: group[s.ASSETCODE] || 'Акции', asset: s.ASSETCODE, aname: FUT_NAMES[s.ASSETCODE] || '',
      kind: perp ? 'вечный' : 'с датой исполнения', near: perp || nearest[s.ASSETCODE] === s.LASTTRADEDATE ? 'ближайший' : 'дальний',
      exp: perp ? null : s.LASTTRADEDATE, days: perp ? null : Math.round((parse_(s.LASTTRADEDATE) - parse_(today)) / 864e5), price: price,
      d1: pos_(m.SETTLEPRICE, m.LAST) && s.PREVSETTLEPRICE ? round_((price / s.PREVSETTLEPRICE - 1) * 100, 1) : null,
      val: round_(val, 0), im: round_(im, 0), lev: round_(val / im, 1), imp: round_(im / val * 100, 1),
      stepv: s.STEPPRICE, step: s.MINSTEP, lot: s.LOTVOLUME, oi: pos_(m.OPENPOSITION, s.PREVOPENPOSITION),
      turn: month.length && turn > 0 ? round_(turn / month.length / 1e6, 2) : (m.VALTODAY ? round_(m.VALTODAY / 1e6, 2) : null), fee: s.BUYSELLFEE || null
    };
    var ind = indicators_(d, null, price, today);
    ['w1', 'm1', 'm3', 'dhi', 'vol', 'rsi', 'ma50', 'hi', 'lo'].forEach(function (k) { row[k] = ind[k]; });
    rows.push(pick_(row, FUT_F));
  });
  if (rows.length < 50) { log_('Фьючерсы: собрано подозрительно мало (' + rows.length + '), прежние данные оставлены'); return; }
  var size = saveJson_('futures', { updated: today, fields: FUT_F, rows: rows });
  log_('Фьючерсы: контрактов ' + rows.length + ', с историей ' + live.length + ', объём ' + Math.round(size / 1024) + ' КБ');
}

// ---------- опционы ----------
// Отдаём только справочник базовых активов: даты исполнения, открытые позиции, позиции физических и юридических лиц.
// Саму доску опционов (страйки, цены) страница берёт у биржи напрямую, когда человек выбрал актив и дату.
var OPT_TYPES = { F: 'фьючерс', S: 'акция', G: 'валюта или металл', I: 'индекс' };
function refreshOptions() {
  var today = ymd_(new Date()), base = ISS + '/statistics/engines/futures/markets/options';
  var j = fetchJson_(base + '/assets.json?iss.meta=off');
  var list = rows_(j && j.asset_volumes);
  if (list.length < 10) { log_('Опционы: Мосбиржа не отдала список базовых активов, обновление пропущено'); return; }
  var codes = list.map(function (r) { return r.asset; });
  var exps = fetchAll_(codes.map(function (a) { return base + '/assets/' + encodeURIComponent(a) + '.json?iss.meta=off'; }));
  var poss = fetchAll_(codes.map(function (a) { return base + '/assets/' + encodeURIComponent(a) + '/openpositions.json?iss.meta=off'; }));
  var possP = fetchAll_(codes.map(function (a) { return base + '/assets/' + encodeURIComponent(a) + '/openpositions.json?iss.meta=off&option_type=P'; }));   // путы биржа отдаёт отдельным запросом
  var assets = [];
  list.forEach(function (r, i) {
    var e = rows_(exps[i] && exps[i].expirations).map(function (x) { return x.expiration_date; }).filter(function (d) { return d >= today; }).sort();
    if (!e.length) return;
    var pos = rows_(poss[i] && poss[i].open_positions).concat(rows_(possP[i] && possP[i].open_positions)), date = '';
    pos.forEach(function (p) { if (p.tradedate > date) date = p.tradedate; });
    assets.push({
      a: r.asset, name: FUT_NAMES[r.asset] || String(r.shortname || r.asset).replace(/\s+/g, ' ').replace(/ -?ао$/, '').trim(),
      type: OPT_TYPES[r.asset_type] || 'другое', price: r.asset_last_price || null, oi: r.openposition || 0, exps: e, posdate: date,
      // [колл или пут, физлица?, участников в покупке, участников в продаже, контрактов в покупке, контрактов в продаже]
      pos: pos.filter(function (p) { return p.tradedate === date; }).map(function (p) {
        return [p.option_type, p.is_fiz, p.persons_long || 0, p.persons_short || 0, p.open_position_long || 0, p.open_position_short || 0];
      })
    });
  });
  assets.sort(function (x, y) { return y.oi - x.oi; });
  if (assets.length < 10) { log_('Опционы: собрано слишком мало активов (' + assets.length + '), прежние данные оставлены'); return; }
  var size = saveJson_('options', { updated: today, assets: assets });
  log_('Опционы: базовых активов ' + assets.length + ', объём ' + Math.round(size / 1024) + ' КБ');
}
