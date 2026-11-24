
/**
 * ===== Скринер облигаций: данные для страницы bulkacapital.ru/skriner-obligaciy =====
 *
 * Раз в сутки забирает с Мосбиржи (ISS) все облигации основных режимов торгов: цены, НКД, доходность,
 * дюрацию, купоны, даты погашения и оферт, объёмы. Считает рыночный риск и свой показатель ликвидности.
 * Рейтинги биржа не даёт: они подставляются с листа «Рейтинги» (по ИНН эмитента), его ведут вручную.
 *
 * Первый запуск: setupBonds(). Справочные данные (эмитент, «только для квалов», история оборотов)
 * догружаются порциями: если за один запуск не успели, следующий продолжит.
 */
var BOND_BOARDS = ['TQOB', 'TQCB', 'TQOD', 'TQOY', 'TQRD'];
var SH_B = 'Облигации', SH_BREF = 'Облигации: справочник', SH_RATE = 'Рейтинги';
var B_HEAD = ['ISIN', 'Название', 'Эмитент', 'ИНН', 'Тип', 'Валюта', 'Цена, %', 'НКД', 'Номинал', 'Лот', 'Доходность, %',
  'Доходность к', 'Дата события', 'Срок, лет', 'Погашение', 'Оферта', 'Купон, %', 'Купон, сумма', 'Купонов в год',
  'След. купон', 'Тип купона', 'Амортизация', 'Дюрация, лет', 'Рыночный риск, %', 'Оборот в день, ₽', 'Ликвидность',
  'Объём выпуска, млн', 'Только квал', 'Дефолт', 'Повыш. риск', 'Рейтинг', 'Полное название'];
var B_KEYS = ['id', 'name', 'issuer', 'inn', 'type', 'cur', 'price', 'nkd', 'face', 'lot', 'y', 'yto', 'ev', 'years', 'mat',
  'offer', 'cp', 'cv', 'freq', 'next', 'ctype', 'amort', 'dur', 'risk', 'turn', 'liq', 'size', 'qual', 'def', 'pir', 'rating', 'full'];
var BREF_HEAD = ['ISIN', 'Эмитент', 'ИНН', 'Тип', 'Только квал', 'Дефолт', 'Купонов в год', 'Средний оборот, ₽', 'Доля дней со сделками', 'Справка загружена'];
var RATE_HEAD = ['ИНН', 'Эмитент', 'АКРА', 'Эксперт РА', 'НКР', 'НРА', 'Дата проверки', 'Комментарий'];
var RATE_AGENCIES = ['АКРА', 'Эксперт РА'];   // основные агентства: по ним считается оценка
var RATE_FALLBACK = ['НКР', 'НРА'];          // запасные: только если у основных рейтинга нет
var GRADES = ['D', 'C', 'CCC', 'B', 'B+', 'BB-', 'BB', 'BBB', 'A', 'AA', 'AAA'];   // номер в списке = балл от 0 до 10
var BOND_TYPES = { ofz_bond: 'ОФЗ', subfederal_bond: 'Региональная', municipal_bond: 'Муниципальная', corporate_bond: 'Корпоративная',
  exchange_bond: 'Корпоративная', euro_bond: 'Еврооблигация', ifi_bond: 'Международная организация', cb_bond: 'Банк России',
  state_bond: 'Государственная', non_exchange_bond: 'Корпоративная' };
var MINFIN_INN = '7710168360';   // Минфин России: только его бумаги считаем ОФЗ
var CARD_VER = 'v2';       // поменять, чтобы перечитать карточки всех бумаг
var LIQ_ALPHA = 0.15;      // вес нового дня в среднем обороте
var LIQ_SEED_DAYS = 16;    // на сколько дней назад заглядываем при первом запуске

function setupBonds() {
  sheet_(SH_B, B_HEAD); sheet_(SH_BREF, BREF_HEAD); sheet_(SH_RATE, RATE_HEAD); sheet_(SH_LOG, ['Время', 'Событие']);
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'refreshBonds') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('refreshBonds').timeBased().everyDays(1).atHour(19).inTimezone('Europe/Moscow').create();
  refreshBonds();
}

function refreshBonds() {
  var started = Date.now();
  var today = ymd_(new Date());
  var props = PropertiesService.getScriptProperties();

  // 1. Рыночные данные по всем режимам торгов
  var boards = fetchAll_(BOND_BOARDS.map(function (b) {
    return ISS + '/engines/stock/markets/bonds/boards/' + b + '/securities.json?iss.meta=off&iss.only=securities,marketdata,marketdata_yields';
  }));
  var bonds = {}, got = 0;
  boards.forEach(function (res, bi) {
    if (!res) return;
    got++;
    var md = {}, my = {};
    rows_(res.marketdata).forEach(function (r) { md[r.SECID] = r; });
    rows_(res.marketdata_yields).forEach(function (r) { my[r.SECID] = r; });
    rows_(res.securities).forEach(function (s) {
      var m = md[s.SECID] || {}, y = my[s.SECID] || {};
      var val = Number(m.VALTODAY_RUR || m.VALTODAY || 0);
      if (bonds[s.SECID] && bonds[s.SECID]._val >= val) return;
      bonds[s.SECID] = bondRow_(s, m, y, BOND_BOARDS[bi], today, val);
    });
  });
  if (!got) { log_('Облигации: Мосбиржа не ответила, обновление пропущено'); return; }
  var ids = Object.keys(bonds);

  // 2. Справочник: эмитент, тип, «только для квалов», история оборотов
  var ref = {};
  read_(SH_BREF).forEach(function (r) {
    ref[r[0]] = { issuer: r[1], inn: r[2], type: r[3], qual: r[4], def: r[5], freq: r[6], turn: r[7] === '' ? '' : Number(r[7]), share: r[8] === '' ? '' : Number(r[8]), done: r[9] };
  });
  var fresh = ids.filter(function (id) { return !ref[id]; });
  fresh.forEach(function (id) { ref[id] = { issuer: '', inn: '', type: '', qual: '', def: '', freq: '', turn: '', share: '', done: '' }; });

  // 2а. Эмитенты и типы: общий список бумаг (только когда появились новые)
  if (fresh.length) {
    for (var start = 0, more = true; more && start < 20000 && Date.now() - started < TIME_LIMIT_MS; start += 100 * BATCH) {
      var urls = [];
      for (var k = 0; k < BATCH; k++) urls.push(ISS + '/securities.json?iss.meta=off&engine=stock&market=bonds&is_trading=1&limit=100&start=' + (start + k * 100) +
        '&securities.columns=secid,emitent_title,emitent_inn,type');
      fetchAll_(urls).forEach(function (res) {
        var rs = rows_(res && res.securities);
        if (rs.length < 100) more = false;
        rs.forEach(function (r) {
          if (!ref[r.secid]) return;
          ref[r.secid].issuer = cleanIssuer_(r.emitent_title); ref[r.secid].inn = r.emitent_inn || ''; ref[r.secid].type = BOND_TYPES[r.type] || '';
        });
      });
    }
  }

  // 2б. Карточка бумаги: «только для квалов», дефолт, число купонов в год (порциями)
  var needCard = ids.filter(function (id) {   // новые бумаги, карточки старого формата и все с дефолтом (их перепроверяем каждый день)
    var r = ref[id]; return String(r.done).indexOf(CARD_VER) < 0 || (r.def !== 'нет' && String(r.done).indexOf(today) < 0);
  }), cards = 0;
  for (var a = 0; a < needCard.length && Date.now() - started < TIME_LIMIT_MS; a += BATCH) {
    var part = needCard.slice(a, a + BATCH);
    var res2 = fetchAll_(part.map(function (id) { return ISS + '/securities/' + id + '.json?iss.meta=off&iss.only=description&description.columns=name,value'; }));
    part.forEach(function (id, i) {
      if (!res2[i] || !res2[i].description) return;
      var dsc = {};
      res2[i].description.data.forEach(function (x) { dsc[x[0]] = x[1]; });
      var r = ref[id];
      r.qual = String(dsc.ISQUALIFIEDINVESTORS) === '1' ? 'да' : 'нет';
      r.def = String(dsc.HASDEFAULT) === '1' ? 'да' : (String(dsc.HASTECHNICALDEFAULT) === '1' ? 'тех' : 'нет');   // «тех» = технический дефолт (задержка выплаты)
      r.freq = dsc.COUPONFREQUENCY || '';
      if (!r.type) r.type = BOND_TYPES[dsc.TYPE] || '';
      r.done = today + ' ' + CARD_VER; cards++;
    });
    Utilities.sleep(200);
  }

  // 3. Ликвидность: средний дневной оборот по итогам прошедших торговых дней
  var liqLast = props.getProperty('LIQ_LAST') || addDays_(today, -LIQ_SEED_DAYS);
  var day = addDays_(liqLast, 1), liqDays = 0;
  while (day < today && Date.now() - started < TIME_LIMIT_MS) {
    var wd = parse_(day).getUTCDay();
    var turn = wd === 0 || wd === 6 ? null : dayTurnover_(day);
    if (turn === false) break;                       // биржа не ответила: продолжим в следующий раз
    if (turn && Object.keys(turn).length) {          // пусто = праздник
      ids.forEach(function (id) {
        var r = ref[id], v = turn[id] || 0, traded = v > 0 ? 1 : 0;
        if (r.turn === '' || isNaN(r.turn)) { r.turn = v; r.share = traded; r._n = 1; }
        else if (r._n) { r._n++; r.turn += (v - r.turn) / r._n; r.share += (traded - r.share) / r._n; }   // первые дни: простое среднее
        else { r.turn += LIQ_ALPHA * (v - r.turn); r.share += LIQ_ALPHA * (traded - r.share); }
      });
      liqDays++;
    }
    liqLast = day; day = addDays_(day, 1);
  }
  props.setProperty('LIQ_LAST', liqLast);

  // 4. Рейтинги: лист «Рейтинги» ведётся вручную (рейтинги агентств по ИНН эмитента), буква по нашей шкале считается здесь
  var rating = ratings_(ids, ref);

  // 5. Сборка строк и запись
  var out = ids.map(function (id) {
    var b = bonds[id], r = ref[id];
    b.issuer = r.issuer; b.inn = r.inn; b.type = r.type || b.type;
    if (b.type === 'ОФЗ' && String(r.inn) !== MINFIN_INN) b.type = 'Иностранное государство';   // бумаги Казахстана и Беларуси биржа относит к тому же виду, что и ОФЗ
    b.qual = r.qual; b.def = r.def; if (r.freq) b.freq = Number(r.freq);
    b.turn = r.turn === '' ? '' : Math.round(r.turn);
    b.liq = r.turn === '' ? '' : liqScore_(r.turn, r.share);
    b.rating = r.def === 'да' ? gradeText_(0, 'по выпуску зафиксирован дефолт', '')
      : (rating[String(r.inn)] || (b.type === 'ОФЗ' ? gradeText_(10, 'государственные облигации России', '') : ''));
    return B_KEYS.map(function (k) { return b[k]; });
  });
  out.sort(function (x, y) { return x[1] < y[1] ? -1 : x[1] > y[1] ? 1 : 0; });
  write_(SH_B, B_HEAD, out);
  write_(SH_BREF, BREF_HEAD, ids.map(function (id) {
    var r = ref[id];
    return [id, r.issuer, r.inn, r.type, r.qual, r.def, r.freq, r.turn === '' ? '' : Math.round(r.turn), r.share === '' ? '' : Math.round(r.share * 1000) / 1000, r.done];
  }));
  props.setProperty('BONDS_UPDATED', today);
  CacheService.getScriptCache().remove('bn');
  var left = needCard.length - cards;
  log_('Облигации: бумаг ' + ids.length + ', новых ' + fresh.length + ', карточек загружено ' + cards + (left > 0 ? ', осталось ' + left : '') +
    ', дней оборота учтено ' + liqDays);
  if (left > 0 || day < today) againBonds_();
}

function againBonds_() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'refreshBondsMore') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('refreshBondsMore').timeBased().after(60 * 1000).create();
}
function refreshBondsMore() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'refreshBondsMore') ScriptApp.deleteTrigger(t); });
  refreshBonds();
}

/** Одна облигация из трёх блоков ответа биржи. */
function bondRow_(s, m, y, board, today, val) {
  function dt(v) { v = String(v || ''); return v && v.slice(0, 4) !== '0000' ? v.slice(0, 10) : ''; }
  var cur = s.FACEUNIT === 'SUR' || s.FACEUNIT === 'RUR' ? 'RUB' : (s.FACEUNIT || 'RUB');
  var mat = dt(s.MATDATE), put = dt(s.PUTOPTIONDATE) || dt(s.OFFERDATE), call = dt(s.CALLOPTIONDATE);
  var offer = put && put > today ? put : '';
  var yDate = dt(y.YIELDDATE), yType = String(y.YIELDDATETYPE || '');
  var ev = offer || (yType && yType !== 'MATDATE' && yDate > today ? yDate : '') || mat;
  var yto = !ev ? '' : (ev === mat ? 'погашению' : (offer ? 'оферте' : (call && ev === call ? 'колл-опциону' : 'оферте')));
  var years = ev ? Math.round((parse_(ev) - parse_(today)) / 864e5 / 365.25 * 100) / 100 : '';
  var price = pos_(m.LAST, m.MARKETPRICE, m.LCURRENTPRICE, s.PREVLEGALCLOSEPRICE, s.PREVPRICE, s.PREVWAPRICE);
  var yld = num_(y.EFFECTIVEYIELD, m.YIELD, s.YIELDATPREVWAPRICE);
  if (yld !== '' && (yld <= -50 || yld > 300)) yld = '';                        // биржевой мусор по неликвидным бумагам
  var durDays = pos_(y.DURATION, m.DURATION);
  var dur = durDays ? Math.round(durDays / 365 * 100) / 100 : '';
  var shock = cur === 'RUB' ? 2 : 1;                                             // рост ставок: 2 п.п. для рублей, 1 п.п. для валют
  var risk = dur !== '' ? -Math.round(dur / (1 + (yld === '' ? 0 : yld) / 100) * shock * 10) / 10 : '';
  var period = Number(s.COUPONPERIOD) || 0;
  return {
    id: s.SECID, name: s.SHORTNAME, full: s.SECNAME || '', issuer: '', inn: '', cur: cur, _val: val,
    type: s.SECTYPE === '3' ? 'ОФЗ' : (s.SECTYPE === '4' ? 'Региональная' : (s.SECTYPE === 'C' ? 'Муниципальная' : 'Корпоративная')),
    price: price || '', nkd: s.ACCRUEDINT === null ? '' : s.ACCRUEDINT, face: s.FACEVALUE || '', lot: s.LOTSIZE || 1,
    y: yld, yto: yld === '' ? '' : yto, ev: ev, years: years, mat: mat, offer: offer,
    cp: s.COUPONPERCENT === null ? '' : s.COUPONPERCENT, cv: s.COUPONVALUE || '', freq: period ? Math.round(365 / period) : '',
    next: dt(s.NEXTCOUPON), ctype: /ОФЗ-ИН/.test(String(s.SECNAME)) ? 'Привязан к инфляции' : couponKind_(s.BONDTYPE, s.COUPON_DETAILS),
    amort: /мортиз/.test(String(s.BONDTYPE)) || (s.FACEVALUE_TYPE && s.FACEVALUE_TYPE !== 'Фиксированный') ? 'да' : 'нет',
    dur: dur, risk: risk, turn: '', liq: '', size: s.ISSUESIZE && s.FACEVALUE ? Math.round(s.ISSUESIZE * s.FACEVALUE / 1e4) / 100 : '',
    qual: '', def: '', pir: board === 'TQRD' ? 'да' : 'нет', rating: ''
  };
}

/** Обороты всех облигаций за один прошедший день: {ISIN: оборот в рублях}. false = биржа не ответила. */
function dayTurnover_(day) {
  var base = function (b, start) {
    return ISS + '/history/engines/stock/markets/bonds/boards/' + b + '/securities.json?iss.meta=off&date=' + day +
      '&history.columns=SECID,VALUE,NUMTRADES&limit=100&start=' + start;
  };
  var first = fetchAll_(BOND_BOARDS.map(function (b) { return base(b, 0); }));
  if (first.some(function (r) { return !r; })) return false;
  var turn = {}, rest = [];
  function take(res) { rows_(res && res.history).forEach(function (r) { if (r.VALUE > 0) turn[r.SECID] = (turn[r.SECID] || 0) + r.VALUE; }); }
  first.forEach(function (res, i) {
    take(res);
    var c = rows_(res['history.cursor'])[0], total = c ? Number(c.TOTAL) : 0;
    for (var st = 100; st < total; st += 100) rest.push(base(BOND_BOARDS[i], st));
  });
  var more = fetchAll_(rest);
  if (more.some(function (r) { return !r; })) return false;
  more.forEach(take);
  return turn;
}

/**
 * Показатель ликвидности от 0 до 100.
 * Оборот 10 тыс. ₽ в день и меньше = 0, 100 млн ₽ в день и больше = 100 (шкала логарифмическая),
 * результат умножается на долю дней, когда по бумаге были сделки.
 */
function liqScore_(turn, share) {
  if (!(turn > 0)) return 0;
  var s = (Math.log(turn) / Math.LN10 - 4) / 4 * 100;
  return Math.round(Math.max(0, Math.min(100, s)) * Math.max(0, Math.min(1, share)));
}

/** Тип купона простыми словами: из двух биржевых полей (вид облигации и описание купона). */
function couponKind_(btype, details) {
  var b = String(btype || ''), c = String(details || '');
  if (/труктурн/.test(b)) return 'Структурная';
  if (/исконтн/.test(b)) return 'Без купона';
  if (/плавающ/.test(b)) return c === 'ИПЦ' ? 'Привязан к инфляции' : 'Плавающий' + (c ? ' (' + c + ')' : '');
  if (/мортиз/.test(b)) return 'С амортизацией';
  if (/неизвестн/.test(b) || /неизвестн/.test(c)) return 'Переменный';
  if (/известн/.test(b) || /известн/.test(c) || /алютн/.test(b)) return 'Фиксированный';
  return c || b || '';
}

/** Рейтинг агентства (AA+(RU), ruAA+, AA+.ru, AA+|ru|) в балл нашей шкалы от 0 до 10. null = не распознан. */
function agencyScore_(v) {
  var x = String(v || '').toUpperCase().replace(/[−–—]/g, '-').replace(/\(RU\)|\.RU|RU|\|/g, '').replace(/\s+/g, '');
  var m = x.match(/^(AAA|AA|A|BBB|BB|B|CCC|CC|C|RD|SD|D)([+-])?/);
  if (!m) return null;
  var l = m[1], sign = m[2] || '';
  if (l === 'AAA') return 10;
  if (l === 'AA') return 9;
  if (l === 'A') return 8;
  if (l === 'BBB') return 7;
  if (l === 'BB') return sign === '-' ? 5 : 6;
  if (l === 'B') return sign === '+' ? 4 : 3;
  if (l === 'CCC') return 2;
  if (l === 'CC' || l === 'C') return 1;
  return 0;
}
/** Что уходит на страницу: буква | балл | на чём основано | дата проверки. */
function gradeText_(score, source, date) { return GRADES[score] + '|' + score + '|' + source + '|' + (date || ''); }

/**
 * Читает лист «Рейтинги» и возвращает {ИНН: оценка}. Если рейтингов несколько, берётся среднее с округлением вниз.
 * Заодно дописывает в лист эмитентов, которых там ещё нет (по убыванию оборота), чтобы было что заполнять.
 */
function ratings_(ids, ref) {
  var sh = sheet_(SH_RATE, RATE_HEAD);
  sh.getRange(1, 1, 1, RATE_HEAD.length).setValues([RATE_HEAD]).setFontWeight('bold');
  var rows = read_(SH_RATE), have = {}, out = {};
  var dateCol = 2 + RATE_AGENCIES.length + RATE_FALLBACK.length;
  rows.forEach(function (r) {
    var inn = String(r[0]).trim(); have[inn] = 1;
    function avg(names, from) {
      var parts = [], sum = 0;
      names.forEach(function (a, i) {
        var v = String(r[from + i] || '').trim(), sc = v ? agencyScore_(v) : null;
        if (sc !== null) { parts.push(a + ' ' + v.replace(/\|/g, '∣')); sum += sc; }   // черта в обозначении НРА мешает разделителю полей
      });
      return parts.length ? gradeText_(Math.floor(sum / parts.length), parts.join(', '), day_(r[dateCol])) : '';
    }
    // запасные агентства учитываются, только если нет рейтинга ни от одного основного
    var g = avg(RATE_AGENCIES, 2) || avg(RATE_FALLBACK, 2 + RATE_AGENCIES.length);
    if (g) out[inn] = g;
  });
  var turn = {}, name = {};
  ids.forEach(function (id) {
    var r = ref[id], inn = String(r.inn || '').trim();
    if (!inn) return;
    turn[inn] = (turn[inn] || 0) + (Number(r.turn) || 0); name[inn] = r.issuer;
  });
  var missing = Object.keys(turn).filter(function (inn) { return !have[inn]; }).sort(function (a, b) { return turn[b] - turn[a]; });
  if (missing.length) {
    var all = rows.map(function (r) { return RATE_HEAD.map(function (h, i) { var v = r[i]; return v instanceof Date ? ymd_(v) : (v === undefined ? '' : v); }); })
      .concat(missing.map(function (inn) { return [inn, name[inn], '', '', '', '', '', '']; }));
    write_(SH_RATE, RATE_HEAD, all);
  }
  return out;
}

function cleanIssuer_(s) {
  return String(s || '').replace(/\\"/g, '"').replace(/\s+/g, ' ').trim();
}
function num_() {
  for (var i = 0; i < arguments.length; i++) { var v = arguments[i]; if (typeof v === 'number' && v !== 0) return v; }
  return '';
}

/** JSON для страницы: {updated, fields:[...], rows:[[...]]} */
function bondsJson_() {
  var cache = CacheService.getScriptCache();
  var n = Number(cache.get('bn') || 0), parts = [];
  if (n) {
    var keys = []; for (var i = 0; i < n; i++) keys.push('b' + i);
    var got = cache.getAll(keys);
    for (var k = 0; k < n; k++) parts.push(got['b' + k]);
    if (parts.every(function (x) { return typeof x === 'string'; })) return parts.join('');
  }
  var numeric = { price: 1, nkd: 1, face: 1, lot: 1, y: 1, years: 1, cp: 1, cv: 1, freq: 1, dur: 1, risk: 1, turn: 1, liq: 1, size: 1 };
  var rows = read_(SH_B).map(function (r) {
    return B_KEYS.map(function (key, i) {
      var v = r[i];
      if (v instanceof Date) v = ymd_(v);
      if (v === '') return null;
      return numeric[key] ? Number(v) : String(v);
    });
  });
  var s = JSON.stringify({ updated: PropertiesService.getScriptProperties().getProperty('BONDS_UPDATED') || '', fields: B_KEYS, rows: rows });
  var put = {}, size = 90000, cnt = Math.ceil(s.length / size);
  for (var c = 0; c < cnt; c++) put['b' + c] = s.substr(c * size, size);
  put.bn = String(cnt);
  try { cache.putAll(put, 3600); } catch (e) {}
  return s;
}
