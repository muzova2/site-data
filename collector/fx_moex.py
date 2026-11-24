# 24.11.2026: биржевые позиции курсов (юань, золото, серебро на бирже; биткоин и эфир — индексы МосБиржи)
# собираем на GitHub прямо с Мосбиржи. Копия логики refreshFx/fxItem_ из Google (Market.gs).
# Запуск: python fx_moex.py data/fx.json  — заменяет/добавляет эти позиции в файле. Если биржа не ответила, позиция остаётся прежней.
import json, math, sys, urllib.request
from datetime import date, datetime, timedelta, timezone

ISS = 'https://iss.moex.com/iss'
BASE = '2014-01-01'
TODAY = (datetime.now(timezone.utc) + timedelta(hours=3)).date().isoformat()
FX_MOEX = [['CNYRUB_TOM', 'Юань на бирже'], ['GLDRUB_TOM', 'Золото на бирже, за грамм'], ['SLVRUB_TOM', 'Серебро на бирже, за грамм']]
FX_CRYPTO = [['MOEXBTC', 'Биткоин, индекс МосБиржи'], ['MOEXETH', 'Эфир, индекс МосБиржи']]


def add(s, n): return (date.fromisoformat(s) + timedelta(days=n)).isoformat()
def rnd(v, d): k = 10 ** d; return math.floor(v * k + 0.5) / k          # как Math.round в JS


def get(url):
    for i in range(3):
        try:
            return json.load(urllib.request.urlopen(urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'}), timeout=60))
        except Exception as e:
            err = e
    raise err


def candles(engine, market, sec, interval, frm):
    out, start = [], 0
    while True:
        j = get(f'{ISS}/engines/{engine}/markets/{market}/securities/{sec}/candles.json?iss.meta=off&interval={interval}'
                f'&from={frm}&start={start}&candles.columns=close,begin')
        part = j['candles']['data']
        out += [[str(r[1])[:10], r[0]] for r in part if r[0]]
        start += 500
        if len(part) < 500 or start > 5000:
            return out


def series(engine, market, sec):
    daily = candles(engine, market, sec, 24, add(TODAY, -400))
    weekly = candles(engine, market, sec, 7, BASE)
    first = daily[0][0] if daily else '9999'
    return [p for p in weekly if p[0] < add(first, -6)] + daily


def thin(s):
    day_from, week_from, out, last = add(TODAY, -190), add(TODAY, -365 * 5 - 10), [], ''
    for d, v in s:
        if d >= day_from: key = d
        elif d >= week_from:
            dd = date.fromisoformat(d); key = 'w' + (dd - timedelta(days=dd.weekday())).isoformat()
        else: key = 'm' + d[:7]
        pt = [(date.fromisoformat(d) - date.fromisoformat(BASE)).days, v]
        if key == last: out[-1] = pt
        else: out.append(pt)
        last = key
    return out


def item(sid, name, group, unit, src, s):
    if len(s) < 2: return None
    ld, price = s[-1]
    def at(iso):
        if s[0][0] > iso: return None
        v = None
        for d, x in s:
            if d > iso: break
            v = x
        return v
    def chg(t): return rnd((price / t - 1) * 100, 1) if t else None
    o = {'id': sid, 'name': name, 'group': group, 'unit': unit, 'src': src, 'price': price, 'date': ld, 'd1': chg(s[-2][1])}
    for k, n in [('w1', 7), ('m1', 30), ('m3', 91), ('m6', 182), ('y1', 365), ('y3', 1095), ('y5', 1825)]:
        o[k] = chg(at(add(ld, -n)))
    o['ytd'] = chg(at(f'{int(ld[:4]) - 1}-12-31'))
    yr = [x for d, x in s if d >= add(ld, -365)]
    o['hi'], o['lo'], o['hist'] = max(yr), min(yr), thin(s)
    return o


def main(path):
    fx = json.load(open(path, encoding='utf-8'))
    items = fx['items']
    todo = [(c, 'Биржевые торги', '₽', 'торги Московской биржи', 'currency', 'selt') for c in FX_MOEX] + \
           [(c, 'Криптовалюта', '$', 'индекс Московской биржи, в долларах', 'stock', 'index') for c in FX_CRYPTO]
    for (sid, name), group, unit, src, eng, mkt in todo:
        try:
            it = item(sid, name, group, unit, src, series(eng, mkt, sid))
        except Exception as e:
            it = None; print(sid, 'ошибка:', e)
        if not it:
            print(sid, 'не получен, остаётся прежний' if any(x['id'] == sid for x in items) else 'не получен'); continue
        pos = next((i for i, x in enumerate(items) if x['id'] == sid), None)
        if pos is None: items.append(it)
        else: items[pos] = it
        print(sid, it['date'], it['price'])
    json.dump(fx, open(path, 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))


if __name__ == '__main__':
    main(sys.argv[1])
