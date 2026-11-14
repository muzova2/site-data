# -*- coding: utf-8 -*-
"""Собирает финансовые показатели российских акций из API Т-Инвестиций в fundamentals_raw.json.

Нужен для страницы «Анализ акций». Токен и правила те же, что у collect_dividends.py:
файл %USERPROFILE%\\.bulka\\tinvest_token.txt, VPN выключен, токен не печатается.

Запуск:  python collect_fundamentals.py            собрать и отправить в Google-таблицу
         python collect_fundamentals.py --upload   отправить уже собранный fundamentals_raw.json
         ... --log                                  писать ход работы в collect.log (так запускает планировщик)
"""
import datetime
import hashlib
import json
import sys
import urllib.request

import collect_dividends as c

KEYS = ["marketCapitalization", "highPriceLast52Weeks", "lowPriceLast52Weeks", "averageDailyVolumeLast4Weeks", "beta",
        "freeFloat", "sharesOutstanding", "revenueTtm", "ebitdaTtm", "netIncomeTtm", "epsTtm", "freeCashFlowTtm",
        "peRatioTtm", "priceToSalesTtm", "priceToBookTtm", "priceToFreeCashFlowTtm", "totalEnterpriseValueMrq",
        "evToEbitdaMrq", "evToSales", "netMarginMrq", "roe", "roa", "roic", "totalDebtMrq", "totalDebtToEquityMrq",
        "netDebtToEbitda", "dividendYieldDailyTtm", "dividendRateTtm", "fiveYearsAverageDividendYield",
        "dividendPayoutRatioFy", "oneYearAnnualRevenueGrowthRate", "threeYearAnnualRevenueGrowthRate",
        "revenueChangeFiveYears", "epsChangeFiveYears", "exDividendDate"]


def upload(out, token):
    """Отправляет показатели в Google-таблицу. Вместо токена уходит его SHA-256: скрипт сверяет со своим."""
    body = {"kind": "fund", "sig": hashlib.sha256(token.encode("utf-8")).hexdigest(), "updated": out["updated"], "items": out["items"]}
    req = urllib.request.Request(c.SHEET_APP, data=json.dumps(body, ensure_ascii=False).encode("utf-8"),
                                 headers={"Content-Type": "text/plain;charset=utf-8"})
    with urllib.request.urlopen(req, timeout=180) as r:
        text = r.read().decode("utf-8", "replace")
    try:
        res = json.loads(text)
    except ValueError:
        raise SystemExit("Google-таблица ответила не по формату: " + text[:200])
    if not res.get("ok"):
        raise SystemExit("Google-таблица отклонила показатели: " + str(res.get("error")))
    print(f"в Google-таблицу отправлено показателей акций: {res.get('items')}")


def main():
    token = c.get_token()
    if "--upload" in sys.argv:
        upload(json.loads((c.HERE / "fundamentals_raw.json").read_text(encoding="utf-8")), token)
        return
    shares = c.call("Shares", {"instrumentStatus": "INSTRUMENT_STATUS_BASE"}, token).get("instruments", [])
    tq = sorted((s for s in shares if s.get("classCode") == "TQBR" and s.get("assetUid")), key=lambda s: s["ticker"])
    by_asset = {}
    for s in tq:
        by_asset.setdefault(s["assetUid"], []).append(s)
    uids = list(by_asset)
    items = []
    for i in range(0, len(uids), 100):
        res = c.call("GetAssetFundamentals", {"assets": uids[i:i + 100]}, token)
        for f in res.get("fundamentals", []):
            for s in by_asset.get(f.get("assetUid"), []):
                row = {"t": s["ticker"], "name": s.get("name", ""), "sector": s.get("sector", ""), "isin": s.get("isin", "")}
                row.update({k: f.get(k) for k in KEYS})
                items.append(row)
    items.sort(key=lambda r: r["t"])
    out = {"updated": datetime.date.today().isoformat(), "items": items}
    (c.HERE / "fundamentals_raw.json").write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
    print(f"акций в основном режиме торгов: {len(tq)}, с показателями: {len(items)}")
    upload(out, token)


if __name__ == "__main__":
    if "--log" in sys.argv:
        log = open(c.HERE / "collect.log", "a", encoding="utf-8", buffering=1)
        sys.stdout = sys.stderr = log
        print(f"\n=== {datetime.datetime.now():%d.%m.%Y %H:%M} показатели акций ===")
        try:
            main()
            print("готово")
        except BaseException as e:
            print("ОШИБКА:", repr(e))
            raise
    else:
        sys.stdout.reconfigure(encoding="utf-8")
        main()
