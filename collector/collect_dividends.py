# -*- coding: utf-8 -*-
"""Собирает дивиденды российских акций из API Т-Инвестиций и сохраняет в dividends_raw.json.

Зачем локально: сервер Т-Инвестиций не принимает соединения с серверов Google, поэтому
скрипт в Google-таблице сам забрать дивиденды не может. Запускать при выключенном VPN.

Токен («только для чтения») лежит в файле  %USERPROFILE%\\.bulka\\tinvest_token.txt
(одна строка, без кавычек). Файл вне OneDrive, в облако не попадает. Скрипт токен не печатает.

Запуск:  python collect_dividends.py            собрать и отправить в Google-таблицу
         python collect_dividends.py --upload   отправить уже собранный dividends_raw.json
         ... --log                              писать ход работы в collect.log (так запускает планировщик Windows,
                                                задание «Bulka Capital - дивиденды» запускает через run_weekly.py: пт 18:00, повторы сб/вс 12:00 и 19:00)
"""
import datetime
import hashlib
import json
import os
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

HERE = Path(__file__).parent
TOKEN_FILE = Path(os.environ.get("USERPROFILE", str(Path.home()))) / ".bulka" / "tinvest_token.txt"
API = "https://invest-public-api.tbank.ru/rest/tinkoff.public.invest.api.contract.v1.InstrumentsService/"
YEARS_BACK = 13
PAUSE = 0.4   # лимит сервиса около 200 запросов в минуту
# веб-приложение скрипта «Календарь дивидендов» (Google Apps Script), принимает дивиденды в таблицу
SHEET_APP = "https://script.google.com/macros/s/AKfycbyQammVyKFaW1iym9aCJV7lxRWvePm8T4xx5wkstlF56SHBSVNlfsji4BX7AXXq_YAG/exec"
# 15.11.2026: на GitHub (задание weekly.yml) токен приходит из секрета TINVEST_TOKEN, а сертификат Т-Банка
# проверяется по корневому сертификату Минцифры из файла TBANK_CA. На её компьютере — как раньше (файл токена, Windows).
_CTX = None
if os.environ.get("TBANK_CA"):
    import ssl
    _CTX = ssl.create_default_context()
    _CTX.load_verify_locations(os.environ["TBANK_CA"])


def get_token():
    tok = os.environ.get("TINVEST_TOKEN", "").strip()
    if tok:
        return tok
    if not TOKEN_FILE.exists():
        raise SystemExit(f"Нет файла с токеном: {TOKEN_FILE}")
    return TOKEN_FILE.read_text(encoding="utf-8").strip()


def call(method, body, token, tries=4):
    req = urllib.request.Request(API + method, data=json.dumps(body).encode("utf-8"),
                                 headers={"Content-Type": "application/json", "Authorization": "Bearer " + token})
    for attempt in range(tries):
        try:
            with urllib.request.urlopen(req, timeout=60, context=_CTX) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code == 429 and attempt < tries - 1:      # превышен лимит: ждём и повторяем
                time.sleep(15)
                continue
            raise SystemExit(f"Т-Инвестиции ответили ошибкой {e.code} на {method}: {e.read().decode('utf-8', 'replace')[:200]}")
        except urllib.error.URLError as e:
            if attempt < tries - 1:
                time.sleep(5)
                continue
            raise SystemExit(f"Нет связи с Т-Инвестициями ({e.reason}). Проверьте, что VPN выключен.")


def money(m):
    return int(m.get("units", 0) or 0) + int(m.get("nano", 0) or 0) / 1e9


def upload(out, token):
    """Отправляет собранное в Google-таблицу. Вместо токена уходит его SHA-256: скрипт сверяет со своим."""
    body = dict(out, sig=hashlib.sha256(token.encode("utf-8")).hexdigest())
    req = urllib.request.Request(SHEET_APP, data=json.dumps(body, ensure_ascii=False).encode("utf-8"),
                                 headers={"Content-Type": "text/plain;charset=utf-8"})
    try:
        with urllib.request.urlopen(req, timeout=180) as r:
            text = r.read().decode("utf-8", "replace")
    except urllib.error.URLError as e:
        raise SystemExit(f"Не удалось отправить в Google-таблицу: {e}")
    try:
        res = json.loads(text)
    except ValueError:
        raise SystemExit("Google-таблица ответила не по формату: " + text[:200])
    if not res.get("ok"):
        raise SystemExit("Google-таблица отклонила данные: " + str(res.get("error")))
    print(f"в Google-таблицу отправлено выплат: {res.get('items')}")


def main():
    token = get_token()
    if not token:
        raise SystemExit("Файл с токеном пустой.")
    if "--upload" in sys.argv:
        upload(json.loads((HERE / "dividends_raw.json").read_text(encoding="utf-8")), token)
        return
    today = datetime.date.today()
    start = f"{today.year - YEARS_BACK + 1}-01-01T00:00:00Z"
    end = f"{today.year + 2}-01-01T00:00:00Z"

    shares = call("Shares", {"instrumentStatus": "INSTRUMENT_STATUS_BASE"}, token).get("instruments", [])
    tqbr = sorted((s for s in shares if s.get("classCode") == "TQBR"), key=lambda s: s["ticker"])
    print(f"акций в основном режиме торгов: {len(tqbr)}")

    items, companies = [], []
    for n, s in enumerate(tqbr, 1):
        res = call("GetDividends", {"instrumentId": s["uid"], "from": start, "to": end}, token)
        rows = []
        for d in res.get("dividends", []):
            net = d.get("dividendNet") or {}
            if (net.get("currency") or "rub").lower() != "rub":
                continue
            reg = (d.get("recordDate") or "")[:10]
            val = round(money(net), 6)
            if not reg or val <= 0:
                continue
            rows.append([s["ticker"], reg, val, (d.get("lastBuyDate") or "")[:10], (d.get("declaredDate") or "")[:10],
                         d.get("dividendType") or ""])
        if rows:
            companies.append([s["ticker"], s.get("name", ""), s.get("sector", "")])
            items.extend(rows)
        if n % 25 == 0:
            print(f"  обработано {n} из {len(tqbr)}")
        time.sleep(PAUSE)

    out = {"updated": today.isoformat(), "companies": companies, "items": items}
    (HERE / "dividends_raw.json").write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
    future = sorted((r for r in items if r[1] >= today.isoformat()), key=lambda r: r[1])
    years = sorted({r[1][:4] for r in items})
    print(f"компаний с дивидендами: {len(companies)}, выплат: {len(items)}, годы: {years[0]}–{years[-1]}")
    print(f"будущих выплат: {len(future)}")
    for r in future[:15]:
        print(f"   {r[1]}  {r[0]:6}  {r[2]} ₽  (последний день покупки {r[3] or '?'})")
    upload(out, token)


if __name__ == "__main__":
    if "--log" in sys.argv:
        log = open(HERE / "collect.log", "a", encoding="utf-8", buffering=1)
        sys.stdout = sys.stderr = log
        print(f"\n=== {datetime.datetime.now():%d.%m.%Y %H:%M} ===")
        try:
            main()
            print("готово")
            # тем же еженедельным запуском обновляем финансовые показатели акций для страницы «Анализ акций»
            if "--upload" not in sys.argv:
                try:
                    import collect_fundamentals
                    print("--- показатели акций ---")
                    collect_fundamentals.main()
                    print("готово")
                except BaseException as e:
                    print("ОШИБКА в показателях акций:", repr(e))
        except SystemExit as e:
            print("ОШИБКА:", e)
            raise
        except Exception as e:
            print("ОШИБКА:", repr(e))
            raise
    else:
        sys.stdout.reconfigure(encoding="utf-8")
        main()
