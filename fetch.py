"""Собирает открытые данные для «Деньги Monee» и сохраняет их в data.json.

Что берёт (всё бесплатно и без ключей, личных данных здесь нет):
  1. Курсы ЦБ РФ — юань, доллар, евро (https://www.cbr-xml-daily.ru).
  2. Производственный календарь РФ — рабочие и нерабочие дни (https://isdayoff.ru).

Как ведёт себя при сбоях:
  - Каждый источник берётся отдельно. Если один не ответил, его часть
    остаётся из старого data.json (с пометкой, когда она была свежей),
    а второй всё равно обновляется.
  - Если не ответил ни один — data.json не трогаем вовсе.
  - Скрипт всегда завершается без ошибки, чтобы приложение не ломалось.

Запуск: python fetch.py   (нужен только Python 3.9+, без доп. библиотек)
"""

import json
import os
import sys
import urllib.request
from datetime import date, datetime, timedelta, timezone

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data.json")

UFA = timezone(timedelta(hours=5))  # Уфа, UTC+5, без перехода на летнее время

RATES_URL = "https://www.cbr-xml-daily.ru/daily_json.js"
CAL_URL = "https://isdayoff.ru/api/getdata?year={year}&pre=1"
CURRENCIES = [("CNY", "Юань"), ("USD", "Доллар"), ("EUR", "Евро")]

# Названия праздников РФ (isdayoff отдаёт только «рабочий / нерабочий»).
HOLIDAY_NAMES = {
    (1, 1): "Новогодние каникулы", (1, 2): "Новогодние каникулы", (1, 3): "Новогодние каникулы",
    (1, 4): "Новогодние каникулы", (1, 5): "Новогодние каникулы", (1, 6): "Новогодние каникулы",
    (1, 7): "Рождество", (1, 8): "Новогодние каникулы",
    (2, 23): "День защитника Отечества", (3, 8): "Международный женский день",
    (5, 1): "Праздник Весны и Труда", (5, 9): "День Победы",
    (6, 12): "День России", (11, 4): "День народного единства",
}


def log(msg):
    print(msg, flush=True)


def get(url, timeout=20):
    req = urllib.request.Request(url, headers={"User-Agent": "monee-dengi-data/1.0"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read().decode("utf-8")


def now_ufa():
    return datetime.now(UFA)


def iso(dt):
    return dt.replace(microsecond=0).isoformat()


# ---------- 1. Курсы ЦБ ----------

def fetch_rates():
    d = json.loads(get(RATES_URL))
    items = []
    for code, title in CURRENCIES:
        v = d["Valute"][code]
        nominal = int(v["Nominal"])
        value = round(float(v["Value"]) / nominal, 4)        # за 1 единицу
        prev = round(float(v["Previous"]) / nominal, 4)
        if not (0 < value < 100000 and 0 < prev < 100000):
            raise ValueError(f"странный курс {code}: {value}")
        items.append({
            "code": code,
            "name": title,
            "value": value,
            "prev": prev,
            "change": round(value - prev, 4),
        })
    return {
        "ok": True,
        "source": "ЦБ РФ (cbr-xml-daily.ru)",
        "date": d["Date"][:10],             # с какого дня действует курс
        "prev_date": d["PreviousDate"][:10],
        "items": items,
        "fetched_at": iso(now_ufa()),
    }


# ---------- 2. Производственный календарь ----------

def year_days(year):
    """Строка по одному символу на день года: 0 — рабочий, 1 — выходной,
    2 — сокращённый рабочий (предпраздничный)."""
    s = get(CAL_URL.format(year=year)).strip()
    days_in_year = (date(year + 1, 1, 1) - date(year, 1, 1)).days
    if len(s) != days_in_year or set(s) - set("0124"):
        raise ValueError(f"календарь {year}: неожиданный ответ")
    return s


def fetch_calendar():
    today = now_ufa().date()
    days = {}
    for y in (today.year, today.year + 1):
        try:
            s = year_days(y)
        except Exception as e:  # следующий год могут ещё не утвердить
            if y == today.year:
                raise
            log(f"  календарь на {y} пока нет: {e}")
            continue
        start = date(y, 1, 1)
        for i, ch in enumerate(s):
            days[start + timedelta(days=i)] = ch

    def is_work(d):
        return days.get(d) in ("0", "2")

    first = today.replace(day=1)
    nxt = (first + timedelta(days=32)).replace(day=1)
    month = [first + timedelta(days=i) for i in range((nxt - first).days)]
    work_total = sum(is_work(d) for d in month)
    work_left = sum(is_work(d) for d in month if d >= today)

    # Ближайшие праздники: нерабочие дни, выпавшие на будни (пн–пт),
    # склеиваем подряд идущие нерабочие дни в одни «каникулы».
    holidays = []
    d = today
    horizon = today + timedelta(days=180)
    while d <= horizon and len(holidays) < 3:
        if d in days and days[d] == "1" and d.weekday() < 5:
            s = d
            while s - timedelta(days=1) in days and days[s - timedelta(days=1)] == "1":
                s -= timedelta(days=1)
            e = d
            while e + timedelta(days=1) in days and days[e + timedelta(days=1)] == "1":
                e += timedelta(days=1)
            if s < today:
                s = today
            name = next((HOLIDAY_NAMES[(x.month, x.day)]
                         for x in (s + timedelta(days=i) for i in range((e - s).days + 1))
                         if (x.month, x.day) in HOLIDAY_NAMES), "Перенос выходного")
            holidays.append({"from": s.isoformat(), "to": e.isoformat(),
                             "days": (e - s).days + 1, "name": name})
            d = e + timedelta(days=1)
            continue
        d += timedelta(days=1)

    tomorrow = today + timedelta(days=1)
    return {
        "ok": True,
        "source": "Производственный календарь РФ (isdayoff.ru)",
        "today": today.isoformat(),
        "today_is_workday": is_work(today),
        "today_is_short": days.get(today) == "2",
        "tomorrow_is_workday": is_work(tomorrow) if tomorrow in days else None,
        "month": first.strftime("%Y-%m"),
        "workdays_in_month": work_total,
        "workdays_left": work_left,
        "next_holidays": holidays,
        "fetched_at": iso(now_ufa()),
    }


# ---------- сборка ----------

def main():
    old = {}
    try:
        with open(DATA, encoding="utf-8") as f:
            old = json.load(f)
    except Exception:
        pass

    out = {"app": "Деньги Monee", "timezone": "Уфа (UTC+5)"}
    fresh = 0
    for key, fn in (("rates", fetch_rates), ("calendar", fetch_calendar)):
        try:
            out[key] = fn()
            fresh += 1
            log(f"✅ {key}: обновлено")
        except Exception as e:
            log(f"⚠ {key}: источник не ответил ({e}). Оставляю прошлые данные.")
            prev = old.get(key)
            if prev:
                prev = dict(prev)
                prev["ok"] = False          # приложение покажет «данные от …»
                out[key] = prev
            else:
                out[key] = {"ok": False}

    if fresh == 0:
        log("❌ Ни один источник не ответил — data.json не меняю.")
        return 0

    out["updated_at"] = iso(now_ufa())
    tmp = DATA + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
        f.write("\n")
    os.replace(tmp, DATA)
    log(f"💾 data.json сохранён, обновлено в {out['updated_at']}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:  # что бы ни случилось — старый файл цел, задание не падает
        log(f"❌ Непредвиденная ошибка: {e}. data.json не тронут.")
        sys.exit(0)
