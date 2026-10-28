"""
Полная проверка регистра и дашборда: не ищет расхождений с чужой книгой, а
смотрит, всё ли внутри нашей таблицы логично и согласовано.

Что проверяет:
  * нет строк с пустым каналом продаж (J) — такие в дашборд не попадают;
  * нет оптовых клиентов (их правило отбора);
  * нет дублей документов внутри дня;
  * даты идут без пропусков до вчерашнего дня;
  * суммы дашборда «Sales by day» сходятся с суммами регистра.

    python tools/audit_register.py
    python tools/audit_register.py --month 9
"""
import argparse
import sys
from collections import Counter, defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(errors="replace")

from src.clock import today as tz_today, yesterday as tz_yesterday
from src.config import load_config
from src.sheets_register import from_serial, get_client, open_register

OK, BAD, WARN = "  [ок]", "  [!!]", "  [~]"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--month", type=int, default=None)
    args = parser.parse_args()
    month = args.month or tz_today().month

    cfg = load_config()
    client = get_client(cfg.credentials_path)
    ss, ws = open_register(client, cfg.spreadsheet_id, cfg.register_sheet)
    print(f"Книга «{ss.title}», лист «{ws.title}»\n")

    rows = ws.get(f"A{cfg.register_header_row + 1}:V{ws.row_count}",
                  value_render_option="UNFORMATTED_VALUE")
    per_day = defaultdict(lambda: {"rows": 0, "revenue": 0.0, "profit": 0.0})
    channels = Counter()
    groups = Counter()
    empty_channel = []
    excluded = []
    docs = defaultdict(Counter)
    problems = 0

    wanted = {g.strip().lower() for g in cfg.exclude_client_groups}
    for offset, raw in enumerate(rows):
        raw = (list(raw) + [""] * 22)[:22]
        day = from_serial(raw[8]) if raw[8] != "" else None
        if day is None:
            continue
        row_index = cfg.register_header_row + 1 + offset
        if day.month != month:
            continue
        stats = per_day[day]
        stats["rows"] += 1
        if isinstance(raw[18], (int, float)):
            stats["revenue"] += raw[18]
        if isinstance(raw[20], (int, float)):
            stats["profit"] += raw[20]

        channel = str(raw[9]).strip()
        channels[channel or "(пусто)"] += 1
        if not channel:
            empty_channel.append(row_index)
        group = str(raw[11]).strip().lower()
        groups[group or "(пусто)"] += 1
        if group in wanted:
            excluded.append((row_index, group))
        key = (str(raw[0]).strip(), str(raw[4]).strip(), str(raw[13]).strip()[:30],
               raw[15] if isinstance(raw[15], (int, float)) else "", raw[18])
        docs[day][key] += 1

    print(f"Строк за месяц {month:02d}: {sum(s['rows'] for s in per_day.values())}")
    print(f"Каналы продаж: {dict(channels.most_common())}")
    print(f"Группы доступа: {dict(groups.most_common())}\n")

    if empty_channel:
        problems += 1
        print(f"{BAD} строк с пустым каналом (J): {len(empty_channel)}, например {empty_channel[:5]}")
    else:
        print(f"{OK} строк с пустым каналом нет")

    if excluded:
        problems += 1
        print(f"{BAD} оптовых строк в регистре: {len(excluded)}, например {excluded[:5]}")
    else:
        print(f"{OK} оптовых клиентов ({', '.join(cfg.exclude_client_groups)}) нет")

    dupes = {day: [k for k, n in counter.items() if n > 1] for day, counter in docs.items()}
    dupes = {day: keys for day, keys in dupes.items() if keys}
    if dupes:
        print(f"{WARN} дни с одинаковыми строками (возможны дубли или два одинаковых товара):")
        for day, keys in sorted(dupes.items())[:5]:
            print(f"       {day:%d.%m.%Y}: {len(keys)}")
    else:
        print(f"{OK} дублей строк внутри дня нет")

    days = sorted(per_day)
    if days:
        missing = []
        cursor = days[0]
        while cursor <= days[-1]:
            if cursor not in per_day:
                missing.append(cursor)
            cursor = cursor.fromordinal(cursor.toordinal() + 1)
        if missing:
            print(f"{WARN} дни без строк: {', '.join(f'{d:%d.%m}' for d in missing)}")
        else:
            print(f"{OK} пропусков по дням нет: {days[0]:%d.%m} … {days[-1]:%d.%m}")
        gap = (tz_yesterday() - days[-1]).days
        if gap > 0:
            problems += 1
            print(f"{BAD} последний день {days[-1]:%d.%m.%Y}, а вчера — {tz_yesterday():%d.%m.%Y}")
        else:
            print(f"{OK} данные доведены до вчера ({days[-1]:%d.%m.%Y})")

    print("\nПо дням (строк | выручка | ВП):")
    for day in days:
        stats = per_day[day]
        print(f"  {day:%d.%m.%Y}  {stats['rows']:>4}  {stats['revenue']:>12,.2f}  {stats['profit']:>11,.2f}")

    # ── сходимость дашборда с регистром ──────────────────────────────────
    print("\nДашборд «Sales by day» против регистра:")
    sales = ss.worksheet(cfg.summary_sheet)
    block = sales.get("A19:F34", value_render_option="UNFORMATTED_VALUE")
    header_day = from_serial(block[0][1]) if len(block) > 0 and len(block[0]) > 1 else None
    total_row = block[-1] if block else []
    day_sales = total_row[1] if len(total_row) > 1 and isinstance(total_row[1], (int, float)) else 0.0
    day_profit = total_row[2] if len(total_row) > 2 and isinstance(total_row[2], (int, float)) else 0.0
    month_sales = total_row[4] if len(total_row) > 4 and isinstance(total_row[4], (int, float)) else 0.0

    expected = per_day.get(header_day, {"revenue": 0.0, "profit": 0.0})
    print(f"  дата в шапке блока: {header_day:%d.%m.%Y}" if header_day else "  дату в шапке не прочитал")
    for name, dashboard, registry in (("Sales за день", day_sales, expected["revenue"]),
                                      ("ВП за день", day_profit, expected["profit"])):
        ok = abs(dashboard - registry) < 0.05
        problems += not ok
        print(f"{OK if ok else BAD} {name}: дашборд {dashboard:,.2f} / регистр {registry:,.2f}")
    month_registry = sum(s["revenue"] for s in per_day.values())
    ok = abs(month_sales - month_registry) < 0.05
    problems += not ok
    print(f"{OK if ok else BAD} Sales за месяц: дашборд {month_sales:,.2f} / регистр {month_registry:,.2f}")

    print("\n" + ("Всё сходится." if not problems else f"Проблем: {problems}"))
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
