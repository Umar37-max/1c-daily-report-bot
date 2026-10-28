"""
Живой тест 1С-части: открыть «Валовая прибыль организаций - копия»,
поставить период, сформировать, выгрузить и разобрать.

    python tools/test_1c.py            # за вчера
    python tools/test_1c.py 19.09.2026
    python tools/test_1c.py 14.09.2026 19.09.2026

Сырой TSV кладётся в out/report_<дата>.tsv — чтобы посмотреть глазами,
что именно отдаёт отчёт.
"""
import sys
from datetime import date, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(errors="replace")

from src.gross_profit import COLUMNS, fetch_report, fmt, parse_report, rows_by_date
from src.clock import yesterday as tz_yesterday
from src.onec_launcher import close_1c, launch_own
from src.reports import (
    dismiss_1c_message_boxes,
    escape_stray_dialogs,
    start_message_box_watchdog,
    stop_message_box_watchdog,
)


def _parse_arg(text: str) -> date:
    day, month, year = (int(p) for p in text.split("."))
    return date(year, month, day)


def main() -> None:
    args = sys.argv[1:]
    if len(args) == 0:
        day_from = day_to = tz_yesterday()
    elif len(args) == 1:
        day_from = day_to = _parse_arg(args[0])
    else:
        day_from, day_to = _parse_arg(args[0]), _parse_arg(args[1])

    # Свой экземпляр 1С, как и в боевом прогоне (см. launch_own): чужое
    # открытое окно не трогаем и за собой закрываем.
    print("[1С] Поднимаю свой экземпляр 1С...")
    _, win, proc = launch_own()
    escape_stray_dialogs()
    dismiss_1c_message_boxes()
    start_message_box_watchdog()
    try:
        print(f"[1С] Отчёт за {fmt(day_from)} – {fmt(day_to)}")
        text = fetch_report(win, day_from, day_to)
    finally:
        stop_message_box_watchdog()   # иначе он мешает закрытию 1С
        close_1c(win, proc)

    out_dir = Path(__file__).resolve().parent.parent / "out"
    out_dir.mkdir(exist_ok=True)
    raw_path = out_dir / f"report_{day_from:%Y-%m-%d}_{day_to:%Y-%m-%d}.tsv"
    raw_path.write_text(text, encoding="utf-8")
    print(f"[1С] Сырой TSV: {raw_path} ({len(text)} символов)")

    lines = [ln for ln in text.split("\n") if ln.strip()]
    print(f"[1С] Строк в выгрузке: {len(lines)}")
    print("--- первые 6 строк выгрузки (по колонкам) ---")
    for line in lines[:6]:
        cols = line.split("\t")
        print(f"  [{len(cols)}] " + " | ".join(c.strip()[:22] for c in cols[:22]))

    rows = parse_report(text, day_from, day_to)
    print(f"\n[Разбор] Строк данных: {len(rows)}")
    if rows:
        print("--- первая строка по колонкам ---")
        for name, value in zip(COLUMNS, rows[0]):
            print(f"   {name:<38} = {value!r}")
        by_day = rows_by_date(rows)
        print("\n--- строк по датам ---")
        for day in sorted(by_day):
            total = sum(r[16] for r in by_day[day] if isinstance(r[16], float))
            profit = sum(r[18] for r in by_day[day] if isinstance(r[18], float))
            print(f"   {fmt(day)}: {len(by_day[day]):>5} строк | выручка {total:>14,.2f} | ВП {profit:>12,.2f}")
    else:
        print("!!! Ни одной строки данных не разобрано — смотри сырой TSV.")


if __name__ == "__main__":
    main()
