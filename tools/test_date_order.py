"""
Проверка распознавания порядка дат в выгрузке 1С.

На рабочем ПК 1С печатает «01.09.2026» (день первым), на сервере —
«9/1/2026» (месяц первым). Пока порядок гадали, дни 1–12 любого месяца
читались как другой месяц и отсеивались «вне периода» — так молча пропали
01–12.09, а дни 13–31 при этом разбирались верно.

    python tools/test_date_order.py
"""
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(errors="replace")

from src.gross_profit import detect_date_order, parse_period, parse_report

HEADER = "\t".join([
    "Клиент", "Менеджер", "Серия", "Подразделение", "Заказ клиента / Реализация.Номер",
    "Имей", "Номенклатура.Марка (бренд)", "Заказ клиента / Реализация.Соглашение",
    "Период, день.Начало дня", "Клиент.Группа доступа", "Склад", "Номенклатура",
    "Вид номенклатуры", "Количество", "Стоимость закупки", "Доп. расходы", "Выручка",
    "Сумма авт. скидки", "Валовая прибыль", "Рентабельность, %",
])
HEAD = ["Клиент X", "Менеджер", "", "Магазин 1", "DOC-1", "", "Samsung",
        "BNPL PARTNER", "", "розница", "Store 1", "Товар", "Смартфоны"]


def _report(period_line, day_cells, total="300,00"):
    body = ["Валовая прибыль предприятия (USD)", f"Параметры:\tПериод: {period_line}", HEADER]
    for cell in day_cells:
        row = list(HEAD)
        row[8] = cell
        body.append("\t".join(row + ["1,000", "100,00", "", "150,00", "", "50,00", ""]))
    body.append("\t".join(["Итого"] + [""] * 12 + ["2,000", "200,00", "", total, "", "100,00", ""]))
    return "\n".join(body)


def main() -> int:
    failed = 0

    print("── порядок дат по шапке отчёта ──")
    cases = [
        ("русский формат", "01.09.2026 - 07.09.2026", ["01.09.2026", "07.09.2026"],
         date(2026, 9, 1), date(2026, 9, 7), "dmy"),
        ("американский формат", "9/1/2026 - 9/7/2026", ["9/1/2026", "9/7/2026"],
         date(2026, 9, 1), date(2026, 9, 7), "mdy"),
    ]
    for name, period, cells, start, end, expected in cases:
        text = _report(period, cells)
        order = detect_date_order(text, start, end)
        ok = order == expected
        failed += not ok
        print(f"  {'ok ' if ok else 'ОШИБКА'} {name}: {order}" + ("" if ok else f", ждали {expected}"))

    print("\n── строки попадают в период (раньше терялись) ──")
    for name, period, cells, start, end in (
        ("русский 01–07.09", "01.09.2026 - 07.09.2026", ["01.09.2026", "03.09.2026", "07.09.2026"],
         date(2026, 9, 1), date(2026, 9, 7)),
        ("американский 01–07.09", "9/1/2026 - 9/7/2026", ["9/1/2026", "9/3/2026", "9/7/2026"],
         date(2026, 9, 1), date(2026, 9, 7)),
        ("американский 20–22.09", "9/20/2026 - 9/22/2026", ["9/20/2026", "9/22/2026"],
         date(2026, 9, 20), date(2026, 9, 22)),
    ):
        rows = parse_report(_report(period, cells), start, end)
        ok = len(rows) == len(cells)
        failed += not ok
        days = ", ".join(f"{r[8]:%d.%m}" for r in rows)
        print(f"  {'ok ' if ok else 'ОШИБКА'} {name}: строк {len(rows)} из {len(cells)} [{days}]")

    print("\n── отдельные даты ──")
    for raw, order, expected in (
        ("01.09.2026", "dmy", date(2026, 9, 1)),
        ("9/1/2026", "mdy", date(2026, 9, 1)),
        ("9/22/2026", "mdy", date(2026, 9, 22)),
        ("22.09.2026", "dmy", date(2026, 9, 22)),
        ("9/22/2026", "dmy", date(2026, 9, 22)),   # 22 > 12 — однозначно
    ):
        got = parse_period(raw, order)
        ok = got == expected
        failed += not ok
        print(f"  {'ok ' if ok else 'ОШИБКА'} {raw!r:>14} ({order}) → {got}")

    print("\n" + ("Все проверки прошли." if not failed else f"ПРОВАЛОВ: {failed}"))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
