"""
Проверка разбора выгрузки, где формат чисел РАЗНЫЙ в разных колонках.

Строки повторяют формат реальной выгрузки сервера (22.09.2026): там «Стоимость
закупки» и «Валовая прибыль» в русском формате (700 100,50), а «Выручка»,
«Доп. расходы» и «Количество» — в английском (707,300.75). Один разделитель
на весь файл превращал 216.06 в 21606 и завышал выручку в десятки раз.

    python tools/test_mixed_formats.py
"""
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(errors="replace")

from src.gross_profit import COLUMNS, clean_number, parse_report, report_totals

HEADER = "\t".join([
    "Клиент", "Менеджер", "Серия", "Подразделение", "Заказ клиента / Реализация.Номер",
    "Имей", "Номенклатура.Марка (бренд)", "Заказ клиента / Реализация.Соглашение",
    "Период, день.Начало дня", "Клиент.Группа доступа", "Склад", "Номенклатура",
    "Вид номенклатуры", "Количество", "Стоимость закупки", "Доп. расходы", "Выручка",
    "Сумма авт. скидки", "Валовая прибыль", "Рентабельность, %",
])

# (кол-во, закупка, доп.расходы, выручка, скидка, ВП, рент) — как в выгрузке сервера
ROWS = [
    ("1.000", "207,36", "", "216.06", "", "8,70", "4,03"),
    ("40.000", "7 777,51", "", "8,480.00", "", "702,49", "8,28"),
    ("500.000", "107 481,29", "", "107,392.66", "", "-88,63", "-0,08"),
    ("-1.000", "-144,51", "", "-168.84", "", "-24,33", "14,41"),
    ("1.000", "12 897,76", "", "12,951.55", "", "53,79", "0,42"),
]
TOTAL = ("2,500.000", "700 100,50", "-50.25", "707,300.75", "150.40", "6 800,25", "0,96")

EXPECTED = [
    (1.0, 207.36, "", 216.06, "", 8.70, 4.03),
    (40.0, 7777.51, "", 8480.00, "", 702.49, 8.28),
    (500.0, 107481.29, "", 107392.66, "", -88.63, -0.08),
    (-1.0, -144.51, "", -168.84, "", -24.33, 14.41),
    (1.0, 12897.76, "", 12951.55, "", 53.79, 0.42),
]


def _line(prefix, tail):
    return "\t".join(list(prefix) + list(tail))


def main() -> int:
    body = [
        "Валовая прибыль предприятия (USD)",
        "Параметры:\tПериод: 9/15/2026 - 9/21/2026",
        HEADER,
    ]
    head = ["Клиент X", "Менеджер", "", "Магазин 1", "DOC-1", "", "Samsung",
            "BNPL PARTNER", "9/19/2026", "розница", "Store 1", "Товар", "Смартфоны"]
    for row in ROWS:
        body.append(_line(head, row))
    body.append(_line(["Итого"] + [""] * 12, TOTAL))
    text = "\n".join(body)

    failed = 0
    rows = parse_report(text, date(2026, 9, 15), date(2026, 9, 21))
    print(f"\nразобрано строк: {len(rows)} (ждали {len(ROWS)})")
    failed += len(rows) != len(ROWS)

    print("\n── значения по колонкам ──")
    for got_row, expected in zip(rows, EXPECTED):
        got = (got_row[13], got_row[14], got_row[15], got_row[16], got_row[17], got_row[18], got_row[19])
        ok = got == expected
        failed += not ok
        print(f"  {'ok ' if ok else 'ОШИБКА'} {got}" + ("" if ok else f"\n         ждали {expected}"))

    revenue = sum(r[16] for r in rows if isinstance(r[16], float))
    expected_revenue = sum(e[3] for e in EXPECTED)
    ok = abs(revenue - expected_revenue) < 0.01
    failed += not ok
    print(f"\n  {'ok ' if ok else 'ОШИБКА'} сумма выручки: {revenue:,.2f} (ждали {expected_revenue:,.2f})")

    totals = report_totals(text)
    ok = totals and abs(totals.get("revenue", 0) - 707300.75) < 0.01
    failed += not ok
    print(f"  {'ok ' if ok else 'ОШИБКА'} «Итого» выручка: {totals.get('revenue') if totals else None} (ждали 707 300.75)")

    print("\n── отдельные числа ──")
    for raw, sep, expected_value in (
        ("216.06", ".", 216.06),
        ("8,480.00", ".", 8480.0),
        ("107,392.66", ".", 107392.66),
        ("700 100,50", ",", 700100.5),
        ("6 800,25", ",", 6800.25),
    ):
        got = clean_number(raw, sep)
        ok = got == expected_value
        failed += not ok
        print(f"  {'ok ' if ok else 'ОШИБКА'} {raw!r:>14} ({sep}) → {got}")

    print("\n" + ("Все проверки прошли." if not failed else f"ПРОВАЛОВ: {failed}"))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
