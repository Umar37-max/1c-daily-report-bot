"""
Проверка разбора чисел и дат в обоих форматах 1С (русская и английская
локаль машины, где работает клиент).

    python tools/test_parsing.py
"""
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(errors="replace")

from src.gross_profit import clean_number, detect_decimal_sep, parse_period, parse_report

ROW_RU = ["Клиент", "Менеджер", "Серия", "Магазин 1", "DOC-26/01", "861", "Samsung", "12 месяцев",
          "19.09.2026", "розница", "Store 1", "Смартфон", "Смартфоны",
          "1,000", "7 777,51", "", "8 424,01", "", "646,50", "7,68"]
ROW_EN = ["Клиент", "Менеджер", "Серия", "Магазин 1", "DOC-26/01", "861", "Samsung", "12 месяцев",
          "9/19/2026", "розница", "Store 1", "Смартфон", "Смартфоны",
          "1.000", "7,777.51", "", "8,424.01", "", "646.50", "7.68"]

HEADER = "Валовая прибыль предприятия (USD)\nПараметры:\tПериод: 19.09.2026 - 19.09.2026\nКлиент\tМенеджер\n"

CASES = [
    # (описание, текст, ожидаемый разделитель, ожидаемые кол-во/закупка/выручка/ВП)
    ("русская локаль", HEADER + "\t".join(ROW_RU), ",", (1.0, 7777.51, 8424.01, 646.50)),
    ("английская локаль", HEADER + "\t".join(ROW_EN), ".", (1.0, 7777.51, 8424.01, 646.50)),
]

UNIT = [
    ("213,97", ",", 213.97),
    ("1 213,97", ",", 1213.97),
    ("1 213 456,97", ",", 1213456.97),
    ("1.213,97", ",", 1213.97),
    ("-28,37", ",", -28.37),
    ("(28,37)", ",", -28.37),
    ("", ",", ""),
    ("213.97", ".", 213.97),
    ("1,213.97", ".", 1213.97),
    ("1,213,456.97", ".", 1213456.97),
    ("1 213.97", ".", 1213.97),
    ("-28.37", ".", -28.37),
    ("текст", ".", ""),
]


def main() -> int:
    failed = 0

    print("── clean_number ──")
    for raw, sep, expected in UNIT:
        got = clean_number(raw, sep)
        ok = got == expected
        failed += not ok
        print(f"  {'ok ' if ok else 'ОШИБКА'} {raw!r:>16} ({sep}) → {got!r}" + ("" if ok else f", ждали {expected!r}"))

    print("\n── определение формата и разбор строки ──")
    for name, text, expected_sep, expected_numbers in CASES:
        sep = detect_decimal_sep(text)
        ok = sep == expected_sep
        failed += not ok
        print(f"  {'ok ' if ok else 'ОШИБКА'} {name}: разделитель «{sep}»" + ("" if ok else f", ждали «{expected_sep}»"))

        rows = parse_report(text, date(2026, 9, 19), date(2026, 9, 19))
        if not rows:
            print(f"  ОШИБКА {name}: строка не разобрана")
            failed += 1
            continue
        got = (rows[0][13], rows[0][14], rows[0][16], rows[0][18])
        ok = got == expected_numbers
        failed += not ok
        print(f"  {'ok ' if ok else 'ОШИБКА'} {name}: кол-во/закупка/выручка/ВП = {got}"
              + ("" if ok else f", ждали {expected_numbers}"))

    print("\n── даты ──")
    for raw, expected in (("19.09.2026", date(2026, 9, 19)), ("9/19/2026", date(2026, 9, 19)),
                          ("19.09.2026 0:00:00", date(2026, 9, 19)), ("Клиент", None)):
        got = parse_period(raw)
        ok = got == expected
        failed += not ok
        print(f"  {'ok ' if ok else 'ОШИБКА'} {raw!r:>20} → {got}")

    print("\n" + ("Все проверки прошли." if not failed else f"ПРОВАЛОВ: {failed}"))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
