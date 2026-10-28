"""
Проверка логики записи в лист, которую можно гонять без Google и без 1С:
перенос формул на другую строку и стоп-кран при обрезанной выгрузке.

    python tools/test_sheets_logic.py
"""
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(errors="replace")

from src.gross_profit import parse_report, report_totals, verify_totals
from src.sheets_register import _shift_formula

FORMULA = (
    '=IF(K17536="BNPL";IF(OR(D17536="Магазин 1";D17536="Магазин 2");"Розница";"");'
    'IF(A17536="MARKETPLACE D";"Маркетплейс D";IF($B$2="x";"y";"")))'
)

HEADER = "\t".join([
    "Клиент", "Менеджер", "Серия", "Подразделение", "Заказ клиента / Реализация.Номер",
    "Имей", "Номенклатура.Марка (бренд)", "Заказ клиента / Реализация.Соглашение",
    "Период, день.Начало дня", "Клиент.Группа доступа", "Склад", "Номенклатура",
    "Вид номенклатуры", "Количество", "Стоимость закупки", "Доп. расходы", "Выручка",
    "Сумма авт. скидки", "Валовая прибыль", "Рентабельность, %",
])
HEAD = ["Клиент X", "Менеджер", "", "Магазин 1", "DOC-1", "", "Samsung",
        "BNPL PARTNER", "19.09.2026", "розница", "Store 1", "Товар", "Смартфоны"]


def _report(rows, total_revenue, total_profit):
    """Строки отчёта + строка «Итого», которую 1С считает по ВСЕМ строкам."""
    body = ["Валовая прибыль предприятия (USD)", "Параметры:\tПериод: 19.09.2026 - 19.09.2026", HEADER]
    for qty, cost, revenue, profit in rows:
        body.append("\t".join(HEAD + [qty, cost, "", revenue, "", profit, ""]))
    body.append("\t".join(["Итого"] + [""] * 12
                          + ["1,000", "0,00", "", total_revenue, "", total_profit, ""]))
    return "\n".join(body)


def main() -> int:
    failed = 0

    print("── перенос формулы на другую строку ──")
    moved = _shift_formula(FORMULA, 17536, 17600)
    checks = [
        ("ссылки своей строки сдвинулись", "K17600" in moved and "D17600" in moved and "A17600" in moved),
        ("старых ссылок не осталось", "17536" not in moved),
        ("абсолютная ссылка не тронута", "$B$2" in moved),
    ]
    for name, ok in checks:
        failed += not ok
        print(f"  {'ok ' if ok else 'ОШИБКА'} {name}")

    print("\n── стоп-кран при обрезанной выгрузке ──")
    full = [("1,000", "100,00", "150.00", "50,00"), ("1,000", "200,00", "260.00", "60,00")]
    text = _report(full, "410.00", "110,00")
    rows = parse_report(text, date(2026, 9, 19), date(2026, 9, 19))
    try:
        verify_totals(rows, report_totals(text))
        print("  ok  полная выгрузка проходит сверку")
    except RuntimeError as e:
        failed += 1
        print(f"  ОШИБКА полная выгрузка не прошла: {e}")

    truncated = _report(full[:1], "410.00", "110,00")  # итог полный, строк пришла половина
    rows = parse_report(truncated, date(2026, 9, 19), date(2026, 9, 19))
    try:
        verify_totals(rows, report_totals(truncated))
        failed += 1
        print("  ОШИБКА обрезанная выгрузка НЕ поймана")
    except RuntimeError:
        print("  ok  обрезанная выгрузка поймана — запись не состоится")

    print("\n" + ("Все проверки прошли." if not failed else f"ПРОВАЛОВ: {failed}"))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
