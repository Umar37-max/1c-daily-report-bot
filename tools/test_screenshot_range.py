"""
Проверка поиска блока «Sales by day» для скриншота.

Диапазон больше не зашит в .env: коллеги дописывают строки в таблицы, и
блок уезжает (27.10.2026 — с A19:K34 на A26:J49). Проверяем, что блок
находится по шапке «Подразделение» и строке «Итого» при любом сдвиге,
а служебный столбец без шапки справа в картинку не попадает.

    python tools/test_screenshot_range.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(errors="replace")

from src.sheet_image import find_block_range

HEADER = ["Подразделение", "Sales", "Gross profit", "Необходимая сумма в день", "Sales",
          "Gross profit", "Таргет", "Выполнение", "Темп в %", "Темп", "0,00"]
LABELS = ["Ритейл", "BNPL", "Интернет Магазин", "Розница", "   Магазин 1", "   Магазин 2",
          "   Магазин 3", "   Магазин 4", "   Магазин 5", "   Магазин 6", "   Магазин 7",
          "Marketplace", " Партнёр A", " Маркетплейс A", " Маркетплейс C", " Маркетплейс B",
          " Маркетплейс D", " Optom 2", " Партнёр B", " Партнёр C"]


def _row(label, width=11):
    return [label] + [" $ 1 "] * (width - 1)


def sheet(top_rows: int, with_margin_row: bool = True):
    """Лист: верхняя таблица на top_rows строк, пустых строк нет, затем блок."""
    rows = [[""] * 15]
    rows += [_row(f"верх {i}", 15) for i in range(top_rows)]
    rows += [_row("Итого ", 15)]
    rows += [["", "10/27/2026"] + [""] * 9]
    rows += [HEADER]
    rows += [_row(label) for label in LABELS]
    rows += [_row("Итого")]
    if with_margin_row:
        rows += [["", "", "45%", "", "", "11%"] + [""] * 5]
    return rows


FAILED = 0


def check(name, got, expected):
    global FAILED
    ok = got == expected
    FAILED += not ok
    print(f"  {'ok ' if ok else 'ОШИБКА'} {name}: {got}" + ("" if ok else f" (ждал {expected})"))


def main() -> int:
    print("── блок на листе ──")
    # Нынешний лист: верхняя таблица до строки 25, блок 26–49.
    check("как на листе 27.10.2026", find_block_range(sheet(23)), "A26:J49")
    # Как было до того, как коллеги дописали строки (верх короче на 7).
    check("верхняя таблица короче", find_block_range(sheet(16)), "A19:J42")
    check("верхняя таблица длиннее", find_block_range(sheet(40)), "A43:J66")

    print("\n── края блока ──")
    check("без строки с маржой", find_block_range(sheet(23, with_margin_row=False)), "A26:J48")
    no_date = sheet(23)
    no_date[25] = [""] * 11
    check("без даты над шапкой", find_block_range(no_date), "A27:J49")
    no_total = sheet(23, with_margin_row=False)[:-1] + [[""] * 11]
    check("без «Итого» — до пустой строки", find_block_range(no_total), "A26:J47")

    print("\n── когда блока нет ──")
    check("нет шапки «Подразделение»", find_block_range([_row("что-то")] * 5), None)
    check("пустой лист", find_block_range([]), None)

    print("\n" + ("Все проверки прошли." if not FAILED else f"ПРОВАЛОВ: {FAILED}"))
    return 1 if FAILED else 0


if __name__ == "__main__":
    sys.exit(main())
