"""
Восстанавливает формулы J/K в строках регистра, где их не оказалось.

Когда нужен: прогон упал между записью данных и протяжкой формул — строки
в таблице есть, а столбцы «Группа»/«Рассрочкка» пустые, и день пропадает из
дашборда (SUMIFS ищет по J). ЖИВЬЁМ 23.09.2026: Google отказал копированию
из-за фильтра на листе, и без формул осталось 11 дней.

Формулы берутся из ближайшей строки выше, где они целы, и переносятся на
свою строку (ссылки сдвигаются). 1С для этого не нужна.

    python tools/repair_formulas.py              # показать, что будет сделано
    python tools/repair_formulas.py --apply      # восстановить
    python tools/repair_formulas.py --apply --drop-empty   # + убрать строки без канала
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(errors="replace")

from src.config import load_config
from src.sheets_register import (
    _shift_formula,
    delete_rows,
    ensure_group_formula,
    from_serial,
    get_client,
    last_data_row,
    open_register,
    read_period_column,
    sweep_excluded_groups,
)

BATCH = 2000


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--apply", action="store_true", help="записать изменения")
    parser.add_argument("--drop-empty", action="store_true",
                        help="после восстановления убрать строки, где «Группа» пустая")
    args = parser.parse_args()

    cfg = load_config()
    ss, ws = open_register(get_client(cfg.credentials_path), cfg.spreadsheet_id, cfg.register_sheet)

    first = cfg.register_header_row + 1
    column = read_period_column(ws, cfg.register_header_row)
    last = last_data_row(column, cfg.register_header_row)
    print(f"Лист «{ws.title}»: строки данных {first}–{last}")

    formulas = ws.get(f"J{first}:K{last}", value_render_option="FORMULA")
    template_row, template = None, None
    broken = []
    for offset in range(last - first + 1):
        row = (formulas[offset] if offset < len(formulas) else []) + ["", ""]
        row_index = first + offset
        if str(row[0]).startswith("="):
            template_row, template = row_index, (row[0], row[1])
        else:
            broken.append(row_index)

    if not broken:
        print("Все строки с формулами — чинить нечего.")
        return 0
    if template is None:
        print("Не нашёл ни одной целой формулы-образца — чинить не из чего.")
        return 1

    days = {}
    for row_index, day in column:
        if row_index in set(broken) and day:
            days[day] = days.get(day, 0) + 1
    print(f"\nБез формул: {len(broken)} строк")
    for day in sorted(days):
        print(f"   {day:%d.%m.%Y}: {days[day]}")
    print(f"\nОбразец формулы — строка {template_row}")

    if not args.apply:
        print("\nЭто предпросмотр. Чтобы восстановить: python tools/repair_formulas.py --apply")
        return 0

    ensure_group_formula(ws, template_row, cfg.retail_divisions)
    fresh = ws.get(f"J{template_row}:K{template_row}", value_render_option="FORMULA")
    template = ((fresh[0] if fresh else ["", ""]) + ["", ""])[:2]

    # Пишем пачками подряд идущих строк: так вместо тысяч запросов уходит
    # несколько, и лимиты Google не задеваются.
    runs = []
    for row_index in broken:
        if runs and row_index == runs[-1][1] + 1:
            runs[-1][1] = row_index
        else:
            runs.append([row_index, row_index])

    written = 0
    for start, end in runs:
        for chunk_start in range(start, end + 1, BATCH):
            chunk_end = min(chunk_start + BATCH - 1, end)
            values = [
                [_shift_formula(template[0], template_row, target),
                 _shift_formula(template[1], template_row, target) if str(template[1]).startswith("=") else ""]
                for target in range(chunk_start, chunk_end + 1)
            ]
            ws.update(values=values, range_name=f"J{chunk_start}:K{chunk_end}",
                      value_input_option="USER_ENTERED")
            written += len(values)
            print(f"   восстановлено {chunk_start}–{chunk_end}")
    print(f"\nВосстановлено строк: {written}")

    if args.drop_empty:
        values = ws.get(f"J{first}:J{last}", value_render_option="UNFORMATTED_VALUE")
        empty = [first + offset for offset, row in enumerate(values)
                 if not row or str(row[0]).strip() == ""]
        if empty:
            print(f"Убираю строки без канала продаж: {len(empty)}")
            delete_rows(ss, ws, empty)
        swept = sweep_excluded_groups(ss, ws, cfg.register_header_row, cfg.exclude_client_groups)
        if swept:
            print(f"Убрано оптовых строк: {swept}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
