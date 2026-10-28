"""
Проверяет и чинит формулу столбца «Группа» (J) в строке-образце регистра.

Формула определяет канал продаж по Подразделению (D) и Соглашению (H).
Всё, чего в ней нет, она отдаёт пустым — и прогон удаляет такие строки как
«без канала». Так терялись:

  * новый маркетплейс (Клиент «MARKETPLACE D», Соглашение «Marketplace») —
    в основном возвраты, минусовые суммы;
  * новый магазин (05.10.2026) — в перечне розницы его не было,
    и его продажи в регистр не попадали вовсе.

Чиним строку-образец — ту, с которой формула протягивается на новые строки.
Старые строки трогать не нужно: их J уже посчитан, а потерянные строки
вернутся, когда перезальёшь соответствующие дни.

Запуск:
    python tools/fix_group_formula.py            # показать, что будет сделано
    python tools/fix_group_formula.py --apply    # записать в книгу
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(errors="replace")

from src.config import load_config
from src.sheets_register import (
    add_retail_divisions,
    add_marketplace_branch,
    get_client,
    last_data_row,
    open_register,
    read_period_column,
)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--apply", action="store_true", help="записать изменения в книгу")
    args = parser.parse_args()

    cfg = load_config()
    client = get_client(cfg.credentials_path)
    ss, ws = open_register(client, cfg.spreadsheet_id, cfg.register_sheet)

    column = read_period_column(ws, cfg.register_header_row)
    last = last_data_row(column, cfg.register_header_row)
    print(f"Лист «{ws.title}»: строка-образец {last}")
    print(f"Розничные подразделения из настроек: {', '.join(cfg.retail_divisions) or '—'}")

    current = ws.get(f"J{last}", value_render_option="FORMULA")
    formula = current[0][0] if current and current[0] else ""
    if not str(formula).startswith("="):
        raise SystemExit(f"В J{last} не формула, а {formula!r} — прерываюсь.")

    updated = add_marketplace_branch(str(formula), last)
    updated = add_retail_divisions(updated, last, cfg.retail_divisions)

    missing = [d for d in cfg.retail_divisions if f'"{d}"' not in updated]
    if missing:
        print(f"\n!! Не нашёл, куда вписать: {', '.join(missing)}. "
              "Похоже, формулу в книге переписали — проверь столбец J вручную.")

    if updated == formula:
        print("\nФормула уже знает всё, что нужно — менять нечего.")
        return 0

    print("\nБыло:")
    print(str(formula).replace("\n", " "))
    print("\nСтанет:")
    print(updated.replace("\n", " "))

    if not args.apply:
        print("\nЭто предпросмотр. Чтобы записать: python tools/fix_group_formula.py --apply")
        return 0

    ws.update(values=[[updated]], range_name=f"J{last}", value_input_option="USER_ENTERED")
    print(f"\nФормула в строке {last} обновлена. Новые строки будут протягиваться уже с ней.")
    print("Теперь перезалей дни, где строки терялись, например:")
    print("   python daily_report.py --from 01.10.2026 --to 05.10.2026")
    return 0


if __name__ == "__main__":
    sys.exit(main())
