"""
Проверка готовности: креды, таблица, структура «Регистра», бот, получатели,
окно 1С. Запускать после любой правки .env и перед постановкой в расписание.

    python tools/check_setup.py
"""
import sys
from datetime import date, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(errors="replace")

from src.clock import today as tz_today, yesterday as tz_yesterday
from src.config import load_config
from src.gross_profit import REPORT_TITLE
from src.sheets_register import from_serial, get_client, last_data_row, open_register, read_period_column
from src.telegram_send import chat_access

OK, FAIL, WARN = "  [ок]", "  [!!]", "  [~]"


def check_google(cfg) -> bool:
    print("Google Sheets:")
    if not Path(cfg.credentials_path).exists():
        print(f"{FAIL} нет файла ключа: {cfg.credentials_path}")
        return False
    print(f"{OK} ключ на месте: {Path(cfg.credentials_path).name}")

    try:
        client = get_client(cfg.credentials_path, retries=1)
        ss, ws = open_register(client, cfg.spreadsheet_id, cfg.register_sheet)
    except Exception as e:  # noqa: BLE001
        print(f"{FAIL} {e}")
        return False

    print(f"{OK} книга: «{ss.title}»")
    print(f"{OK} лист «{ws.title}»: строк {ws.row_count}, колонок {ws.col_count}")

    headers = ws.row_values(cfg.register_header_row)
    expected = ["Клиент", "Менеджер", "Серия", "Подразделение"]
    if headers[:4] != expected:
        print(f"{FAIL} шапка в строке {cfg.register_header_row} не та: {headers[:4]} (ждал {expected})")
        return False
    print(f"{OK} шапка в строке {cfg.register_header_row}: {', '.join(headers[:4])}, ... "
          f"J={headers[9] if len(headers) > 9 else '?'}, K={headers[10] if len(headers) > 10 else '?'}")

    column = read_period_column(ws, cfg.register_header_row)
    last = last_data_row(column, cfg.register_header_row)
    days = sorted({day for _, day in column if day})
    if days:
        print(f"{OK} данные по строку {last}; даты: {days[0]:%d.%m.%Y} … {days[-1]:%d.%m.%Y}")
        gap = (tz_yesterday() - days[-1]).days
        if gap > 0:
            print(f"{WARN} не хватает дней до вчера: {gap} — догони "
                  f"`python daily_report.py --from {days[-1] + timedelta(days=1):%d.%m.%Y} "
                  f"--to {tz_yesterday():%d.%m.%Y}`")
    else:
        print(f"{WARN} в колонке «Период» нет ни одной даты")

    try:
        summary = ss.worksheet(cfg.summary_sheet)
        print(f"{OK} лист для скриншота: «{summary.title}» {cfg.summary_range}")
    except Exception:  # noqa: BLE001
        names = [w.title for w in ss.worksheets()]
        print(f"{FAIL} нет листа «{cfg.summary_sheet}». Есть: {', '.join(names)}")
        return False
    return True


def check_telegram(cfg) -> bool:
    print("\nTelegram:")
    if not cfg.bot_token:
        print(f"{FAIL} не задан TELEGRAM_BOT_TOKEN")
        return False
    import requests
    me = requests.get(f"https://api.telegram.org/bot{cfg.bot_token}/getMe", timeout=30).json()
    if not me.get("ok"):
        print(f"{FAIL} токен не принят: {me.get('description')}")
        return False
    print(f"{OK} бот: @{me['result']['username']}")

    targets = list(cfg.chat_ids) + ([cfg.log_chat_id] if cfg.log_chat_id else [])
    access = chat_access(cfg.bot_token, targets)
    ok = True
    for chat_id in cfg.chat_ids:
        reachable, note = access.get(str(chat_id), (False, "не проверен"))
        role = "админ" if int(chat_id) in cfg.admin_ids else "наблюдатель"
        if reachable:
            print(f"{OK} получатель {chat_id} ({role}): {note}")
        else:
            print(f"{FAIL} получатель {chat_id} ({role}): {note} — "
                  "пусть откроет бота и нажмёт Start, иначе отчёт ему не уйдёт")
            ok = False
    if cfg.log_chat_id:
        reachable, note = access.get(str(cfg.log_chat_id), (False, "не проверен"))
        if reachable:
            print(f"{OK} лог-группа {cfg.log_chat_id}: {note}")
        else:
            print(f"{WARN} лог-группа {cfg.log_chat_id}: {note} — добавь бота в группу")
    return ok


def check_1c() -> bool:
    print("\n1С:")
    try:
        from src.onec_launcher import find_main_window
        win = find_main_window()
    except Exception as e:  # noqa: BLE001
        print(f"{FAIL} {e}")
        return False
    if win is None:
        print(f"{WARN} окно 1С не найдено — скрипт поднимет клиент сам при прогоне")
        return True
    print(f"{OK} окно открыто: {win.window_text()}")
    print(f"{OK} отчёт: «{REPORT_TITLE}»")
    return True


def main() -> int:
    cfg = load_config()
    print("=" * 64)
    print("  Проверка настроек ежедневного отчёта")
    print("=" * 64)
    results = [check_google(cfg), check_telegram(cfg), check_1c()]
    print("\n" + "=" * 64)
    if all(results):
        print("  Всё готово.")
        return 0
    print("  Есть замечания — смотри пометки [!!] и [~] выше.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
