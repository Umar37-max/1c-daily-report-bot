"""
Ежедневный отчёт по продажам: 1С → лист «Регистр» Google-таблицы →
скриншот вкладки «Sales by day» → Telegram.

Что делает один прогон:
  1. открывает в 1С «Продажи → Отчеты по продажам → Валовая прибыль
     организаций - копия», ставит период и формирует отчёт;
  2. выгружает строки и разбирает их (числа очищаются от пробелов-
     разделителей, дата становится настоящей датой);
  3. в листе «Регистр» удаляет прежние строки за эти даты, дописывает
     новые, протягивает формулы J/K и убирает строки без канала продаж;
  4. снимает картинку блока «Sales by day» и рассылает её в Telegram.

Период:
  * обычный день — только вчера;
  * раз в неделю (WEEKLY_REFRESH_WEEKDAY в .env, по умолчанию понедельник)
    — последние 7 дней целиком: прошлая неделя удаляется и заливается
    заново, потому что задним числом документы в 1С ещё правятся.

Запуск:
    python daily_report.py                      # как по расписанию
    python daily_report.py --date 19.09.2026    # один конкретный день
    python daily_report.py --from 14.09.2026 --to 19.09.2026
    python daily_report.py --weekly             # принудительно недельный перезалив
    python daily_report.py --dry-run            # всё, кроме записи и отправки
    python daily_report.py --no-telegram        # записать, но не рассылать
"""
import argparse
import sys
from contextlib import nullcontext
import traceback
from datetime import date, datetime, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(errors="replace")

from src.clock import now as tz_now, today as tz_today
from src.config import load_config
from src.desktop_session import describe as describe_desktop, ensure_interactive_desktop
from src.gross_profit import (
    PERIOD_IDX,
    drop_excluded_groups,
    fetch_report,
    fmt,
    parse_report,
    report_totals,
    rows_by_date,
    verify_parsed,
    verify_totals,
)
from src.onec_launcher import (
    attach_or_launch,
    close_1c,
    close_if_ours,
    kill_tracked_1c,
    launch_own,
)
from src.reports import (
    dismiss_1c_message_boxes,
    escape_stray_dialogs,
    start_message_box_watchdog,
    stop_message_box_watchdog,
)
from src.sheet_image import capture
from src.sheets_register import get_client, open_register, write_days
from src.telegram_send import broadcast_photo, broadcast_text
from src.tg_log import StreamToTelegram, TelegramLog

OUT_DIR = Path(__file__).resolve().parent / "out"

WEEKDAYS_RU = ["понедельник", "вторник", "среда", "четверг", "пятница", "суббота", "воскресенье"]


def _parse_date(text: str) -> date:
    for pattern in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(text.strip(), pattern).date()
        except ValueError:
            continue
    raise SystemExit(f"Не понимаю дату {text!r} — жду ДД.ММ.ГГГГ")


def period_for_run(today: date, weekly_weekday: int, args) -> tuple:
    """(дата с, дата по, как это назвать в логе)."""
    yesterday = today - timedelta(days=1)
    if args.date:
        day = _parse_date(args.date)
        return day, day, "выбранный день"
    if getattr(args, "from_date", None) or args.to:
        start = _parse_date(args.from_date) if args.from_date else yesterday
        end = _parse_date(args.to) if args.to else yesterday
        if start > end:
            raise SystemExit("Начало периода позже конца.")
        return start, end, "выбранный период"
    if args.weekly or today.weekday() == weekly_weekday:
        return yesterday - timedelta(days=6), yesterday, "недельный перезалив"
    return yesterday, yesterday, "за вчера"


def _totals(rows) -> tuple:
    revenue = sum(r[16] for r in rows if isinstance(r[16], float))
    profit = sum(r[18] for r in rows if isinstance(r[18], float))
    return revenue, profit


def _money(value: float) -> str:
    return f"{value:,.2f}".replace(",", " ")


def build_caption(day_from, day_to, kind, rows, summary) -> str:
    revenue, profit = _totals(rows)
    if day_from == day_to:
        header = f"<b>Продажи за {fmt(day_from)}</b> ({WEEKDAYS_RU[day_from.weekday()]})"
    else:
        header = f"<b>Продажи {fmt(day_from)} – {fmt(day_to)}</b> ({kind})"
    lines = [
        header,
        f"Выручка: <b>{_money(revenue)}</b> · ВП: <b>{_money(profit)}</b>",
        f"Строк в регистр: {summary['kept']}"
        + (f" (удалено прежних: {summary['deleted']})" if summary["deleted"] else "")
        + (f", без канала: {summary['dropped']}" if summary["dropped"] else ""),
    ]
    # Если перезалив изменил прошлые дни — говорим об этом прямо, иначе
    # итог месяца «прыгает» без объяснений (1С правит документы задним числом).
    # Строки, для которых формула «Группа» не дала канал, прогон удаляет.
    # Молча это делать нельзя: так в сентябре потерялся один из маркетплейсов, а 05.10 —
    # весь новый магазин. Показываем, что именно выпало.
    unknown = summary.get("unknown") or {}
    if unknown:
        lines.append("")
        lines.append("<b>Не распознан канал (строки удалены):</b>")
        for (dep, agreement), count in sorted(unknown.items(), key=lambda kv: -kv[1])[:5]:
            lines.append(f"{dep or '—'} / {agreement or '—'} → {count}")

    changed = summary.get("changed_days") or {}
    if changed:
        lines.append("")
        lines.append("<b>1С изменила прошлые дни:</b>")
        for day, (was, now) in list(changed.items())[:6]:
            lines.append(f"{fmt(day)}: {_money(was)} → {_money(now)} ({now - was:+,.2f})".replace(",", " "))
    return "\n".join(lines)


def error_hint(error) -> str:
    """
    Перевод технической ошибки на человеческий — прямо в сообщение.

    «call succeeded but had no visible effect» выглядит как поломка
    навигации, а на деле почти всегда означает, что у сессии нет рабочего
    стола: с сервера вышли, закрыв окно RDP. Без подсказки это каждый раз
    выясняется заново.
    """
    text = str(error).lower()
    if "without a cursor" in text or "no visible effect" in text:
        return ("\n\n⚠️ Похоже, у сервера пропал рабочий стол — так бывает, "
                "когда из RDP вышли, просто закрыв окно, или окно RDP свёрнуто. "
                "1С в такой сессии не нажимается.\nСостояние сейчас: <code>"
                + describe_desktop() + "</code>\nЛечится ярлыком "
                "«Отключиться (оставить рабочим)» на рабочем столе сервера "
                "(tools/setup_rdp_console.ps1).")
    return ""


def date_chunks(day_from: date, day_to: date, size: int):
    """Период → куски не длиннее size дней."""
    chunks = []
    cursor = day_from
    while cursor <= day_to:
        last = min(cursor + timedelta(days=size - 1), day_to)
        chunks.append((cursor, last))
        cursor = last + timedelta(days=1)
    return chunks


class OneCSession:
    """
    Открытая 1С на время прогона: поднять, отдать окно, закрыть за собой.

    Вынесено в отдельный объект, потому что длинный период берётся из 1С
    несколькими заходами (см. date_chunks), и клиент между ними нужно
    держать открытым — а при сбое навигации перезапускать.
    """

    def __init__(self, cfg):
        self.cfg = cfg
        self.win = None
        self.proc = None
        self._close = None

    def open(self):
        if self.cfg.reuse_open_1c:
            print("[1С] Подключаюсь к открытому окну (ONEC_REUSE_OPEN=1)...")
            _, self.win, self.proc = attach_or_launch()
            self._close = close_if_ours
        else:
            print("[1С] Поднимаю свой экземпляр 1С (чужие окна не трогаю)...")
            _, self.win, self.proc = launch_own()
            self._close = close_1c
        # Свежая 1С любит встретить прогон посторонним окном (новости,
        # «база изменена», подсказка) — гасим всё, что успело всплыть, и
        # оставляем фонового сторожа, который жмёт Escape по ходу работы.
        escape_stray_dialogs()
        dismiss_1c_message_boxes()
        start_message_box_watchdog()
        return self.win

    def restart(self):
        """Поднять клиент заново — после сбоя навигации в старом окне."""
        stop_message_box_watchdog()
        self._close(self.win, self.proc)
        # Добиваем экземпляр прогона принудительно: close_1c иногда
        # оставляет процесс живым (подтверждение выхода не нажалось), а
        # launch_own такой «живой» экземпляр переиспользует — и вторая
        # попытка упирается ровно в то же зависшее окно.
        kill_tracked_1c()
        _, self.win, self.proc = launch_own()
        self._close = close_1c
        escape_stray_dialogs()
        dismiss_1c_message_boxes()
        start_message_box_watchdog()
        return self.win

    def close(self):
        # Сторожа обязательно снять ДО закрытия: иначе он гасит Escape'ом
        # диалог «Завершить работу?», который close_1c в это же время
        # подтверждает, и клиент не закрывается штатно.
        stop_message_box_watchdog()
        if self._close is not None:
            self._close(self.win, self.proc)


FETCH_RETRIES = 2  # сколько раз поднимать 1С заново, если отчёт не взялся


def _clicks_had_no_effect(error) -> bool:
    """Клик «прошёл», а в 1С ничего не поменялось — признак, что окна не рисуются."""
    text = str(error).lower()
    return "no visible effect" in text or "without a cursor" in text


def fetch_chunk(session, cfg, start_day, end_day):
    """Выгрузка за отрезок; при сбое — до FETCH_RETRIES повторов на свежей 1С."""
    for attempt in range(FETCH_RETRIES + 1):
        try:
            return fetch_report(session.win, start_day, end_day)
        except Exception as error:  # noqa: BLE001 — повтор на свежей 1С
            if attempt == FETCH_RETRIES:
                raise
            # Навигация по разделам 1С иногда просто не срабатывает: клик
            # «проходит», но страница не меняется. Повтор в том же окне
            # упрётся в то же место, поэтому клиент поднимаем заново.
            print(f"  ! Не получилось (попытка {attempt + 1}/{FETCH_RETRIES + 1}): {error}")
            # Если окна не рисуются (RDP закрыт или СВЁРНУТ), перезапуск
            # 1С не поможет — сначала даём сессии настоящий рабочий стол.
            # ЖИВЬЁМ 28.10.2026: обе попытки упали с «no visible effect»,
            # хотя сессия числилась активной.
            ensure_interactive_desktop(force_console=_clicks_had_no_effect(error))
            print("[1С] Перезапускаю клиент и пробую ещё раз...")
            session.restart()


def run(args) -> int:
    cfg = load_config()
    # Дата — по часовому поясу бизнеса (см. src/clock.py), а не по
    # локальной дате машины: на сервере она отличается на полсуток.
    today = tz_today()
    day_from, day_to, kind = period_for_run(today, cfg.weekly_refresh_weekday, args)
    span = (day_to - day_from).days + 1

    print("=" * 64)
    print(f"  Прогон {tz_now():%d.%m.%Y %H:%M} — {kind}: {fmt(day_from)} – {fmt(day_to)}")
    print("=" * 64)

    # Длинный период 1С отдаёт через буфер обмена не целиком: ЖИВЬЁМ
    # 23.09.2026 заливка за 01–22.09 вернула пустоту по двенадцати дням.
    # Поэтому берём частями — каждая часть проверяется своим «Итого».
    chunks = date_chunks(day_from, day_to, max(cfg.max_chunk_days, 1))
    if len(chunks) > 1:
        print(f"  Период {span} дн. — беру из 1С частями по {cfg.max_chunk_days}: {len(chunks)} шт.")

    # 1С водим по окнам, а для этого сессии нужен живой рабочий стол.
    # Когда с сервера выходят, закрыв окно RDP, стол пропадает и любые
    # нажатия «проходят без видимого эффекта» (см. src/desktop_session.py).
    print(f"[Стол] {describe_desktop()}")
    ensure_interactive_desktop()

    session = OneCSession(cfg)
    session.open()
    collected = []
    try:
        for index, (start_day, end_day) in enumerate(chunks, 1):
            label = f"[{index}/{len(chunks)}] " if len(chunks) > 1 else ""
            print(f"[1С] {label}Отчёт «{cfg.report_title}» за {fmt(start_day)} – {fmt(end_day)}")
            text = fetch_chunk(session, cfg, start_day, end_day)

            OUT_DIR.mkdir(exist_ok=True)
            raw_path = OUT_DIR / f"report_{start_day:%Y-%m-%d}_{end_day:%Y-%m-%d}.tsv"
            raw_path.write_text(text, encoding="utf-8")

            failures = []
            rows = parse_report(text, start_day, end_day, failures=failures)
            # Жёсткая проверка: все ли числа отчёта стали числами.
            verify_parsed(rows, failures)
            # Стоп-кран против обрезанной выгрузки.
            verify_totals(rows, report_totals(text))

            # Оптовые клиенты в регистр не идут — как галка «Клиент.Группа
            # доступа (не в списке)» в форме отчёта.
            rows, dropped_groups, by_group = drop_excluded_groups(rows, cfg.exclude_client_groups)
            if dropped_groups:
                detail = ", ".join(f"{name} — {count}" for name, count in sorted(by_group.items()))
                print(f"  Исключено строк по группе доступа: {dropped_groups} ({detail})")
            collected.append((start_day, end_day, rows))
    finally:
        session.close()

    all_rows = [row for _, _, rows in collected for row in rows]
    by_day = rows_by_date(all_rows)
    print(f"[1С] Строк данных: {len(all_rows)}")
    for day in sorted(by_day):
        revenue, profit = _totals(by_day[day])
        print(f"     {fmt(day)}: {len(by_day[day]):>5} строк | выручка {_money(revenue):>16} | ВП {_money(profit):>14}")

    if not all_rows:
        message = (
            f"⚠️ Отчёт за {fmt(day_from)} – {fmt(day_to)} вернул 0 строк. "
            "В таблицу ничего не записано (чтобы не удалить уже имеющиеся данные)."
        )
        print(message)
        if not args.no_telegram and cfg.chat_ids:
            broadcast_text(cfg.bot_token, cfg.chat_ids, message)
        return 1

    # Перезаливаем только те даты, по которым 1С реально что-то отдала:
    # день, за который отчёт пуст, а в регистре строки есть — это скорее
    # сбой выгрузки, чем «продаж не было», и старые строки надёжнее
    # оставить на месте.
    empty_days = [
        day for day in (day_from + timedelta(days=i) for i in range(span))
        if day not in by_day
    ]
    if empty_days:
        print(f"  ! Дни без строк в отчёте (в регистре не трогаю): {', '.join(fmt(d) for d in empty_days)}")

    if args.dry_run:
        print("[dry-run] В таблицу не пишу, скриншот не отправляю.")
        return 0

    # ── Google Sheets ────────────────────────────────────────────────────
    # Пишем по частям в том же порядке: если что-то сорвётся на середине
    # длинного перезалива, уже записанные дни останутся на месте.
    print(f"[Sheets] Пишу в «{cfg.register_sheet}»...")
    summary = {"deleted": 0, "appended": 0, "dropped": 0, "kept": 0, "swept": 0, "changed_days": {}}
    for start_day, end_day, rows in collected:
        if not rows:
            continue
        part = write_days(cfg, rows, set(rows_by_date(rows)))
        for key in ("deleted", "appended", "dropped", "kept", "swept"):
            summary[key] += part.get(key, 0)
        summary["changed_days"].update(part.get("changed_days") or {})
    print(f"[Sheets] Готово: дописано {summary['kept']} строк, "
          f"удалено прежних {summary['deleted']}, без канала {summary['dropped']}")

    # ── Скриншот + Telegram ──────────────────────────────────────────────
    print(f"[Скриншот] «{cfg.summary_sheet}»...")
    client = get_client(cfg.credentials_path)
    ss, _ = open_register(client, cfg.spreadsheet_id, cfg.register_sheet)
    shot = OUT_DIR / f"sales_{day_to:%Y-%m-%d}.png"
    capture(cfg.credentials_path, ss, cfg.summary_sheet, cfg.summary_range, str(shot))
    print(f"[Скриншот] {shot}")

    if args.no_telegram or not cfg.chat_ids:
        print("[Telegram] Отправка отключена.")
        return 0

    caption = build_caption(day_from, day_to, kind, all_rows, summary)
    result = broadcast_photo(cfg.bot_token, cfg.chat_ids, str(shot), caption)
    print(f"[Telegram] Отправлено: {', '.join(result['sent']) or '—'}")
    if result["failed"]:
        print(f"[Telegram] Не доставлено: {result['failed']}")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Ежедневный отчёт по продажам (1С → Google Sheets → Telegram)")
    parser.add_argument("--date", help="один день, ДД.ММ.ГГГГ")
    parser.add_argument("--from", dest="from_date", help="начало периода, ДД.ММ.ГГГГ")
    parser.add_argument("--to", help="конец периода, ДД.ММ.ГГГГ")
    parser.add_argument("--weekly", action="store_true", help="принудительно перезалить последние 7 дней")
    parser.add_argument("--dry-run", action="store_true", help="только выгрузка из 1С, без записи и отправки")
    parser.add_argument("--no-telegram", action="store_true", help="не рассылать картинку")
    parser.add_argument("--no-log", action="store_true", help="не писать ход работы в лог-группу")
    args = parser.parse_args()

    # Ход работы — в лог-группу (TELEGRAM_LOG_CHAT_ID): за прогоном видно
    # прямо в телефоне, без захода на сервер. StreamToTelegram перехватывает
    # обычные print() и шлёт их пачками, поэтому ниже по коду ничего
    # специально логировать не нужно.
    log = TelegramLog("Ежедневный отчёт по продажам")
    if args.no_log:
        log.enabled = False

    # При выключенном логе stdout не перехватываем вовсе: TelegramLog.send()
    # печатает текст и в консоль, поэтому обёртка дублировала бы каждую
    # строку прогона <pre>-блоком.
    with (StreamToTelegram(log) if log.enabled else nullcontext()):
        log.start()
        try:
            code = run(args)
        except Exception as e:  # noqa: BLE001 — падение должно быть видно в Telegram, а не только в логе
            traceback.print_exc()
            log.fail("Ежедневный отчёт", e)
            try:
                cfg = load_config()
                if not args.no_telegram and cfg.chat_ids:
                    broadcast_text(
                        cfg.bot_token,
                        cfg.chat_ids,
                        f"❌ Ежедневный отчёт не собрался: <code>{type(e).__name__}: {e}</code>"
                        + error_hint(e),
                    )
            except Exception:  # noqa: BLE001 — сообщить не удалось, это не должно прятать исходную ошибку
                pass
            return 1

    log.send("🏁 Прогон завершён." if code == 0 else f"⚠️ Прогон завершён с кодом {code}.")
    return code


if __name__ == "__main__":
    sys.exit(main())
