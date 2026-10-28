"""
Бот управления ежедневным отчётом.

Что умеет (только для id из BOT_ADMIN_IDS):
  ▶️ Обновить за вчера      — обычный дневной прогон
  🗓 Перезалить неделю      — удалить последние 7 дней и залить заново
  📅 За дату / период       — догнать пропущенные дни
  📸 Скриншот               — прислать картинку «Sales by day» без захода в 1С
  📊 Статус                 — что в регистре, идёт ли прогон, когда автозапуск
  ⏹ Остановить             — прервать текущий прогон

Автозапуск: каждый день в RUN_AT (.env, по умолчанию 06:30, Asia/Tashkent) —
тот же прогон, что и по кнопке «Обновить за вчера». В понедельник
(WEEKLY_REFRESH_WEEKDAY) прогон сам берёт неделю, а не один день.

Сам прогон запускается ОТДЕЛЬНЫМ ПРОЦЕССОМ (python daily_report.py ...), а
не внутри бота: автоматизация 1С работает с COM/окнами и держит поток
минутами — в асинхронном цикле бота это повесило бы и бота, и его ответы.
Ход работы прогон сам пишет в лог-группу (TELEGRAM_LOG_CHAT_ID), бот лишь
сообщает старт и итог.
"""
import asyncio
import logging
import calendar
import os
import re
import subprocess
import sys
from datetime import date, datetime, time as dtime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import requests

from telegram import BotCommand, BotCommandScopeChat, ReplyKeyboardMarkup, Update
from telegram.constants import ParseMode
from telegram.ext import Application, CommandHandler, ContextTypes, MessageHandler, filters

PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT))

from src.clock import yesterday as tz_yesterday
from src.config import load_config
from src.desktop_session import describe as describe_desktop, desktop_alive
from src.gross_profit import fmt
from src.onec_launcher import kill_tracked_1c
from src.sheet_image import capture
from src.sheets_register import get_client, last_data_row, open_register, read_period_column
from src.telegram_send import chat_access
from src.tg_log import TelegramLog

logging.basicConfig(format="%(asctime)s %(levelname)s %(name)s: %(message)s", level=logging.INFO)
logging.getLogger("httpx").setLevel(logging.WARNING)
log = logging.getLogger("dailysales")

CFG = load_config()
TASHKENT = ZoneInfo("Asia/Tashkent")

BTN_RUN = "▶️ Обновить за вчера"
BTN_WEEK = "🗓 Перезалить неделю"
BTN_RANGE = "📅 За период"
BTN_SHOT = "📸 Скриншот"
BTN_STATUS = "📊 Статус"
BTN_STOP = "⏹ Остановить"

# Сколько раз за жизнь бота пробовать догнать одну и ту же дату.
_MAX_CATCHUP = 3

MENU = ReplyKeyboardMarkup(
    [[BTN_RUN, BTN_WEEK], [BTN_RANGE, BTN_SHOT], [BTN_STATUS, BTN_STOP]],
    resize_keyboard=True,
)

# Меню наблюдателя: отчёт посмотреть можно, заливку запускать — нет.
MENU_VIEWER = ReplyKeyboardMarkup([[BTN_SHOT, BTN_STATUS]], resize_keyboard=True)

# Две даты в одном сообщении: «01.09.2026 30.09.2026», «1.9.2026 - 30.9.2026».
_RANGE_RE = re.compile(r"(\d{1,2}[.\-/]\d{1,2}[.\-/]\d{4})\D+(\d{1,2}[.\-/]\d{1,2}[.\-/]\d{4})")

HELP = (
    "<b>Ежедневный отчёт по продажам</b>\n\n"
    f"Автозапуск каждый день в {CFG.run_at} (Ташкент): 1С → лист «{CFG.register_sheet}» → "
    f"скриншот «{CFG.summary_sheet}» → сюда.\n"
    "По понедельникам вместо одного дня перезаливается вся прошлая неделя.\n\n"
    "<b>Команды</b>\n"
    "/run — обновить за вчера\n"
    "/week — перезалить последние 7 дней\n"
    "/date 19.09.2026 — один конкретный день\n"
    "/range 14.09.2026 19.09.2026 — период (или кнопка «📅 За период»)\n"
    "/month 9 — залить весь месяц заново\n"
    "/shot — прислать текущий скриншот\n"
    "/status — состояние\n"
    "/stop — прервать прогон"
)

HELP_VIEWER = (
    "<b>Ежедневный отчёт по продажам</b>\n\n"
    f"Каждый день в {CFG.run_at} (Ташкент) сюда приходит картинка «{CFG.summary_sheet}» "
    "за вчерашний день.\n\n"
    "<b>Что можно запросить в любой момент</b>\n"
    "/shot — прислать отчёт прямо сейчас\n"
    "/status — на какую дату данные в таблице\n\n"
    "Заливку данных запускает администратор."
)


def is_admin(update: Update) -> bool:
    """Кто может ЗАПУСКАТЬ заливку — BOT_ADMIN_IDS."""
    user = update.effective_user
    return bool(user) and user.id in CFG.admin_ids


def is_viewer(update: Update) -> bool:
    """
    Кто может СМОТРЕТЬ отчёт — наблюдатели и администраторы.

    Наблюдатель просит скриншот и статус когда угодно, но заливку не
    запускает: данные в таблице трогает только администратор.
    """
    user = update.effective_user
    return bool(user) and (user.id in CFG.admin_ids or user.id in CFG.viewer_ids)


def menu_for(update: Update):
    if is_admin(update):
        return MENU
    return MENU_VIEWER if is_viewer(update) else None


async def deny(update: Update) -> None:
    await update.effective_message.reply_text(
        "Заливку данных запускает только администратор.\n"
        "Тебе доступны «📸 Скриншот» и «📊 Статус».",
        reply_markup=menu_for(update),
    )


async def deny_stranger(update: Update) -> None:
    await update.effective_message.reply_text(
        "Этот бот доступен только сотрудникам из списка рассылки."
    )


# ─────────────────────────── запуск прогона ───────────────────────────

def _running(app) -> bool:
    proc = app.bot_data.get("proc")
    return proc is not None and proc.poll() is None


def _spawn(args: list) -> subprocess.Popen:
    cmd = [sys.executable, str(PROJECT_ROOT / "daily_report.py")] + args
    log.info("Запускаю: %s", " ".join(cmd))
    logs_dir = PROJECT_ROOT / "logs"
    logs_dir.mkdir(exist_ok=True)
    out = open(logs_dir / f"run_{datetime.now():%Y%m%d_%H%M%S}.log", "w", encoding="utf-8")
    return subprocess.Popen(
        cmd, cwd=str(PROJECT_ROOT), stdout=out, stderr=subprocess.STDOUT,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )


async def launch(update: Update, context: ContextTypes.DEFAULT_TYPE, args: list, label: str) -> None:
    app = context.application
    message = update.effective_message

    if _running(app):
        await message.reply_text("⏳ Прогон уже идёт. Дождись окончания или нажми «⏹ Остановить».",
                                 reply_markup=MENU)
        return

    user = update.effective_user
    who = f"{user.full_name} (id {user.id})" if user else "расписание"
    try:
        proc = _spawn(args)
    except Exception as e:  # noqa: BLE001
        log.exception("Не смог запустить прогон")
        await message.reply_text(f"❌ Не смог запустить: {type(e).__name__}: {e}")
        return

    app.bot_data["proc"] = proc
    app.bot_data["started_at"] = datetime.now(TASHKENT)
    app.bot_data["label"] = label
    TelegramLog(label).send(f"▶️ <b>{label}</b> — запустил {who}")
    await message.reply_text(f"{label} — запустил. Ход работы пишу в лог-группу.", reply_markup=MENU)

    app.create_task(_watch(context, proc, message.chat_id, label))


async def _watch(context, proc, chat_id: int, label: str) -> None:
    app = context.application
    try:
        code = await asyncio.to_thread(proc.wait)
        if app.bot_data.get("proc") is not proc:
            return  # остановили вручную — итог уже сообщён
        text = f"🏁 {label} — готово." if code == 0 else f"⚠️ {label} — завершилось с кодом {code}. Смотри лог-группу."
        await context.bot.send_message(chat_id, text)
    finally:
        if app.bot_data.get("proc") is proc:
            app.bot_data["proc"] = None


# ─────────────────────────── команды ───────────────────────────

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if is_admin(update):
        text = HELP
    elif is_viewer(update):
        text = HELP_VIEWER
    else:
        return await deny_stranger(update)
    await update.effective_message.reply_text(text, parse_mode=ParseMode.HTML,
                                              reply_markup=menu_for(update))


async def cmd_run(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_admin(update):
        return await deny(update)
    await launch(update, context, [], "Обновление за вчера")


async def cmd_week(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_admin(update):
        return await deny(update)
    await launch(update, context, ["--weekly"], "Перезалив недели")


async def cmd_date(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_admin(update):
        return await deny(update)
    if not context.args:
        await update.effective_message.reply_text("Формат: /date 19.09.2026")
        return
    day = context.args[0]
    await launch(update, context, ["--date", day], f"Обновление за {day}")


async def cmd_range(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_admin(update):
        return await deny(update)
    if len(context.args) < 2:
        await update.effective_message.reply_text("Формат: /range 14.09.2026 19.09.2026")
        return
    start, end = context.args[0], context.args[1]
    await launch(update, context, ["--from", start, "--to", end], f"Обновление {start} – {end}")


async def cmd_month(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Перезалить весь месяц: /month 9 или /month 9 2026.

    Конец периода не дальше вчерашнего дня — сегодняшние продажи ещё идут,
    и день закрывать рано.
    """
    if not is_admin(update):
        return await deny(update)
    today = tz_yesterday() + timedelta(days=1)
    try:
        month = int(context.args[0]) if context.args else today.month
        year = int(context.args[1]) if len(context.args) > 1 else today.year
        first = date(year, month, 1)
    except (ValueError, IndexError):
        await update.effective_message.reply_text("Формат: /month 9  (или /month 9 2026)")
        return
    last_day = calendar.monthrange(year, month)[1]
    last = min(date(year, month, last_day), tz_yesterday())
    if last < first:
        await update.effective_message.reply_text("Этот месяц ещё не начался.")
        return
    await launch(update, context, ["--from", fmt(first), "--to", fmt(last)],
                 f"Перезалив {fmt(first)} – {fmt(last)}")


async def cmd_shot(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Скриншот без захода в 1С — просто текущее состояние таблицы."""
    if not is_viewer(update):
        return await deny_stranger(update)
    message = await update.effective_message.reply_text("Снимаю…")
    try:
        path = await asyncio.to_thread(_make_shot)
    except Exception as e:  # noqa: BLE001
        await message.edit_text(f"❌ Не получилось: {type(e).__name__}: {e}")
        return
    await message.delete()
    with open(path, "rb") as fh:
        await context.bot.send_photo(
            update.effective_chat.id, fh,
            caption=f"«{CFG.summary_sheet}» на {datetime.now(TASHKENT):%d.%m.%Y %H:%M}",
        )


def _make_shot() -> str:
    client = get_client(CFG.credentials_path)
    ss, _ = open_register(client, CFG.spreadsheet_id, CFG.register_sheet)
    out = PROJECT_ROOT / "out" / "shot_manual.png"
    out.parent.mkdir(exist_ok=True)
    return capture(CFG.credentials_path, ss, CFG.summary_sheet, CFG.summary_range, str(out))


def _register_state() -> str:
    client = get_client(CFG.credentials_path)
    _, ws = open_register(client, CFG.spreadsheet_id, CFG.register_sheet)
    column = read_period_column(ws, CFG.register_header_row)
    last_row = last_data_row(column, CFG.register_header_row)
    days = sorted({day for _, day in column if day})
    if not days:
        return "в регистре нет дат"
    yesterday = tz_yesterday()
    gap = (yesterday - days[-1]).days
    tail = f"\nНе хватает дней до вчера: <b>{gap}</b>" if gap > 0 else "\nДанные актуальны на вчера."
    return (f"Строк: <b>{last_row - CFG.register_header_row}</b>\n"
            f"Последняя дата: <b>{fmt(days[-1])}</b>{tail}")


async def cmd_who(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Кому уходит отчёт и дойдёт ли он.

    Telegram запрещает боту писать первым: пока человек не нажал Start,
    отправка ему падает с «chat not found», и отчёт он не получит. Эта
    команда показывает такие случаи сразу, а не по факту пропавшего утром
    сообщения.
    """
    if not is_admin(update):
        return await deny(update)
    message = await update.effective_message.reply_text("Проверяю получателей…")
    targets = list(CFG.chat_ids) + ([CFG.log_chat_id] if CFG.log_chat_id else [])
    access = await asyncio.to_thread(chat_access, CFG.bot_token, targets)

    lines = ["<b>Получатели отчёта</b>"]
    problems = []
    for chat_id in CFG.chat_ids:
        reachable, note = access.get(str(chat_id), (False, "не проверен"))
        role = "админ" if int(chat_id) in CFG.admin_ids else "наблюдатель"
        mark = "✅" if reachable else "❌"
        lines.append(f"{mark} <code>{chat_id}</code> · {role} · {note}")
        if not reachable:
            problems.append(str(chat_id))
    if CFG.log_chat_id:
        reachable, note = access.get(str(CFG.log_chat_id), (False, "не проверен"))
        lines.append(f"\n{'✅' if reachable else '❌'} лог-группа <code>{CFG.log_chat_id}</code> · {note}")
    if problems:
        lines.append("\nКому отчёт НЕ дойдёт: " + ", ".join(problems))
        lines.append("Пусть откроют бота и нажмут <b>Start</b> — Telegram не даёт "
                     "боту писать первым.")
    await message.edit_text("\n".join(lines), parse_mode=ParseMode.HTML)


async def cmd_status(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_viewer(update):
        return await deny_stranger(update)
    app = context.application
    lines = []
    if _running(app):
        started = app.bot_data.get("started_at")
        minutes = (datetime.now(TASHKENT) - started).total_seconds() / 60 if started else 0
        lines.append(f"⏳ Идёт: <b>{app.bot_data.get('label', 'прогон')}</b> ({minutes:.0f} мин)")
    else:
        lines.append("Прогон сейчас не идёт.")

    jobs = context.job_queue.get_jobs_by_name("daily")
    if jobs and jobs[0].next_t:
        lines.append(f"Следующий автозапуск: <b>{jobs[0].next_t.astimezone(TASHKENT):%d.%m.%Y %H:%M}</b>")

    # Рабочий стол сервера: без него 1С не нажимается, и прогон упадёт
    # ещё до таблицы (см. src/desktop_session.py). Лучше узнать заранее.
    try:
        alive = desktop_alive()
        lines.append(("🖥 Рабочий стол: в порядке" if alive
                      else "🖥 Рабочий стол: <b>нет</b> (из RDP вышли крестиком) — "
                           "прогон вернёт его сам перед запуском 1С")
                     + f"\n<code>{describe_desktop()}</code>")
    except Exception as e:  # noqa: BLE001 — диагностика не должна ронять /status
        lines.append(f"🖥 Рабочий стол: не проверился ({type(e).__name__})")

    message = await update.effective_message.reply_text("\n".join(lines) + "\n\nСмотрю регистр…",
                                                        parse_mode=ParseMode.HTML)
    try:
        state = await asyncio.to_thread(_register_state)
    except Exception as e:  # noqa: BLE001
        state = f"❌ Регистр не прочитался: {type(e).__name__}: {e}"
    # Без reply_markup: editMessageText принимает ТОЛЬКО инлайн-клавиатуру,
    # на обычную (ReplyKeyboardMarkup) Telegram отвечает
    # «BadRequest: Inline keyboard expected» — живьём поймано на сервере
    # 20.09.2026 первым же нажатием «📊 Статус». Меню и так остаётся на
    # экране с прошлого сообщения, заново его слать не нужно.
    await message.edit_text("\n".join(lines) + "\n\n" + state, parse_mode=ParseMode.HTML)


async def cmd_stop(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_admin(update):
        return await deny(update)
    app = context.application
    proc = app.bot_data.get("proc")
    if proc is None or proc.poll() is not None:
        await update.effective_message.reply_text("Нечего останавливать.", reply_markup=MENU)
        return
    app.bot_data["proc"] = None
    proc.terminate()
    # Прогон умирает вместе со своим процессом Python, а поднятая им 1С
    # осталась бы висеть окном на сервере — снимаем и её.
    killed = await asyncio.to_thread(kill_tracked_1c)
    await update.effective_message.reply_text(
        "⏹ Прогон остановлен" + (" (1С прогона тоже закрыта)." if killed else ".")
        + " Таблица могла остаться недозаполненной — перезапусти за нужный день.",
        reply_markup=MENU,
    )


async def on_button(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    text = (update.effective_message.text or "").strip()

    # «За период»: кнопка просит даты, следующее сообщение с двумя датами
    # запускает заливку. Даты принимаем и без кнопки — если человек просто
    # прислал «01.09.2026 30.09.2026», понятно, чего он хочет.
    if text in (BTN_RUN, BTN_WEEK, BTN_RANGE, BTN_STOP) and not is_admin(update):
        return await deny(update) if is_viewer(update) else await deny_stranger(update)

    if text == BTN_RANGE:
        context.user_data["awaiting_range"] = True
        await update.effective_message.reply_text(
            "Пришли период одним сообщением, например:\n<code>01.09.2026 30.09.2026</code>\n\n"
            "Длинный период беру из 1С частями по неделе — можно хоть весь месяц.",
            parse_mode=ParseMode.HTML, reply_markup=MENU)
        return

    match = _RANGE_RE.search(text)
    if match:
        context.user_data.pop("awaiting_range", None)
        if not is_admin(update):
            return await deny(update)
        start, end = match.group(1), match.group(2)
        await launch(update, context, ["--from", start, "--to", end], f"Заливка {start} – {end}")
        return

    if context.user_data.pop("awaiting_range", False):
        await update.effective_message.reply_text(
            "Не разобрал даты. Нужно две даты в одном сообщении: <code>01.09.2026 30.09.2026</code>",
            parse_mode=ParseMode.HTML, reply_markup=MENU)
        return

    if text == BTN_RUN:
        await cmd_run(update, context)
    elif text == BTN_WEEK:
        await cmd_week(update, context)
    elif text == BTN_SHOT:
        await cmd_shot(update, context)
    elif text == BTN_STATUS:
        await cmd_status(update, context)
    elif text == BTN_STOP:
        await cmd_stop(update, context)


# ─────────────────────────── расписание ───────────────────────────

async def daily_job(context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Автозапуск. Период выбирает сам daily_report.py: обычный день — вчера,
    в день недельного перезалива — последние 7 дней.
    """
    app = context.application
    if _running(app):
        log.warning("Автозапуск пропущен: предыдущий прогон ещё идёт")
        TelegramLog("Автозапуск").send("⚠️ Пропустил автозапуск: предыдущий прогон ещё идёт.")
        return
    label = "Автозапуск"
    proc = _spawn([])
    app.bot_data["proc"] = proc
    app.bot_data["started_at"] = datetime.now(TASHKENT)
    app.bot_data["label"] = label
    TelegramLog(label).send(f"▶️ <b>{label}</b> в {datetime.now(TASHKENT):%H:%M}")

    code = await asyncio.to_thread(proc.wait)
    if app.bot_data.get("proc") is proc:
        app.bot_data["proc"] = None
    for admin_id in CFG.admin_ids:
        try:
            await context.bot.send_message(
                admin_id,
                "🏁 Автообновление готово." if code == 0
                else f"⚠️ Автообновление завершилось с кодом {code}. Смотри лог-группу.",
            )
        except Exception:  # noqa: BLE001 — недоступный админ не должен ронять задачу
            log.warning("Не смог написать админу %s", admin_id)


def _register_last_date():
    """Последняя дата в регистре — по ней видно, отстали мы или нет."""
    client = get_client(CFG.credentials_path)
    _, ws = open_register(client, CFG.spreadsheet_id, CFG.register_sheet)
    days = [day for _, day in read_period_column(ws, CFG.register_header_row) if day]
    return max(days) if days else None


WATCHDOG_INTERVAL = 300  # секунд между проверками приёма сообщений
_watchdog = {"stuck_since": None}


def _pending_updates() -> int:
    """Сколько сообщений боту лежит у Telegram и ещё не забрано."""
    resp = requests.get(f"https://api.telegram.org/bot{CFG.bot_token}/getWebhookInfo",
                        timeout=30).json()
    return int(resp.get("result", {}).get("pending_update_count", 0))


async def watchdog_job(context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Сторож приёма сообщений.

    ЖИВЬЁМ 27–28.10.2026: после сетевой ошибки (httpx.ReadError) бот
    перестал забирать сообщения, а процесс остался жить — расписание
    срабатывало, утренний отчёт ушёл, но на кнопки и команды бот молчал.
    Снаружи это выглядит как «бот умер», а start_bot.cmd его не
    перезапускает: процесс-то не упал.

    Признак зависания прямой: у Telegram копятся незабранные сообщения.
    Видим их две проверки подряд (5 минут) — завершаем процесс, и
    start_bot.cmd через 10 секунд поднимает бота заново, уже с живым
    соединением. Во время прогона не трогаем — дождёмся его конца.
    """
    app = context.application
    try:
        pending = await asyncio.to_thread(_pending_updates)
    except Exception as e:  # noqa: BLE001 — нет сети: это не зависание бота
        log.warning("Сторож: не смог спросить Telegram (%s)", e)
        return

    polling = bool(app.updater and app.updater.running)
    if pending == 0 and polling:
        if _watchdog["stuck_since"] is not None:
            log.info("Сторож: приём сообщений в порядке.")
        _watchdog["stuck_since"] = None
        return

    now = datetime.now(TASHKENT)
    if _watchdog["stuck_since"] is None:
        _watchdog["stuck_since"] = now
        log.warning("Сторож: незабранных сообщений %s, приём %s — проверю ещё раз через %s мин.",
                    pending, "идёт" if polling else "ОСТАНОВЛЕН", WATCHDOG_INTERVAL // 60)
        return

    if _running(app):
        log.warning("Сторож: приём сообщений завис, но идёт прогон — перезапущусь после него.")
        return

    minutes = (now - _watchdog["stuck_since"]).total_seconds() / 60
    log.error("Сторож: сообщения не забираются %.0f мин (в очереди %s) — перезапускаю бота.",
              minutes, pending)
    logging.shutdown()
    # os._exit, а не sys.exit: из задачи планировщика исключение поймает
    # сам планировщик, и процесс останется висеть в том же состоянии.
    os._exit(3)


async def catchup_job(context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Страховка на случай пропущенного автозапуска: раз в час сверяем
    последнюю дату регистра со вчерашней и, если отстали, догоняем сами.

    Зачем: расписание держит процесс бота, и если в 6:30 он не работал
    (сервер перезагрузили, сеанс разлогинился, бот упал) — день просто
    теряется, и это замечает человек, а не программа. ЖИВЬЁМ 22.09.2026:
    за 21.09 данные не появились вовсе.

    Чтобы не долбить 1С впустую, если за день реально нет продаж (или
    прогон падает по другой причине), на каждую дату не больше
    _MAX_CATCHUP попыток за время жизни бота.
    """
    app = context.application
    if _running(app):
        return
    now_local = datetime.now(TASHKENT)

    # Догонять можно только ПОСЛЕ времени автозапуска, с запасом в час.
    # Иначе при RUN_AT=09:30 проверка в 07:00 увидела бы, что вчерашнего дня
    # в регистре ещё нет, и запустила бы прогон раньше расписания — то есть
    # настройка времени просто перестала бы работать.
    try:
        hour, minute = (int(part) for part in CFG.run_at.split(":"))
    except ValueError:
        hour, minute = 6, 30
    scheduled = now_local.replace(hour=hour, minute=minute, second=0, microsecond=0)
    if now_local < scheduled + timedelta(hours=1) or now_local.hour >= 23:
        return

    target = tz_yesterday()
    attempts = app.bot_data.setdefault("catchup_attempts", {})
    if attempts.get(target, 0) >= _MAX_CATCHUP:
        return

    try:
        last = await asyncio.to_thread(_register_last_date)
    except Exception as e:  # noqa: BLE001 — сеть/Google, попробуем через час
        log.warning("Догоняющая проверка не удалась: %s", e)
        return
    if last is None or last >= target:
        return

    attempts[target] = attempts.get(target, 0) + 1
    start = last + timedelta(days=1)
    label = f"Догоняю {fmt(start)} – {fmt(target)}"
    log.info("%s (попытка %s)", label, attempts[target])
    TelegramLog("Догоняющий прогон").send(
        f"⏰ <b>{label}</b> — автозапуск в {CFG.run_at} не сработал "
        f"(в регистре последняя дата {fmt(last)}), запускаю сам."
    )
    proc = _spawn(["--from", fmt(start), "--to", fmt(target)])
    app.bot_data["proc"] = proc
    app.bot_data["started_at"] = now_local
    app.bot_data["label"] = label

    code = await asyncio.to_thread(proc.wait)
    if app.bot_data.get("proc") is proc:
        app.bot_data["proc"] = None
    for admin_id in CFG.admin_ids:
        try:
            await context.bot.send_message(
                admin_id,
                f"🏁 {label} — готово." if code == 0
                else f"⚠️ {label} — не получилось (код {code}). Смотри лог-группу.",
            )
        except Exception:  # noqa: BLE001
            log.warning("Не смог написать админу %s", admin_id)


async def _post_init(app: Application) -> None:
    commands = [
        BotCommand("run", "Обновить за вчера"),
        BotCommand("week", "Перезалить неделю"),
        BotCommand("date", "За конкретную дату"),
        BotCommand("range", "За период"),
        BotCommand("month", "Залить весь месяц"),
        BotCommand("shot", "Скриншот сейчас"),
        BotCommand("status", "Состояние"),
        BotCommand("who", "Кому уходит отчёт"),
        BotCommand("stop", "Остановить прогон"),
    ]
    for admin_id in CFG.admin_ids:
        try:
            await app.bot.set_my_commands(commands, scope=BotCommandScopeChat(chat_id=admin_id))
        except Exception as e:  # noqa: BLE001
            log.warning("Команды для админа %s не выставились: %s", admin_id, e)

    viewer_commands = [
        BotCommand("shot", "Прислать отчёт сейчас"),
        BotCommand("status", "На какую дату данные"),
    ]
    for viewer_id in CFG.viewer_ids:
        if viewer_id in CFG.admin_ids:
            continue
        try:
            await app.bot.set_my_commands(viewer_commands, scope=BotCommandScopeChat(chat_id=viewer_id))
        except Exception as e:  # noqa: BLE001 — не писавший боту получатель не должен ронять старт
            log.warning("Команды для наблюдателя %s не выставились: %s", viewer_id, e)


def main() -> None:
    if not CFG.bot_token:
        raise SystemExit("В .env нет TELEGRAM_BOT_TOKEN")

    app = Application.builder().token(CFG.bot_token).post_init(_post_init).build()
    app.add_handler(CommandHandler(["start", "help"], cmd_start))
    app.add_handler(CommandHandler("run", cmd_run))
    app.add_handler(CommandHandler("week", cmd_week))
    app.add_handler(CommandHandler("date", cmd_date))
    app.add_handler(CommandHandler("range", cmd_range))
    app.add_handler(CommandHandler("month", cmd_month))
    app.add_handler(CommandHandler(["shot", "screenshot"], cmd_shot))
    app.add_handler(CommandHandler("status", cmd_status))
    app.add_handler(CommandHandler("who", cmd_who))
    app.add_handler(CommandHandler("stop", cmd_stop))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_button))

    hour, minute = (int(part) for part in CFG.run_at.split(":"))
    app.job_queue.run_daily(daily_job, time=dtime(hour, minute, tzinfo=TASHKENT), name="daily")
    # Страховка: раз в час проверяем, не отстал ли регистр (см. catchup_job).
    app.job_queue.run_repeating(catchup_job, interval=3600, first=180, name="catchup")
    # Сторож: бот, переставший забирать сообщения, сам перезапускается.
    app.job_queue.run_repeating(watchdog_job, interval=WATCHDOG_INTERVAL,
                                first=WATCHDOG_INTERVAL, name="watchdog")

    log.info("Бот запущен. Автозапуск в %s (Ташкент), догоняющая проверка раз в час. "
             "Админы: %s. Наблюдатели: %s",
             CFG.run_at, CFG.admin_ids, CFG.viewer_ids)
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
