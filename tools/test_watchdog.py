"""
Проверка сторожа приёма сообщений (bot.watchdog_job) без Telegram.

Сторож должен:
  * ничего не делать, пока сообщения забираются;
  * при первом признаке зависания только запомнить время;
  * при втором подряд — завершить процесс (start_bot.cmd поднимет бота);
  * не трогать бота, пока идёт прогон;
  * не считать зависанием отсутствие сети у сервера.

    python tools/test_watchdog.py
"""
import asyncio
import sys
from pathlib import Path
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(errors="replace")

import bot

FAILED = 0


class Exited(Exception):
    pass


def _exit(code):
    raise Exited(code)


def _context(polling=True, running=False):
    app = SimpleNamespace(updater=SimpleNamespace(running=polling), bot_data={})
    if running:
        app.bot_data["proc"] = SimpleNamespace(poll=lambda: None)
    return SimpleNamespace(application=app)


def tick(pending, polling=True, running=False):
    """Один запуск сторожа. True — если он завершил бы процесс."""
    if isinstance(pending, Exception):
        def fake():
            raise pending
    else:
        def fake():
            return pending
    bot._pending_updates = fake
    try:
        asyncio.run(bot.watchdog_job(_context(polling, running)))
    except Exited:
        return True
    return False


def check(name, got, expected):
    global FAILED
    ok = got == expected
    FAILED += not ok
    print(f"  {'ok ' if ok else 'ОШИБКА'} {name}" + ("" if ok else f" (получил {got!r})"))


def reset():
    bot._watchdog["stuck_since"] = None


def main() -> int:
    bot.os._exit = _exit
    bot.logging.shutdown = lambda: None

    print("── всё в порядке ──")
    reset()
    check("очередь пуста — не перезапускает", tick(0), False)
    check("и не запоминает зависание", bot._watchdog["stuck_since"], None)

    print("\n── зависание ──")
    reset()
    check("первая проверка: только запоминает", tick(3), False)
    check("время зависания запомнено", bot._watchdog["stuck_since"] is not None, True)
    check("вторая подряд: перезапуск", tick(3), True)

    print("\n── приём остановлен, хотя очередь пуста ──")
    reset()
    check("первая проверка", tick(0, polling=False), False)
    check("вторая: перезапуск", tick(0, polling=False), True)

    print("\n── разовый сбой проходит сам ──")
    reset()
    check("первая проверка: в очереди 1", tick(1), False)
    check("вторая: очередь разобрана — не перезапускает", tick(0), False)
    check("зависание забыто", bot._watchdog["stuck_since"], None)

    print("\n── идёт прогон ──")
    reset()
    tick(3, running=True)
    check("во время прогона не перезапускает", tick(3, running=True), False)
    check("после прогона — перезапуск", tick(3), True)

    print("\n── нет сети ──")
    reset()
    check("ошибка сети — не зависание", tick(ConnectionError("нет сети")), False)
    check("и ничего не запомнено", bot._watchdog["stuck_since"], None)

    print("\n" + ("Все проверки прошли." if not FAILED else f"ПРОВАЛОВ: {FAILED}"))
    return 1 if FAILED else 0


if __name__ == "__main__":
    sys.exit(main())
