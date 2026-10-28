"""Запуск и подключение к окну 1С:Предприятие через pywinauto (backend UIA)."""
import time

from pywinauto import Application, Desktop
from pywinauto.findwindows import ElementNotFoundError
import re

MAIN_WINDOW_TITLE_RE = ".*(1C|1С):Предприятие.*"


def launch(exe_path: str, infobase_arg: str, user: str = "", password: str = "") -> Application:
    """Запускает 1С с указанной базой. Возвращает объект Application (pywinauto)."""
    cmd = f'"{exe_path}" {infobase_arg}'
    if user:
        cmd += f' /N"{user}"'
    if password:
        cmd += f' /P"{password}"'
    app = Application(backend="uia").start(cmd)
    return app


def wait_main_window(app: Application, timeout: int = 60):
    """Ждёт появления главного окна 1С и возвращает его."""
    end = time.time() + timeout
    last_err = None
    while time.time() < end:
        try:
            win = app.window(title_re=MAIN_WINDOW_TITLE_RE)
            win.wait("visible", timeout=2)
            return win
        except (ElementNotFoundError, TimeoutError) as e:
            last_err = e
            time.sleep(1)
    raise TimeoutError(f"Главное окно 1С не появилось за {timeout} сек") from last_err


def attach_running() -> Application:
    """Подключается к уже открытому окну 1С вместо запуска нового процесса."""
    return Application(backend="uia").connect(title_re=MAIN_WINDOW_TITLE_RE)


def find_pid_by_title(title_substring: str) -> int | None:
    """Ищет среди открытых окон рабочего стола то, чей заголовок содержит title_substring."""
    for w in Desktop(backend="uia").windows():
        if title_substring.lower() in w.window_text().lower():
            return w.process_id()
    return None


def attach_by_title(title_substring: str) -> tuple[Application, "object"]:
    """
    Подключается к уже запущенной 1С по подстроке в заголовке окна
    (например, название базы из .env → ONEC_WINDOW_TITLE).
    Возвращает (app, main_window).
    """
    pid = find_pid_by_title(title_substring)
    if pid is None:
        raise RuntimeError(
            f"Не найдено открытое окно 1С с заголовком, содержащим '{title_substring}'. "
            "Убедись, что 1С запущена и база открыта."
        )
    app = Application(backend="uia").connect(process=pid)
    win = app.top_window()
    _remember_active_window(win)
    return app, win


def _remember_active_window(win) -> None:
    """
    Сообщает src/reports.py, в какое окно адресовать клавиатурные посылы.
    Курсор-независимому вводу (см. src/virtual_input.py) нужен конкретный
    hwnd — в отличие от прежней реальной клавиатуры, которая шла туда, где
    фокус ОС.
    """
    try:
        from src.reports import set_active_window
        set_active_window(win)
    except Exception as e:  # noqa: BLE001 — не критично, есть фолбэк на foreground
        print(f"  ⚠ Не удалось запомнить активное окно 1С: {e}")


def wait_for_window_by_title_substring(title_substring: str, timeout: int = 10, poll: float = 0.25):
    """
    Ожидает появления любого открытого окна рабочего стола, чей заголовок
    содержит `title_substring` (регистронезависимо). Возвращает найденный
    объект окна (wrapper) или `None`, если таймаут истек.

    Использование: удобно после нажатия в 1С на пункт/кнопку, который открывает
    новую форму/отчет — скрипт будет ждать появления окна, но прекратит ожидание
    сразу при его появлении или по истечении `timeout`.
    """
    end = time.time() + timeout
    substr = title_substring.lower()
    while time.time() < end:
        for w in Desktop(backend="uia").windows():
            try:
                text = w.window_text() or ""
            except Exception:
                continue
            if substr in text.lower():
                return w
        time.sleep(poll)
    return None
