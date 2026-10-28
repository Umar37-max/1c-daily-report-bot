"""
Живой рабочий стол для автоматизации 1С.

ЗАЧЕМ. Прогон управляет окнами 1С, а не считает цифры в файле: ему нужен
настоящий рабочий стол сессии. Пока по RDP подключён человек, стол есть.
Стоит закрыть окно RDP — сессия остаётся жить (процессы работают, бот
отвечает), но её рабочий стол перестаёт существовать: окна не
перерисовываются, элементы не получают фокус. UI Automation при этом не
падает, а честно сообщает, что нажатие «прошло, но ничего не изменилось»:

    Could not click 'Отчеты по продажам' without a cursor - tried:
    message-child: call succeeded but had no visible effect; ...

Ровно эта ошибка и приходила утром 08.10.2026 — не баг навигации, а
отключённый RDP.

ЧТО ДЕЛАЕМ. Windows умеет перевести сессию на физическую консоль сервера:
`tscon <id> /dest:console`. Сессия остаётся той же (1С, бот и расписание
продолжают работать), но получает настоящий стол — как будто человек сидит
за сервером. Перед прогоном проверяем состояние сессии и, если она
отключена, возвращаем её на консоль сами.

Человеку при этом ничего делать не нужно: он как обычно закрывает RDP, а
прогон перед работой с 1С сам «включает экран».
"""
import ctypes
import os
import subprocess
import time
from ctypes import wintypes

# Состояния сессии терминальных служб (WTS_CONNECTSTATE_CLASS).
STATE_NAMES = {
    0: "активна",
    1: "подключена",
    2: "запрос подключения",
    3: "теневая",
    4: "отключена (RDP закрыт)",
    5: "простаивает",
    6: "слушает",
    7: "сброшена",
    8: "выключена",
    9: "инициализация",
}
STATE_ACTIVE = 0
STATE_DISCONNECTED = 4

_WTS_CURRENT_SERVER_HANDLE = 0
_WTS_CONNECT_STATE = 8
_WTS_CLIENT_PROTOCOL_TYPE = 16
_DESKTOP_SWITCHDESKTOP = 0x0100


def session_id() -> int:
    """Номер сессии, в которой выполняется прогон."""
    sid = wintypes.DWORD()
    ok = ctypes.windll.kernel32.ProcessIdToSessionId(os.getpid(), ctypes.byref(sid))
    return int(sid.value) if ok else -1


def connect_state(sid: int = None) -> int:
    """
    Состояние сессии: 0 — активна, 4 — отключена. -1, если спросить не вышло.

    Это главный признак: при закрытом RDP сессия именно «отключена», и
    любой клик по 1С в ней бесполезен.
    """
    if sid is None:
        sid = session_id()
    if sid < 0:
        return -1
    buffer = ctypes.c_void_p()
    size = wintypes.DWORD()
    ok = ctypes.windll.wtsapi32.WTSQuerySessionInformationW(
        _WTS_CURRENT_SERVER_HANDLE, sid, _WTS_CONNECT_STATE,
        ctypes.byref(buffer), ctypes.byref(size),
    )
    if not ok or not buffer:
        return -1
    try:
        return int(ctypes.cast(buffer, ctypes.POINTER(ctypes.c_int)).contents.value)
    finally:
        ctypes.windll.wtsapi32.WTSFreeMemory(buffer)


def via_rdp(sid: int = None) -> bool:
    """
    Сессия сейчас показывается через RDP (а не на консоли сервера).

    Нужно, чтобы распознать СВЁРНУТОЕ окно RDP: сессия при этом «активна»
    и рабочий стол у неё есть, но клиент RDP просит сервер ничего не
    рисовать — и клики по 1С проходят без видимого эффекта, ровно как при
    закрытом RDP. Снаружи эти случаи не различить, поэтому после такой
    ошибки сессию переводим на консоль — там окна рисуются всегда.
    """
    if sid is None:
        sid = session_id()
    if sid < 0:
        return False
    buffer = ctypes.c_void_p()
    size = wintypes.DWORD()
    ok = ctypes.windll.wtsapi32.WTSQuerySessionInformationW(
        _WTS_CURRENT_SERVER_HANDLE, sid, _WTS_CLIENT_PROTOCOL_TYPE,
        ctypes.byref(buffer), ctypes.byref(size),
    )
    if not ok or not buffer:
        return False
    try:
        # 0 — консоль, 2 — RDP.
        return int(ctypes.cast(buffer, ctypes.POINTER(ctypes.c_ushort)).contents.value) == 2
    finally:
        ctypes.windll.wtsapi32.WTSFreeMemory(buffer)


def input_desktop_available() -> bool:
    """
    Есть ли у сессии рабочий стол, принимающий ввод.

    Вторая проверка в придачу к состоянию сессии: стол пропадает и когда
    сессия заблокирована (экран блокировки — это другой рабочий стол, в
    который нашему процессу хода нет).
    """
    handle = ctypes.windll.user32.OpenInputDesktop(0, False, _DESKTOP_SWITCHDESKTOP)
    if not handle:
        return False
    ctypes.windll.user32.CloseDesktop(handle)
    return True


def desktop_alive(sid: int = None) -> bool:
    """Можно ли прямо сейчас работать с окнами."""
    state = connect_state(sid)
    if state not in (STATE_ACTIVE, -1):
        return False
    return input_desktop_available()


def _tscon(sid: int) -> tuple:
    """Перевести сессию на физическую консоль сервера."""
    try:
        done = subprocess.run(
            ["tscon", str(sid), "/dest:console"],
            capture_output=True, text=True, timeout=60,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except FileNotFoundError:
        return False, "не нашёл tscon.exe"
    except subprocess.TimeoutExpired:
        return False, "tscon не ответил за минуту"
    if done.returncode == 0:
        return True, ""
    message = (done.stderr or done.stdout or "").strip().replace("\n", " ")
    return False, message or f"tscon вернул код {done.returncode}"


def ensure_interactive_desktop(attempts: int = 3, pause: float = 3.0,
                               force_console: bool = False) -> bool:
    """
    Убедиться, что у сессии живой рабочий стол, и вернуть его, если нет.

    force_console=True — перевести на консоль и сессию, которая выглядит
    живой, но показывается через RDP: так бывает со СВЁРНУТЫМ окном RDP
    (см. via_rdp). Зовём это только после того, как клики по 1С уже прошли
    вхолостую. Подключённый по RDP человек при этом отключится — зайдёт
    заново, сессия та же.

    Возвращает True, если со столом можно работать. False — если вернуть
    не удалось; тогда прогон всё равно попробует (вдруг проверка ошиблась),
    но в логе уже будет написано, почему 1С не слушается.
    """
    sid = session_id()
    if desktop_alive(sid) and not (force_console and via_rdp(sid)):
        return True

    if desktop_alive(sid):
        print(f"[Стол] Сессия {sid} открыта через RDP, а 1С не нажимается — похоже, окно RDP "
              "свёрнуто и сервер не рисует окна. Перевожу сессию на консоль сервера "
              "(RDP отключится, зайти можно заново).")
    else:
        state = connect_state(sid)
        print(f"[Стол] Сессия {sid}: {STATE_NAMES.get(state, f'состояние {state}')} — "
              "окна 1С в таком виде не нажимаются. Возвращаю сессию на консоль сервера.")

    for attempt in range(1, attempts + 1):
        ok, error = _tscon(sid)
        if not ok:
            print(f"[Стол] Попытка {attempt}: tscon не сработал — {error}")
        time.sleep(pause)
        if desktop_alive(sid) and not (force_console and via_rdp(sid)):
            print("[Стол] Рабочий стол на месте, продолжаю.")
            return True

    print("[Стол] ! Вернуть рабочий стол не удалось. Прогон 1С, скорее всего, "
          "упрётся в «call succeeded but had no visible effect». Подключись к "
          "серверу по RDP и выходи через ярлык «Отключиться (оставить рабочим)» "
          "— см. tools/setup_rdp_console.ps1.")
    return False


def describe() -> str:
    """Короткая строка о состоянии сессии — для /status и диагностики."""
    sid = session_id()
    state = connect_state(sid)
    desktop = "есть" if input_desktop_available() else "нет"
    where = "через RDP" if via_rdp(sid) else "на консоли"
    return (f"сессия {sid}, {STATE_NAMES.get(state, f'состояние {state}')}, {where}, "
            f"рабочий стол: {desktop}")


if __name__ == "__main__":
    import sys

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    print(describe())
    print("Пригоден для 1С:", "да" if desktop_alive() else "нет")
    if "--fix" in sys.argv:
        ensure_interactive_desktop()
        print(describe())
