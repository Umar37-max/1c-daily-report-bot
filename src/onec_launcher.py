"""
Запуск тонкого клиента 1С с автологином и его закрытие после работы — чтобы
скрипты этого проекта можно было ставить на сервер и запускать по расписанию,
не требуя, чтобы кто-то заранее открыл 1С и залогинился руками.

Портировано из соседнего проекта emit_transactions (launch_1c.py), где этот
механизм уже работает на сервере. Все "живьём подтверждённые" грабли оттуда
сохранены — они одинаковые, клиент 1С тот же:
  * путь к 1cv8c.exe НЕ хардкодим: на разных машинах разные версии в
    C:\\Program Files\\1cv8\\<версия>\\bin — ищем сами, берём старшую;
  * 1cv8c.exe умеет ПЕРЕДАВАТЬ работу дочернему процессу, и окно в итоге
    принадлежит не тому pid, который вернул Popen — поэтому окно ищем по
    всему дереву процессов (_pids_in_tree), а не по одному pid;
  * окно появляется РАНЬШЕ, чем 1С готова принимать ввод — ждём по
    IsHungAppWindow (в отличие от UIA-вызова, он сам зависнуть не может);
  * на WM_CLOSE 1С показывает диалог "Завершить работу с программой?" —
    его надо подтвердить, иначе окно не закрывается и процесс убивается
    только по таймауту.

Подключаемся, как emit_transactions на сервере: СТРОКОЙ СОЕДИНЕНИЯ
(/IBConnectionString ws=...), а не по имени базы (/IBName). Имя ищется в
списке баз 1С той учётки, под которой идёт процесс, а на сервере бот
работает службой — список у неё свой или пустой, и запуск по имени там
молча не срабатывает. Адрес берём в таком порядке: запись базы в
ibases.v8i (если нашлась) → ONEC_CONNECTION_STRING из .env → значение по
умолчанию.

Отличия от emit_transactions:
  * креды и строка подключения берутся из .env ЭТОГО проекта
    (1c_login / 1c_password / ONEC_CONNECTION_STRING) — проект остаётся
    самодостаточным;
  * окно опознаётся по ONEC_WINDOW_TITLE (подстрока заголовка), как во всех
    остальных скриптах этого проекта (см. src/onec_app.py).
"""
import ctypes
import os
import subprocess
import time
from pathlib import Path

import win32api
import win32com.client
import win32con
import win32gui
import win32process
from pywinauto import Application, Desktop

_user32 = ctypes.windll.user32

PROJECT_ROOT = Path(__file__).resolve().parent.parent
ENV_PATH = PROJECT_ROOT / ".env"

# Заголовок окна 1С этого проекта (та же подстрока, что и в onec_app.attach_by_title).
WINDOW_TITLE_HINT = os.environ.get("ONEC_WINDOW_TITLE", "1С:Предприятие")

# Куда подключаться, если база не нашлась в ibases.v8i. Задаётся в .env:
# ONEC_CONNECTION_STRING=ws=https://<сервер-1С>/<база>;  Значение ниже —
# только заглушка-пример.
DEFAULT_CONNECTION_STRING = "ws=https://1c.example.com/base;"

# PID последнего запущенного НАМИ клиента — чтобы следующий запуск переиспользовал
# уже прогретый экземпляр, а не поднимал холодный (у свежего логина домашняя
# страница может долго догружаться). Файл, а не переменная: каждый запуск
# скрипта — отдельный процесс Python.
_OWN_PID_FILE = PROJECT_ROOT / ".own_1c_pid"


def _load_env(path=ENV_PATH) -> dict:
    env = {}
    if not Path(path).exists():
        return env
    for line in Path(path).read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        env[key.strip()] = value.strip()
    return env


def _ibases_paths() -> list:
    """
    Где искать ibases.v8i (список баз, зарегистрированных в самой 1С).

    СЕРВЕР: список баз у КАЖДОЙ учётки свой, а бот там крутится службой —
    Path.home() у неё может указывать вообще не туда, где базы регистрировал
    человек. Поэтому смотрим сначала свой профиль (через %APPDATA%, он
    корректен и под службой), затем профили всех пользователей машины.
    Файла может не быть вовсе — это НЕ ошибка: тогда подключаемся строкой
    соединения из .env (так и работает emit_transactions на сервере).
    """
    paths = []
    appdata = os.environ.get("APPDATA")
    if appdata:
        paths.append(Path(appdata) / "1C" / "1CEStart" / "ibases.v8i")
    paths.append(Path.home() / "AppData" / "Roaming" / "1C" / "1CEStart" / "ibases.v8i")
    try:
        for profile in Path("C:/Users").iterdir():
            paths.append(profile / "AppData" / "Roaming" / "1C" / "1CEStart" / "ibases.v8i")
    except OSError:
        pass
    seen, unique = set(), []
    for path in paths:
        key = str(path).lower()
        if key not in seen:
            seen.add(key)
            unique.append(path)
    return unique


def _read_ibases() -> dict:
    """
    Список баз, зарегистрированных в самой 1С (ibases.v8i): имя -> параметры
    (Connect, DefaultVersion). Именно отсюда берём и адрес базы, и нужную
    ВЕРСИЮ клиента — на машине их может стоять несколько.
    """
    bases: dict = {}
    source = next((p for p in _ibases_paths() if p.exists()), None)
    if source is None:
        return bases
    current = None
    for line in source.read_text(encoding="utf-8-sig", errors="replace").splitlines():
        line = line.strip()
        if line.startswith("[") and line.endswith("]"):
            current = line[1:-1]
            bases[current] = {}
        elif current and "=" in line:
            key, value = line.split("=", 1)
            bases[current][key.strip()] = value.strip()
    return bases


def _selected_base() -> tuple[str | None, dict]:
    """
    Какая база нужна: ONEC_IB_NAME из .env, иначе — единственная/первая
    подходящая из ibases.v8i. Возвращает (имя, параметры).
    """
    env = _load_env()
    bases = _read_ibases()
    name = env.get("ONEC_IB_NAME")
    if name:
        if name not in bases:
            print(f"  ⚠ База '{name}' не найдена в ibases.v8i — подключусь строкой соединения из .env.")
            return name, {}
        return name, bases[name]
    return None, {}


def _find_thin_client_exe() -> str:
    """
    Путь к 1cv8c.exe НУЖНОЙ версии.

    ЖИВЬЁМ ПОДТВЕРЖДЕНО: на этой машине установлены две версии (8.3.13.1644
    и 8.3.22.2283), а база работает на 8.3.13.1644 (её DefaultVersion в
    ibases.v8i, и именно с ней запущен клиент у пользователя). Брать просто
    "самую новую" нельзя — клиент другой версии к этой базе не подключится.
    Поэтому: версия из записи базы, затем — самая новая как запасной вариант.
    """
    env = _load_env()
    override = env.get("THIN_CLIENT_EXE_OVERRIDE")
    if override:
        if not Path(override).exists():
            raise RuntimeError(
                f"THIN_CLIENT_EXE_OVERRIDE в .env указывает на несуществующий файл: {override}"
            )
        return override

    candidates = []
    for root in (r"C:\Program Files\1cv8", r"C:\Program Files (x86)\1cv8"):
        candidates.extend(Path(root).glob("*/bin/1cv8c.exe"))
    if not candidates:
        raise RuntimeError(
            "Не нашёл 1cv8c.exe в C:\\Program Files\\1cv8\\*\\bin — 1С точно установлена? "
            "Если стоит в другом месте, пропиши путь в .env: THIN_CLIENT_EXE_OVERRIDE=..."
        )

    # ONEC_CLIENT_VERSION в .env главнее списка баз: на сервере базы в
    # списке может не быть вовсе, а версия клиента там всё равно важна —
    # клиент не той версии к этой базе просто не подключится.
    _, base = _selected_base()
    wanted = env.get("ONEC_CLIENT_VERSION") or base.get("DefaultVersion")
    if wanted:
        for path in candidates:
            if path.parent.parent.name == wanted:
                return str(path)
        print(f"  ⚠ Версия {wanted} (нужна базе) не установлена — беру самую новую из имеющихся.")

    def version_key(p: Path):
        try:
            return tuple(int(part) for part in p.parent.parent.name.split("."))
        except ValueError:
            return (0,)

    candidates.sort(key=version_key, reverse=True)
    return str(candidates[0])


def _connection_string() -> str:
    """
    Куда подключаться. Порядок:
      1) Connect из записи базы в ibases.v8i (если база там нашлась) — это
         ровно тот адрес, по которому базу открывает человек;
      2) ONEC_CONNECTION_STRING из .env — ГЛАВНЫЙ вариант для сервера, где
         базы в списке может не быть;
      3) DEFAULT_CONNECTION_STRING.

    В ibases.v8i адрес записан как  ws="https://...";  — кавычки и точку с
    запятой снимаем, 1С их в /IBConnectionString не ждёт.
    """
    _, base = _selected_base()
    connect = (base or {}).get("Connect", "")
    if connect:
        return connect.strip().rstrip(";").replace('"', "") + ";"
    return _load_env().get("ONEC_CONNECTION_STRING", DEFAULT_CONNECTION_STRING)


def _credentials() -> tuple[str, str]:
    env = _load_env()
    login = env.get("1c_login")
    password = env.get("1c_password")
    if not login or not password:
        raise RuntimeError(
            f"В {ENV_PATH} нет 1c_login / 1c_password — без них 1С не запустить автоматически."
        )
    return login, password


def find_main_window():
    """Окно 1С этого проекта, если оно открыто, иначе None."""
    matches = []
    for w in Desktop(backend="uia").windows():
        try:
            text = w.window_text()
        except Exception:
            continue
        if WINDOW_TITLE_HINT in text:
            matches.append(w)
    return matches[0] if len(matches) == 1 else None


def _wait_until_responsive(hwnd, timeout: float = 120, poll: float = 1.5) -> None:
    """
    Окно 1С появляется задолго до того, как клиент реально готов принимать
    ввод (на первом за день логине он ещё минуту грузит конфигурацию/кэши).
    IsHungAppWindow — правильная проверка: это вопрос к ОС "отвечает ли
    цикл сообщений окна", и она, в отличие от UIA-вызова, сама зависнуть
    не может.
    """
    end = time.time() + timeout
    while time.time() < end:
        if not _user32.IsHungAppWindow(hwnd):
            return
        time.sleep(poll)
    print(f"  ⚠ Окно 1С всё ещё не отвечает через {timeout} сек — продолжаю, но следующий шаг может быть медленным.")


def _pids_in_tree(root_pid: int) -> set:
    """
    root_pid плюс все его живые потомки 1cv8c.exe. ЖИВЬЁМ ПОДТВЕРЖДЕНО (в
    emit_transactions, на сервере): запущенный нами 1cv8c.exe иногда
    передаёт работу ДОЧЕРНЕМУ процессу, и видимое окно принадлежит уже ему —
    поиск по одному исходному pid в этом случае просто истекает по таймауту,
    хотя 1С явно открылась.
    """
    try:
        wmi = win32com.client.GetObject("winmgmts:")
        rows = wmi.ExecQuery("SELECT ProcessId, ParentProcessId FROM Win32_Process WHERE Name='1cv8c.exe'")
        children_of = {}
        for row in rows:
            children_of.setdefault(row.ParentProcessId, []).append(row.ProcessId)
    except Exception:
        return {root_pid}

    pids = {root_pid}
    frontier = [root_pid]
    while frontier:
        current = frontier.pop()
        for child in children_of.get(current, []):
            if child not in pids:
                pids.add(child)
                frontier.append(child)
    return pids


def _hwnd_for_pid(pid: int, timeout: float = 180, poll: float = 1.5):
    """
    Ждёт окно с нашим заголовком, принадлежащее процессу pid ИЛИ его живым
    потомкам. Именно по дереву процессов, а не по заголовку: на сервере
    рядом могут быть открыты чужие сессии 1С с таким же заголовком, и
    поиск по тексту увёл бы весь прогон в чужое окно.
    """
    end = time.time() + timeout
    while time.time() < end:
        tree = _pids_in_tree(pid)
        matches = []
        for w in Desktop(backend="uia").windows():
            try:
                hwnd = w.handle
                if not hwnd:
                    continue
                _, win_pid = win32process.GetWindowThreadProcessId(hwnd)
                if win_pid not in tree:
                    continue
                text = w.window_text()
            except Exception:
                continue
            if WINDOW_TITLE_HINT in text:
                matches.append(w)
        if len(matches) == 1:
            return matches[0]
        time.sleep(poll)
    return None


def _diagnose_stuck_launch(pid: int) -> str:
    """
    Что именно пошло не так, когда окно не дождались: какие процессы 1С живы,
    какие окна с подходящим заголовком есть и наши ли они. Без этого в логе
    остаётся только "не дождался окна", и на сервере это ничего не объясняет
    (порт из emit_transactions, где та же диагностика читается прямо из лога).
    """
    lines = []
    try:
        tree = _pids_in_tree(pid)
        lines.append(f"дерево процессов от pid {pid}: {sorted(tree)}")
    except Exception as e:  # noqa: BLE001
        lines.append(f"не смог построить дерево процессов: {e}")
        tree = {pid}
    try:
        wmi = win32com.client.GetObject("winmgmts:")
        alive = [r.ProcessId for r in wmi.ExecQuery(
            "SELECT ProcessId FROM Win32_Process WHERE Name='1cv8c.exe'")]
        lines.append(f"живые 1cv8c.exe: {sorted(alive) or 'ни одного'}")
    except Exception as e:  # noqa: BLE001
        lines.append(f"не смог перечислить процессы 1cv8c.exe: {e}")
    try:
        found_any = False
        for w in Desktop(backend="uia").windows():
            try:
                hwnd = w.handle
                if not hwnd:
                    continue
                _, win_pid = win32process.GetWindowThreadProcessId(hwnd)
                text = w.window_text()
            except Exception:
                continue
            if WINDOW_TITLE_HINT in text or "1С" in text or "1C" in text:
                found_any = True
                tag = "НАШЕ дерево" if win_pid in tree else "чужой процесс"
                lines.append(f"  окно: pid={win_pid} ({tag}) заголовок={text!r}")
        if not found_any:
            lines.append(f"  ни одного окна с '{WINDOW_TITLE_HINT}' в заголовке не найдено")
    except Exception as e:  # noqa: BLE001
        lines.append(f"не смог перечислить окна: {e}")
    return "\n".join(lines)


def _kill_proc(proc) -> None:
    """
    Не оставляем зависший процесс без окна болтаться: в emit_transactions
    такие неудачные запуски за день накапливались в шесть одновременных
    1cv8c.exe и делали каждый следующий запуск ещё медленнее.
    """
    if proc is None:
        return
    try:
        proc.terminate()
        proc.wait(timeout=10)
    except Exception:
        try:
            proc.kill()
        except Exception:
            pass


def _pid_alive(pid: int) -> bool:
    try:
        handle = win32api.OpenProcess(0x1000, False, pid)  # PROCESS_QUERY_LIMITED_INFORMATION
    except Exception:
        return False
    win32api.CloseHandle(handle)
    return True


def _read_tracked_pid():
    try:
        return int(_OWN_PID_FILE.read_text().strip())
    except (OSError, ValueError):
        return None


def _write_tracked_pid(pid: int) -> None:
    try:
        _OWN_PID_FILE.write_text(str(pid))
    except OSError:
        pass


def launch_fresh_1c(timeout: float = 180, poll: float = 1.5):
    """
    Всегда поднимает НОВЫЙ экземпляр 1С и ждёт именно его окно (по дереву
    процессов) — никогда не подхватывает чужие уже открытые окна.
    Возвращает (window, process).
    """
    login, password = _credentials()
    exe = _find_thin_client_exe()
    name, base = _selected_base()

    # Подключаемся СТРОКОЙ СОЕДИНЕНИЯ (/IBConnectionString), а не по имени
    # базы (/IBName) — так же, как emit_transactions, который уже живёт на
    # сервере. Причина: /IBName ищет базу в списке 1С ТОЙ учётки, под которой
    # идёт процесс, а на сервере бот работает службой — список у неё свой
    # (или пустой), и запуск по имени там молча не находит базу. Строка
    # соединения ни от какого профиля не зависит.
    conn = _connection_string()
    base_args = ["/IBConnectionString", conn]
    print(f"[1С] База: {conn}" + (f"  ({name})" if name else ""))

    print(f"[1С] Клиент: {exe}")
    proc = subprocess.Popen(
        [
            exe,
            "ENTERPRISE",
            *base_args,
            "/N", login,
            "/P", password,
            "/DisableStartupDialogs",
            "/DisableStartupMessages",
        ],
        # Без этого 1cv8c.exe наследует наши stdin/stdout/stderr и, пережив
        # наш процесс, держит пайп открытым — вызывающий (напр. бот с
        # capture_output=True) будет ждать EOF вечно.
        stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        close_fds=True,
    )

    win = _hwnd_for_pid(proc.pid, timeout=timeout, poll=poll)
    if win is None:
        # ЗАПАСНОЙ ВАРИАНТ: окно есть, но опознать его по дереву процессов не
        # вышло. На сервере это реально бывает — 1cv8c.exe передаёт работу
        # процессу, который WMI в дерево не связал (другое имя процесса,
        # запуск через 1cestart). Раз других 1С на сервере нет, окно с нашим
        # заголовком в единственном экземпляре — это наше окно; emit на
        # сервере именно так его и находит (ensure_1c_open по заголовку).
        by_title = find_main_window()
        if by_title is not None:
            print("  ⚠ Окно не опознал по дереву процессов — нашёл по заголовку, работаю с ним.")
            win = by_title
    if win is None:
        pid = proc.pid
        diagnostics = _diagnose_stuck_launch(pid)
        _kill_proc(proc)
        raise RuntimeError(
            f"1С (pid {pid}) не показала своё окно за {timeout} сек — возможно, не прошёл логин "
            f"(проверь 1c_login/1c_password в .env) или сервер медленный. Процесс остановлен.\n"
            f"{diagnostics}"
        )

    print("[1С] Окно появилось — жду, пока клиент начнёт отвечать...")
    _wait_until_responsive(win.handle)
    time.sleep(3)
    print("[1С] Запущена, залогинена и отвечает.")
    return win, proc


def get_or_launch_own_1c(timeout: float = 180, poll: float = 1.5):
    """
    Переиспользует экземпляр 1С, который запускал ПРЕДЫДУЩИЙ прогон этого
    проекта (если он ещё жив), иначе поднимает новый. Возвращает
    (window, process|None) — process=None означает "переиспользовали
    чужой по времени запуск", такой закрывать по окончании не нужно
    (см. close_1c).
    """
    pid = _read_tracked_pid()
    if pid is not None and _pid_alive(pid):
        win = _hwnd_for_pid(pid, timeout=8, poll=0.5)
        if win is not None:
            print(f"[1С] Переиспользую уже открытый нами экземпляр (pid {pid}).")
            return win, None
        print(f"[1С] Отслеживаемый pid {pid} жив, но окна у него нет — поднимаю новый экземпляр.")

    win, proc = launch_fresh_1c(timeout=timeout, poll=poll)
    _write_tracked_pid(proc.pid)
    return win, proc


def _is_exit_button(text: str) -> bool:
    """
    Кнопка "да, выходим" на подтверждении выхода — в обоих языках интерфейса
    (на сервере в emit_transactions 1С была с английским интерфейсом и
    показывала "Exit"/"Continue", а не "Завершить работу"/"Продолжить работу").
    Намеренно узко: не должна совпасть с кнопками, которые оставляют
    программу работать.
    """
    t = text.strip().lower().rstrip(".")
    return t.startswith("заверш") or t == "exit"


def _extra_windows(pid: int, main_hwnd: int) -> list:
    """Видимые окна процесса pid, кроме главного — кандидаты в диалог подтверждения выхода."""
    found = []

    def _cb(hwnd, _):
        if hwnd == main_hwnd or not win32gui.IsWindowVisible(hwnd):
            return True
        try:
            _, wpid = win32process.GetWindowThreadProcessId(hwnd)
        except Exception:
            return True
        if wpid == pid:
            found.append(hwnd)
        return True

    try:
        win32gui.EnumWindows(_cb, None)
    except Exception as e:
        print(f"  ⚠ Не удалось перечислить окна 1С: {e}")
    return found


def _confirm_exit_dialog(main_hwnd: int, pid: int, still_alive_fn) -> None:
    """
    Если 1С показала "Завершить работу с программой?" — подтверждаем, иначе
    WM_CLOSE ничего не даёт и процесс уходит только по жёсткому таймауту.
    Кликаем координатным посылом сообщений (vi.click_at): ЖИВЬЁМ
    ПОДТВЕРЖДЕНО в emit_transactions, что UIA Invoke на этой кнопке
    отрабатывает "успешно" и диалог даже исчезает, а процесс продолжает
    жить.
    """
    from src import virtual_input as vi

    if not main_hwnd or not pid:
        return

    buttons = []
    for hwnd in _extra_windows(pid, main_hwnd):
        try:
            dlg = Desktop(backend="uia").window(handle=hwnd)
            buttons.extend(dlg.descendants(control_type="Button"))
        except Exception:
            continue
    if not buttons:
        return

    for b in buttons:
        try:
            text = b.window_text().strip()
        except Exception:
            continue
        if _is_exit_button(text):
            try:
                r = b.rectangle()
                vi.click_at(b, (r.left + r.right) // 2, (r.top + r.bottom) // 2, settle=0.5)
                for _ in range(10):
                    if not still_alive_fn():
                        print(f"  [1С] Подтвердил выход (кнопка {text!r}).")
                        return
                    time.sleep(0.3)
                print(f"  [1С] Нажал {text!r}, но процесс ещё жив — повторю.")
                return
            except Exception as e:
                print(f"  ⚠ Клик по {text!r} не удался: {e}")

    for hwnd in _extra_windows(pid, main_hwnd):
        try:
            vi.send_keys(hwnd, "{ENTER}", settle=0.3, required=False)
            print("  [1С] Подтвердил выход Enter'ом (кнопки 'Завершить' не нашлось).")
            return
        except Exception as e:
            print(f"  ⚠ Enter в диалог выхода не прошёл: {e}")


def close_1c(win, proc=None, timeout: float = 20) -> None:
    """
    Закрывает ИМЕННО тот экземпляр 1С, с которым работал этот прогон, и
    никогда не трогает чужие окна. Сначала аккуратно (WM_CLOSE + подтверждение
    диалога выхода), и только если за timeout секунд процесс не умер — жёстко.
    """
    hwnd = win.handle if win is not None else None
    pid = proc.pid if proc is not None else None
    window_pid = None
    if hwnd:
        try:
            _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
        except Exception:
            window_pid = None
    if pid is None:
        pid = window_pid

    def _still_alive():
        if proc is not None and proc.poll() is None:
            return True
        if proc is None and pid is not None and _pid_alive(pid):
            return True
        return bool(window_pid and window_pid != pid and _pid_alive(window_pid))

    print("[1С] Закрываю 1С...")
    try:
        if hwnd:
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        end = time.time() + timeout
        while time.time() < end:
            if not _still_alive():
                print("[1С] Закрыта.")
                _clear_tracked_pid()
                return
            _confirm_exit_dialog(hwnd, window_pid or pid, _still_alive)
            time.sleep(0.5)
    except Exception as e:
        print(f"  ⚠ Аккуратное закрытие не удалось: {e}")

    print(f"  ⚠ 1С не закрылась сама за {timeout} сек — закрываю принудительно.")
    if proc is not None:
        _kill_proc(proc)
    for p in {pid, window_pid} - {None}:
        if _pid_alive(p):
            try:
                handle = win32api.OpenProcess(1, False, p)  # PROCESS_TERMINATE
                win32api.TerminateProcess(handle, 0)
                win32api.CloseHandle(handle)
            except Exception as e:
                print(f"  ⚠ Не удалось убить процесс {p}: {e}")
    _clear_tracked_pid()


def _clear_tracked_pid() -> None:
    try:
        _OWN_PID_FILE.unlink(missing_ok=True)
    except OSError:
        pass


def attach_or_launch(timeout: float = 180):
    """
    То, что нужно скриптам: вернуть (app, win, proc) — работающее окно 1С,
    подняв клиент при необходимости.

    proc=None означает "1С была уже открыта (не нами)" — такую по окончании
    работы закрывать НЕ надо, и close_if_ours() это учитывает.
    """
    from src.onec_app import _remember_active_window

    existing = find_main_window()
    if existing is not None:
        print(f"[1С] Окно '{WINDOW_TITLE_HINT}' уже открыто — использую его.")
        _, pid = win32process.GetWindowThreadProcessId(existing.handle)
        app = Application(backend="uia").connect(process=pid)
        win = app.window(handle=existing.handle)
        _remember_active_window(win)
        return app, win, None

    win, proc = get_or_launch_own_1c(timeout=timeout)
    _, pid = win32process.GetWindowThreadProcessId(win.handle)
    app = Application(backend="uia").connect(process=pid)
    win = app.window(handle=win.handle)
    _remember_active_window(win)
    return app, win, proc


def restart_1c(win=None, proc=None, timeout: float = 180):
    """
    Закрывает текущий экземпляр 1С и поднимает новый. Нужен, когда шаг
    прогона упал: причина чаще всего в застрявшем состоянии самой 1С
    (модальное окно, зависшая форма, отвалившееся соединение с сервером), и
    повтор в том же окне упрётся в то же самое. Возвращает (app, win, proc)
    — как attach_or_launch.

    Закрывает ИМЕННО то окно, с которым работали (в том числе если 1С была
    открыта не нами — по прямому указанию пользователя: "если не
    получается, закрыть 1С и пробовать заново").
    """
    from src.onec_app import _remember_active_window

    if win is not None:
        try:
            close_1c(win, proc)
        except Exception as e:
            print(f"  ⚠ Не удалось закрыть 1С перед перезапуском: {e}")
    time.sleep(3)

    fresh_win, fresh_proc = launch_fresh_1c(timeout=timeout)
    _write_tracked_pid(fresh_proc.pid)
    _, pid = win32process.GetWindowThreadProcessId(fresh_win.handle)
    app = Application(backend="uia").connect(process=pid)
    wrapped = app.window(handle=fresh_win.handle)
    _remember_active_window(wrapped)
    return app, wrapped, fresh_proc


def kill_tracked_1c() -> bool:
    """
    Убивает 1С, которую поднял последний прогон (pid из .own_1c_pid).
    Нужен кнопке "Стоп" в боте: прогон останавливается вместе с процессом
    Python, и без этого его 1С осталась бы висеть.
    """
    pid = _read_tracked_pid()
    if pid is None or not _pid_alive(pid):
        _clear_tracked_pid()
        return False
    try:
        handle = win32api.OpenProcess(1, False, pid)  # PROCESS_TERMINATE
        win32api.TerminateProcess(handle, 0)
        win32api.CloseHandle(handle)
        print(f"[1С] Остановил 1С прогона (pid {pid}).")
        return True
    except Exception as e:
        print(f"  ⚠ Не удалось остановить 1С (pid {pid}): {e}")
        return False
    finally:
        _clear_tracked_pid()


def close_if_ours(win, proc) -> None:
    """Закрыть 1С только если её подняли мы сами (proc не None)."""
    if proc is None:
        print("[1С] 1С была открыта до нас — оставляю открытой.")
        return
    close_1c(win, proc)


def launch_own(timeout: float = 180):
    """
    Работать ВСЕГДА в своём экземпляре 1С, не трогая чужие открытые окна —
    так же, как reconcile_ekvayring в соседнем проекте emit_transactions
    («Launching a fresh 1C instance, ignoring any other 1C windows already
    open»).

    Почему так, а не подключением к любому открытому окну:
      * на сервере под той же учёткой человек может держать 1С открытой и
        работать в ней руками — прогон переключал бы вкладки и перебивал
        период прямо у него под курсором;
      * чужое окно приходит в неизвестном состоянии (открытый документ,
        модальный диалог, другой пользователь 1С со своими сохранёнными
        настройками отчёта), и первый же шаг навигации падает;
      * свой экземпляр логинится тем пользователем из .env, чьи настройки
        отчёта нам и нужны.

    Экземпляр от предыдущего прогона (pid в .own_1c_pid) переиспользуется,
    если он ещё жив — это подхватывает 1С, оставшуюся после аварийно
    прерванного прогона, вместо того чтобы плодить новые окна.

    Возвращает (app, win, proc) — как attach_or_launch; закрывать по
    окончании нужно close_1c(win, proc), а не close_if_ours.
    """
    from src.onec_app import _remember_active_window

    win, proc = get_or_launch_own_1c(timeout=timeout)
    _, pid = win32process.GetWindowThreadProcessId(win.handle)
    app = Application(backend="uia").connect(process=pid)
    win = app.window(handle=win.handle)
    _remember_active_window(win)
    return app, win, proc
