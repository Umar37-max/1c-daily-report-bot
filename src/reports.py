"""
Работа с отчётом "Сводная ведомость расчетов - копия 2.0" (раздел Закупки ->
Отчеты по закупкам).

Текущий (актуальный) сценарий на контрагента:
  1. open_report() — открыть раздел Закупки -> Отчеты по закупкам -> ссылка отчёта.
  2. configure_contractor_group_filter() — один раз за сессию: Настройки ->
     Расширенный -> Отборы -> фильтр "Контрагент" в режиме "В группе из
     списка" + избранное (тогда поле "Контрагент:" с кнопкой "..." появляется
     в шапке отчёта). Идемпотентна — если уже настроено, пропускает себя.
  3. set_contractor() — кнопка "..." рядом с "Контрагент:" -> список текущего
     выбора (снять старые отметки) -> "Подбор" -> ввести имя в поиск -> ждать
     конца поиска -> двойной клик по найденной строке -> Escape -> "ОК".
     Если фильтр ещё не в режиме "В группе из списка" (шаг 2 не сработал) —
     автоматически откатывается на _set_contractor_legacy() (см. ниже).
  4. Ctrl+Enter — сформировать отчёт (см. generate()).

_set_contractor_legacy() — прежний способ (до перехода на "В группе из
списка"), оставлен как запасной вариант: F4 -> двойной клик по первой строке
свободного списка (входит в режим редактирования и сразу выделяет весь текст)
-> Home -> Shift+End (выделить всё) -> Shift+Insert (текст заранее кладётся в
буфер обмена) -> Esc (закрыть подсказку автоподбора) -> ОК. Программный
Ctrl+V с клавиатуры не работает в этом поле 1С, а правый клик -> "Вставить"
из контекстного меню оказался ненадёжным — Shift+Insert единственный
проверенный стабильный способ там, где он ещё применяется.
"""
import csv
import ctypes
import datetime as _dt
import io
import os
import re
import threading
import time
from pathlib import Path

import win32api
import win32clipboard
import win32con
import win32gui
import win32process
from pywinauto import Application, keyboard

from src import virtual_input as vi

REPORT_TITLE = "Сводная ведомость расчетов - копия 2.0"

# Таймаут по умолчанию для ожидания элементов (секунды)
_DEFAULT_TIMEOUT = 15


def _open_clipboard_with_retry(retries: int = 20, delay: float = 0.15) -> None:
    """
    OpenClipboard изредка падает с (5, 'OpenClipboard', 'Отказано в
    доступе.') — буфер обмена на долю секунды занят другим процессом
    (антивирус, менеджер буфера обмена, сама 1С и т.п.). Это временно и
    почти всегда проходит за 1-2 повтора, поэтому вместо падения всего
    контрагента просто ждём и пробуем снова.
    """
    last_err = None
    for _ in range(retries):
        try:
            win32clipboard.OpenClipboard()
            return
        except Exception as e:
            last_err = e
            time.sleep(delay)
    raise last_err


def _set_clipboard_text(text: str) -> None:
    _open_clipboard_with_retry()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
    finally:
        win32clipboard.CloseClipboard()


_ACTIVE_HWND = None


def set_active_window(win) -> None:
    """
    Запоминает окно 1С, в которое по умолчанию уходят клавиатурные посылы
    (_send_keys/_send_ctrl_key без явной цели). Раньше такие вызовы шли
    через реальную клавиатуру в то окно, у которого фокус ОС; курсор-
    независимый ввод адресуется конкретному hwnd, поэтому его надо знать.
    Вызывается из onec_app.attach_by_title() и из onec_launcher.
    """
    global _ACTIVE_HWND
    hwnd = getattr(win, "handle", win)
    if isinstance(hwnd, int) and hwnd:
        _ACTIVE_HWND = hwnd
        # Сразу поднимаем сторожа всплывающих окон 1С: любой прогон,
        # с какой бы точки входа он ни начался, дальше защищён от
        # модального окна, которое иначе намертво останавливает работу.
        start_message_box_watchdog()


def _focused_hwnd_of_1c():
    """
    Окно 1С, у которого СЕЙЧАС клавиатурный фокус.

    Критично для виртуального ввода: в диалоге "Подбор" фокус держит окно
    каталога "Контрагенты", а не главное окно. Реальная клавиатура сама
    попадала куда надо, а посланные клавиши адресуются конкретному hwnd —
    и если слать их в главное окно, имя контрагента вставляется не туда
    (именно на этом ломался Подбор).

    GUITHREADINFO.hwndFocus спрашивает у самой системы, какое окно в потоке
    1С владеет фокусом, — это ровно та цель, куда ушла бы настоящая клавиша.
    """
    if not _ACTIVE_HWND:
        return win32gui.GetForegroundWindow()

    class _GUITHREADINFO(ctypes.Structure):
        _fields_ = [
            ("cbSize", ctypes.c_ulong),
            ("flags", ctypes.c_ulong),
            ("hwndActive", ctypes.c_void_p),
            ("hwndFocus", ctypes.c_void_p),
            ("hwndCapture", ctypes.c_void_p),
            ("hwndMenuOwner", ctypes.c_void_p),
            ("hwndMoveSize", ctypes.c_void_p),
            ("hwndCaret", ctypes.c_void_p),
            ("rcCaret", ctypes.c_long * 4),
        ]

    try:
        thread_id, _ = win32process.GetWindowThreadProcessId(_ACTIVE_HWND)
        info = _GUITHREADINFO()
        info.cbSize = ctypes.sizeof(_GUITHREADINFO)
        if ctypes.windll.user32.GetGUIThreadInfo(thread_id, ctypes.byref(info)):
            for candidate in (info.hwndFocus, info.hwndActive):
                if candidate:
                    return int(candidate)
    except Exception:
        pass
    return _ACTIVE_HWND


def _default_hwnd():
    return _focused_hwnd_of_1c() or _ACTIVE_HWND or win32gui.GetForegroundWindow()


def _auto_verify(element):
    """
    Готовая проверка "клик сработал" для переключаемых элементов, или None,
    если проверять нечем.

    Радио-кнопки ("Вид: Расширенный"), чекбоксы, вкладки и строки списков
    сами сообщают своё состояние через UIA — если после клика элемент стал
    выбранным/отмеченным, клик точно дошёл. Именно на таких элементах
    паттерны 1С чаще всего "срабатывают" вхолостую, поэтому проверка тут
    даёт максимум пользы и не требует писать её руками на каждый вызов.
    """
    for getter in (
        lambda: element.is_selected(),
        lambda: element.get_toggle_state() == 1,
    ):
        try:
            if getter() is not None:
                return lambda: bool(getter())
        except Exception:
            continue
    return None


def _click_input_preserve_cursor(element, *args, **kwargs):
    """
    Клик по элементу БЕЗ реального курсора (см. src/virtual_input.py):
    UIA-паттерны -> LegacyIAccessible -> посланное WM_LBUTTONDOWN/UP в тот
    дочерний hwnd, которому реально принадлежат пиксели элемента.

    Имя оставлено прежним (вызовов по всему проекту десятки), но курсор
    теперь не трогается вообще — поэтому всё работает и на сервере без
    подключённой сессии, где SetCursorPos просто падает
    ("There is no active desktop required for moving mouse cursor").

    coords=(x, y) в kwargs (как у pywinauto click_input) трактуется как
    точка в КЛИЕНТСКИХ координатах элемента — так её и передавали здесь.
    """
    coords = kwargs.get("coords")
    if coords is not None:
        try:
            rect = element.rectangle()
            vi.click_at(element, rect.left + coords[0], rect.top + coords[1], settle=0.15)
            return
        except Exception:
            pass

    # Есть активный рабочий стол -> ТОЧНО исходное поведение проекта:
    # настоящий click_input по элементу с возвратом курсора на место.
    # Никаких подмен на UIA-паттерны: у 1С они регулярно "срабатывают"
    # вхолостую, и именно на этом ломались Подбор и переходы по разделам.
    if vi.ALLOW_REAL_INPUT:
        original_pos = win32api.GetCursorPos()
        try:
            element.click_input(*args, **kwargs)
        finally:
            try:
                win32api.SetCursorPos(original_pos)
            except Exception:
                pass
        return

    # Рабочего стола нет (сервер) -> курсор-независимый путь: посланные
    # сообщения и UIA-паттерны (см. src/virtual_input.py).
    #
    # Для переключаемых элементов (радио-кнопка "Расширенный", чекбоксы,
    # вкладки, строки списков) проверяем результат АВТОМАТИЧЕСКИ: элемент
    # должен стать выбранным/отмеченным. Без этого клик "успешно" уходил в
    # паттерн, который ничего не делает, и код шёл дальше как ни в чём не
    # бывало. Для остальных элементов проверить нечем — там как раньше.
    vi.click(element, settle=0.15, verify=_auto_verify(element),
             verify_timeout=2.5, verify_poll=0.25)


def _double_click_input_preserve_cursor(element, *args, **kwargs):
    """
    Двойной клик по элементу (в Подборе им открывается найденная строка
    каталога контрагентов).

    Есть рабочий стол -> настоящий двойной клик с возвратом курсора
    (исходное поведение проекта). Нет -> посланные сообщения.
    """
    if vi.ALLOW_REAL_INPUT:
        original_pos = win32api.GetCursorPos()
        try:
            element.double_click_input(*args, **kwargs)
        finally:
            try:
                win32api.SetCursorPos(original_pos)
            except Exception:
                pass
        return
    vi.double_click(element, settle=0.2)


def _force_foreground(hwnd: int) -> None:
    """
    Принудительно делает hwnd активным окном Windows. Обычный
    SetForegroundWindow часто молча блокируется системой, если вызывающий
    процесс не является "активным" — AttachThreadInput на время вызова
    снимает это ограничение. Без реального OS-фокуса клавиатурные события
    (даже отправленные через keybd_event) уходят не в то окно.
    """
    # В виртуальном режиме OS-фокус не нужен вообще: и клавиши, и клики
    # адресуются конкретному hwnd (см. src/virtual_input.py). А на сервере
    # с отключённой сессией SetForegroundWindow ещё и не работает в
    # принципе — падает с "SetForegroundWindow ... No error message is
    # available" и только засоряет лог. Поднимаем окно в z-order и выходим.
    if not vi.ALLOW_REAL_INPUT:
        try:
            win32gui.BringWindowToTop(hwnd)
        except Exception:
            pass
        return

    fg_hwnd = win32gui.GetForegroundWindow()
    cur_thread = win32api.GetCurrentThreadId()
    fg_thread, _ = win32process.GetWindowThreadProcessId(fg_hwnd) if fg_hwnd else (0, 0)
    target_thread, _ = win32process.GetWindowThreadProcessId(hwnd)

    attached_fg = attached_target = False
    try:
        if fg_thread and fg_thread != cur_thread:
            win32process.AttachThreadInput(cur_thread, fg_thread, True)
            attached_fg = True
        if target_thread and target_thread != cur_thread:
            win32process.AttachThreadInput(cur_thread, target_thread, True)
            attached_target = True
        win32gui.BringWindowToTop(hwnd)
        win32gui.SetForegroundWindow(hwnd)
    finally:
        if attached_fg:
            win32process.AttachThreadInput(cur_thread, fg_thread, False)
        if attached_target:
            win32process.AttachThreadInput(cur_thread, target_thread, False)


def _send_ctrl_key(vk_code: int, target=None) -> None:
    """
    Ctrl+<клавиша> в окно 1С.

    Есть активный рабочий стол -> старый проверенный путь: keybd_event
    (аппаратный уровень) — именно он годами работает в этом проекте для
    Ctrl+A/Ctrl+C/Ctrl+S в табличном документе отчёта, куда pywinauto
    type_keys/send_keys не доходят.

    Нет рабочего стола (сервер с отключённой сессией) -> посланные
    WM_KEYDOWN/UP с явными virtual-key кодами и синхронизацией состояния
    Ctrl в потоке-получателе (см. virtual_input.send_vk_combo — без неё
    Ctrl+C приходит как обычная 'C').
    """
    hwnd = getattr(target, "handle", target) if target is not None else _default_hwnd()
    if not vi.ALLOW_REAL_INPUT:
        vi.send_vk_combo(hwnd, vk_code, ctrl=True, settle=0.05)
        return

    win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
    time.sleep(0.05)
    win32api.keybd_event(vk_code, 0, 0, 0)
    time.sleep(0.05)
    win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.05)
    win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)


def _send_keys(target, keys: str, settle: float = 0.15, required: bool = True) -> bool:
    """
    Клавиши в 1С. Синтаксис — как у pywinauto ("^a", "{ENTER}", "{DOWN 3}",
    "+{INSERT}").

    Есть активный рабочий стол -> ТОЧНО как раньше: element.type_keys(...)
    для элемента и keyboard.send_keys(...) в АКТИВНОЕ окно, когда цели нет.

    Это принципиально: в диалоге "Подбор" фокус держит окно каталога
    "Контрагенты", а не главное окно 1С. Курсор-независимый путь адресует
    клавиши конкретному hwnd и для этого выводит его вперёд — если так
    отправить Shift+Insert, фокус уедет на главное окно, и имя контрагента
    вставится не туда (именно это и сломало Подбор). Поэтому при живом
    рабочем столе ничего не перенаправляем.

    Рабочего стола нет (сервер с отключённой сессией) -> курсор-независимый
    путь (см. virtual_input.send_keys).
    """
    if vi.ALLOW_REAL_INPUT:
        try:
            if target is not None and hasattr(target, "type_keys"):
                target.type_keys(keys)
            else:
                keyboard.send_keys(keys)
            time.sleep(settle)
            return True
        except Exception:
            if required:
                raise
            return False

    hwnd = target if target is not None else _default_hwnd()
    return vi.send_keys(hwnd, keys, settle=settle, required=required)


def _scroll(hwnd_or_ctrl, x: int, y: int, notches: int) -> None:
    """
    Колесо мыши в точке (x, y) экрана. notches > 0 — вверх по содержимому.

    Есть активный рабочий стол -> настоящее колесо (1С его обрабатывает
    гарантированно). Нет -> посланное WM_MOUSEWHEEL в тот hwnd, которому
    принадлежит точка.
    """
    if vi.ALLOW_REAL_INPUT:
        try:
            from pywinauto import mouse as _mouse

            original = win32api.GetCursorPos()
            try:
                _mouse.scroll(coords=(x, y), wheel_dist=notches)
            finally:
                win32api.SetCursorPos(original)
            return
        except Exception:
            pass

    hwnd = vi.hwnd_at_point(x, y, fallback=getattr(hwnd_or_ctrl, "handle", hwnd_or_ctrl))
    if not hwnd:
        return
    lparam = (y << 16) | (x & 0xFFFF)
    for _ in range(abs(notches)):
        wparam = ((120 if notches > 0 else -120) << 16) & 0xFFFFFFFF
        win32gui.PostMessage(hwnd, win32con.WM_MOUSEWHEEL, wparam, lparam)
        time.sleep(0.01)


def _wait_for(fn, description: str, timeout: float = _DEFAULT_TIMEOUT, poll: float = 0.5):
    """
    Вызывает fn() каждые poll секунд, пока не вернёт не-None / не выкинет
    исключение, или пока не истечёт timeout.  Возвращает результат fn().
    """
    deadline = time.time() + timeout
    last_err = None
    while time.time() < deadline:
        try:
            result = fn()
            if result is not None:
                return result
        except Exception as exc:
            last_err = exc
        time.sleep(poll)
    msg = f"Не дождались: {description} (таймаут {timeout} сек)"
    if last_err:
        msg += f" — последняя ошибка: {last_err}"
    raise TimeoutError(msg)


def _report_snapshot(win):
    """Возвращает снимок текущих строк отчёта: количество и текст каждой строки."""
    try:
        items = win.descendants(control_type="DataItem")
        return len(items), tuple(item.window_text() for item in items)
    except Exception:
        return 0, ()


def _wait_for_report_refresh(win, old_snapshot, timeout: float = 20.0, poll: float = 0.5, stable_time: float = 1.0):
    """Ждёт обновления отчёта после его формирования в 1С."""
    deadline = time.time() + timeout
    changed = False
    stable_start = None
    min_wait = time.time() + min(0.3, timeout)

    while time.time() < deadline:
        current_snapshot = _report_snapshot(win)
        if current_snapshot != old_snapshot:
            if not changed:
                changed = True
                stable_start = time.time()
            elif time.time() - stable_start >= stable_time:
                return
        else:
            if changed:
                if time.time() - stable_start >= stable_time:
                    return
            elif time.time() >= min_wait:
                return

        time.sleep(poll)

    # Таймаут прошёл, считаем, что больше ждать нельзя.
    return


def _find_generation_indicator(win):
    """
    Ищет всплывающую подсказку "Отчёт формируется..." — 1С показывает её
    поверх формы, пока отчёт реально считается. Ищем по тексту среди ВСЕХ
    потомков без привязки к конкретному control_type — так надёжнее, чем
    гадать точный тип элемента (тултип/панель/группа).
    """
    try:
        for el in win.descendants():
            try:
                text = el.window_text()
            except Exception:
                continue
            if text and "формируется" in text:
                return el
    except Exception:
        pass
    return None


def _wait_for_generation_indicator(
    win,
    appear_timeout: float = 3.0,
    disappear_timeout: float = 120.0,
    poll: float = 0.2,
) -> None:
    """
    Ждёт появления индикатора "Отчёт формируется..." и затем — его
    исчезновения. Это прямой и надёжный сигнал готовности отчёта: не важно,
    формируется он быстро или долго — мы просто смотрим на сам индикатор,
    а не гадаем по таймауту.

    Если индикатор не появился за appear_timeout секунд — отчёт, видимо,
    сформировался быстрее, чем мы успели его увидеть; тогда просто идём
    дальше (это нормально для маленьких отчётов).
    """
    deadline_appear = time.time() + appear_timeout
    appeared = False
    while time.time() < deadline_appear:
        if _find_generation_indicator(win) is not None:
            appeared = True
            break
        time.sleep(poll)

    if not appeared:
        return

    print("  [1С] Вижу 'Отчёт формируется...' — жду, пока исчезнет...")
    deadline_disappear = time.time() + disappear_timeout
    while time.time() < deadline_disappear:
        if _find_generation_indicator(win) is None:
            print("  [1С] 'Отчёт формируется...' исчез — отчёт готов.")
            return
        time.sleep(poll)

    print(f"  [1С] ⚠ 'Отчёт формируется...' не исчез за {disappear_timeout} сек — иду дальше.")


def get_report_filter_line(win) -> str:
    """
    Возвращает строку с "Отбор" из тела сформированного отчёта — 1С
    печатает её прямо над таблицей, например:
    'Отбор: Контрагент В группе из списка "Client-Example-A"'.

    Ищем "отбор" ГДЕ УГОДНО в тексте элемента (не только в начале) —
    в реальном отчёте это может быть частью более длинного многострочного
    блока (например, вместе с "Тип расчетов: Все" в одном элементе), и
    строгая проверка "начинается с" такое пропускает. Пустая строка, если
    не нашли вообще ничего похожего.
    """
    try:
        for el in win.descendants():
            try:
                text = el.window_text()
            except Exception:
                continue
            if text and "отбор" in text.lower():
                return text.strip()
    except Exception:
        pass
    return ""


def verify_contractor_filter(win, expected_name: str):
    """
    Проверяет, что отчёт реально сформирован по нужному контрагенту —
    сверяет expected_name со строкой "Отбор: ..." в теле отчёта.

    Возвращает:
      True  — строка "Отбор:" найдена и содержит expected_name.
      False — строка "Отбор:" найдена, но с ДРУГИМ контрагентом (реальная
              проблема: set_contractor не сработал, отчёт по чужому фильтру).
      None  — строку "Отбор:" вообще не удалось найти в дереве элементов —
              не доказательство ошибки, а неубедительный результат (скорее
              баг поиска по UI, чем реальная проблема с фильтром);
              вызывающий код не должен считать это провалом и дёргать
              переустановку контрагента заново на пустом месте.
    """
    filter_line = get_report_filter_line(win)
    if not filter_line:
        return None
    return _names_match(expected_name, filter_line)


def open_section_report(win, section_title: str, menu_item_title: str, report_title: str,
                        wait_after: float = 4.0, attempts: int = 3) -> None:
    """
    Полная навигация до отчёта: раздел -> пункт меню -> ссылка отчёта, с
    ПРОВЕРКОЙ на каждом шаге.

    ЖИВЬЁМ ПОДТВЕРЖДЕНО, почему проверки обязательны: клик по разделу в
    левой панели 1С срабатывает НЕ ВСЕГДА с первого раза (раздел остаётся
    прежним, а в UIA-дереве при этом висят элементы старой страницы).
    Раньше код шёл дальше вслепую и падал уже на следующем шаге —
    "не найден пункт меню" или, хуже, кликал по координатам чужой
    страницы, открывая посторонние окна ("Выбор поля", "Расшифровка"),
    которые модально блокировали весь прогон.
    """
    for attempt in range(1, attempts + 1):
        dismiss_1c_message_boxes()
        _click_input_preserve_cursor(win.child_window(title=section_title, control_type="TabItem"))
        time.sleep(1.5)

        menu_item = win.child_window(title=menu_item_title, control_type="MenuItem")
        if not menu_item.exists(timeout=4):
            print(f"  [1С] Раздел '{section_title}' не открылся (попытка {attempt}/{attempts}) — повторяю...")
            continue

        _click_input_preserve_cursor(menu_item)
        time.sleep(1.5)

        link = win.child_window(title=report_title, control_type="Hyperlink")
        if not link.exists(timeout=5):
            print(f"  [1С] Ссылка '{report_title}' не появилась (попытка {attempt}/{attempts}) — повторяю...")
            continue

        open_report_link(win, link, wait_after=wait_after)
        return

    raise RuntimeError(
        f"Не удалось открыть отчёт '{report_title}' "
        f"(раздел '{section_title}' -> '{menu_item_title}') за {attempts} попыток."
    )


def click_until(element, verify, timeout: float = 4.0, attempts: int = 3, what: str = "") -> bool:
    """
    Кликает по элементу, пока verify() не подтвердит результат.

    Зачем: при ВИРТУАЛЬНОМ вводе (сервер без активной сессии) у 1С клик
    часто "проходит" вхолостую — UIA-паттерн отрабатывает без ошибки, а на
    экране ничего не меняется. virtual_input умеет перебирать слои
    (invoke -> select -> expand -> legacy -> посланные сообщения), но без
    проверки он останавливается на ПЕРВОМ, который не упал. verify даёт
    ему повод идти дальше — ровно так это решено в соседнем проекте
    emit_transactions, где тот же приём и работает на сервере.
    """
    for attempt in range(1, attempts + 1):
        try:
            vi.click(element, settle=0.2, verify=verify, verify_timeout=timeout, verify_poll=0.25)
            return True
        except Exception as e:
            if attempt == attempts:
                raise RuntimeError(f"Не удалось нажать {what or 'элемент'}: {e}") from e
            print(f"  [1С] {what or 'Клик'}: попытка {attempt}/{attempts} не дала результата — повторяю...")
            dismiss_1c_message_boxes()
            time.sleep(1.0)
    return False


def open_section_report(win, section_title: str, menu_item_title: str, report_title: str,
                        wait_after: float = 4.0) -> None:
    """
    Навигация до отчёта: раздел -> пункт меню -> ссылка отчёта, где КАЖДЫЙ
    шаг подтверждается появлением следующего элемента (см. click_until).
    Без подтверждения виртуальный клик по разделу 1С регулярно не
    срабатывает, а код идёт дальше вслепую и падает уже на следующем шаге.
    """
    dismiss_1c_message_boxes()

    menu_item = win.child_window(title=menu_item_title, control_type="MenuItem")
    click_until(
        win.child_window(title=section_title, control_type="TabItem"),
        verify=lambda: menu_item.exists(timeout=0.5),
        what=f"раздел '{section_title}'",
    )

    link = win.child_window(title=report_title, control_type="Hyperlink")
    click_until(
        menu_item,
        verify=lambda: link.exists(timeout=0.5),
        what=f"пункт '{menu_item_title}'",
    )

    open_report_link(win, link, wait_after=wait_after)


def open_report_link(win, link, wait_after: float = 4.0) -> None:
    """
    Клик по ссылке отчёта в разделе + Enter ТОЛЬКО если клик отчёт не открыл.

    ЖИВЬЁМ ПОДТВЕРЖДЁННЫЙ БАГ (после перехода на курсор-независимый ввод):
    раньше здесь всегда слались и клик, и Enter. Клик теперь доходит
    надёжно и сам открывает отчёт — а лишний Enter уходит уже В ТАБЛИЦУ
    открытого отчёта и вызывает расшифровку ячейки, на что 1С отвечает
    модальным "Поле объекта недоступно для чтения (Расшифровка)". Это
    окно блокирует главное окно, и ВСЕ следующие шаги прогона падают с
    ElementNotFoundError, будто исчезли все элементы.

    Поэтому: кликаем, ждём появления кнопки "Сформировать" (признак, что
    отчёт открылся) и жмём Enter только если её так и нет.
    """
    # Открытие проверяем ПО ФАКТУ: когда отчёт открылся, страница раздела
    # сменилась и сама ссылка пропала из активного дерева.
    #
    # Без проверки при виртуальном вводе отчёт не открывался вовсе:
    # vi.click останавливается на первом слое, который не упал, а у ссылок
    # 1С это invoke() — он отрабатывает вхолостую. Дальше Enter уходил в
    # никуда (ссылка так и не получила фокус), и следующий шаг искал поля
    # формы, которой нет ("не дождались поля периода").
    def _opened() -> bool:
        return not link.exists(timeout=0.5)

    try:
        click_until(link, verify=_opened, timeout=4.0, attempts=2, what="ссылку отчёта")
    except Exception:
        # Запасной путь: дать ссылке фокус и нажать Enter (в списке раздела
        # клик только выделяет ссылку, открывает именно Enter).
        vi.focus(link)
        _send_keys(link, "{ENTER}", required=False)
        time.sleep(1.5)
        if not _opened():
            raise RuntimeError(
                "Не удалось открыть отчёт по ссылке: страница раздела не сменилась."
            )

    time.sleep(wait_after)
    dismiss_1c_message_boxes()


def open_report(win) -> None:
    """
    Открывает отчёт заново через навигацию Закупки -> Отчеты по закупкам.
    Верхние вкладки документов в 1С не имеют доступных через UIA заголовков
    (проверено: control_type="TabItem" — это только пункты левого меню разделов,
    у document-вкладок сверху вообще нет отдельного именованного элемента),
    поэтому надёжно проверить "уже открыта ли вкладка" нельзя — каждый вызов
    открывает новую вкладку с отчётом (старые остаются в фоне, это не мешает
    работе — их Edit/CheckBox не попадают в дерево активной вкладки).
    """
    open_section_report(win, "Закупки", "Отчеты по закупкам", REPORT_TITLE)
    # Увеличиваем паузу — форме отчёта нужно время, чтобы полностью прогрузить
    # все элементы (Edit-поля, CheckBox-ы, кнопки). 2 сек часто не хватало.
    time.sleep(4)


def _contractor_edit(win, timeout: float = _DEFAULT_TIMEOUT):
    """
    Поле ввода контрагента — 3-е поле Edit на форме (после двух полей периода).
    Ждём, пока на форме появятся хотя бы 3 Edit-поля (форма полностью загружена).
    """
    def _find():
        edits = win.descendants(control_type="Edit")
        if len(edits) >= 3:
            return edits[2]
        return None

    return _wait_for(_find, "поле ввода контрагента (3-й Edit)", timeout=timeout)


_FORM_CHECKBOX_LABELS = ("Период:", "Контрагент:")


def _names_match(expected: str, actual: str) -> bool:
    """
    Регистронезависимое сравнение имени контрагента (подстрока).

    ЖИВЬЁМ ПОДТВЕРЖДЁННЫЙ БАГ: 'OPT-Example-Client' (реальное отображение в
    1С) vs 'Opt-example-client' (имя вкладки Google Sheets, передаётся как
    name) — один и тот же контрагент, разный регистр отдельных букв.
    Обычное `in` регистрозависимо, поэтому set_contractor() решал, что
    перезапись на месте "не сработала" (хотя поле показывало ПРАВИЛЬНОЕ
    значение), зря уходил в 'Подбор', а тот падал с необработанным
    исключением по той же причине (поиск строки результата в каталоге
    тоже сравнивал регистрозависимо и не находил свою же строку).
    """
    return expected.strip().casefold() in actual.casefold()


def _dialog_first_row(win, timeout: float = _DEFAULT_TIMEOUT):
    """
    Первая строка списка в диалоге "Контрагент" — CheckBox с непустым
    текстом, который не совпадает с постоянными подписями самой формы
    отчёта ("Период:", "Контрагент:"). Это надёжнее индекса [0], который
    зависит от порядка обхода дерева.

    ПОДТВЕРЖДЕНО ЖИВЬЁМ (реальный баг, найденный на этой функции): в
    диалоге "Контрагент" есть unlabeled CheckBox-иконки тулбара (клон/
    отметить все и т.п.) — их window_text() пустой, но раньше отсеивались
    только по совпадению с _FORM_CHECKBOX_LABELS, поэтому пустая иконка
    проходила фильтр и ложно принималась за "уже есть строка", хотя
    настоящей строки не было вообще — это и ломало ветвление "первый через
    Подбор / остальные как обычно" в set_contractor. Теперь дополнительно
    требуем непустой (не whitespace) текст — реальная строка выбора всегда
    содержит имя контрагента, а не пустую иконку.

    Ждём появления такого CheckBox-а, потому что диалог может прогружаться
    с задержкой.
    """
    def _find():
        for cb in win.descendants(control_type="CheckBox"):
            try:
                text = cb.window_text()
            except Exception:
                continue
            if not text.strip():
                continue
            if text not in _FORM_CHECKBOX_LABELS:
                return cb
        return None

    return _wait_for(
        _find,
        "первая строка CheckBox в диалоге контрагента",
        timeout=timeout,
    )


def _find_text(win, text: str, control_type: str | None = None, exact: bool = True, visible_only: bool = False):
    """
    Первый элемент дерева win, чей window_text() совпадает с text (или
    содержит его, если exact=False).

    visible_only=True — пропускает элементы, которые технически есть в
    дереве, но реально не видны на экране (напр. controls с других вкладок
    диалога настроек — 1С держит их в UIA-дереве даже когда вкладка не
    активна). Без этого фильтра поиск по всему дереву мог одинаково
    ошибочно как найти "Контрагент" не там (на другой вкладке), так и НЕ
    найти видимую строку (если control_type угадан неверно и элемент
    отфильтровался раньше, чем дошло до проверки видимости) — оба случая
    уже встречались на реальных тестах, поэтому здесь ищем ПО ТЕКСТУ без
    привязки к control_type, но с проверкой видимости.
    """
    elements = win.descendants(control_type=control_type) if control_type else win.descendants()
    for el in elements:
        try:
            t = el.window_text()
        except Exception:
            continue
        if not ((t.strip() == text) if exact else (text in t)):
            continue
        if visible_only:
            try:
                if not el.is_visible():
                    continue
            except Exception:
                pass
        return el
    return None


def _wait_click_text(win, text: str, control_type: str | None = None, timeout: float = _DEFAULT_TIMEOUT,
                     exact: bool = True, double: bool = False, verify=None):
    """
    Ждёт появления элемента с текстом text и кликает по нему.

    verify — необязательная проверка "клик реально сработал" (напр. открылся
    нужный диалог). Если передана, клик повторяется/перебирает способы, пока
    она не выполнится (см. click_until). Это важно при виртуальном вводе: у
    кнопок 1С UIA-паттерны часто отрабатывают вхолостую, и без проверки код
    идёт дальше, считая, что нажал.
    """
    el = _wait_for(lambda: _find_text(win, text, control_type=control_type, exact=exact), f"элемент '{text}'", timeout=timeout)
    if double:
        _double_click_input_preserve_cursor(el)
    elif verify is not None:
        click_until(el, verify=verify, what=f"'{text}'")
    else:
        _click_input_preserve_cursor(el)
    return el


# Условие отбора "В группе из списка" — русский и английский варианты.
# ЖИВЬЁМ ПОДТВЕРЖДЕНО на сервере: интерфейс платформы там английский, и
# список условий показывает "In list group" (при том что сама форма настроек
# и тело отчёта остаются русскими — они из конфигурации, а не из платформы).
#
# Сравнение СТРОГОЕ, не подстрокой: "Не в группе из списка" / "Not in list
# group" целиком содержат нужный текст и выбрались бы вместо него — это
# прямо противоположный фильтр.
_IN_LIST_GROUP_OPTIONS = ("в группе из списка", "in list group")

# Остальные подписи, которые рисует сама платформа, — тоже в двух языках.
_IN_REPORT_HEADER_OPTIONS = ("В шапке отчета", "In report header")
_PICK_BUTTON_TITLES = ("Подбор", "Pick")
_SELECT_BUTTON_TITLES = ("Выбрать", "Select")
_SEARCH_PLACEHOLDERS = ("Поиск...", "Search...")


def _click_in_list_group_option(win, timeout: float = 6.0) -> bool:
    """
    Находит в раскрытом списке условий пункт "В группе из списка" (на любом
    из двух языков) и кликает по нему. True — кликнули, False — не нашли.

    Сначала среди ListItem (так этот список размечен обычно), потом среди
    всех элементов — на случай, если конкретный пункт не типизирован.
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        for control_type in ("ListItem", None):
            try:
                items = win.descendants(control_type=control_type) if control_type else win.descendants()
            except Exception:
                continue
            for item in items:
                try:
                    text = (item.window_text() or "").strip().lower()
                except Exception:
                    continue
                if text in _IN_LIST_GROUP_OPTIONS:
                    _click_input_preserve_cursor(item)
                    return True
        time.sleep(0.3)
    return False


def _wait_click_any_text(win, texts, control_type: str | None = None,
                         timeout: float = _DEFAULT_TIMEOUT, exact: bool = True,
                         double: bool = False, verify=None):
    """
    То же, что _wait_click_text, но принимает НЕСКОЛЬКО вариантов текста и
    кликает по первому найденному.

    Нужно из-за смешанного языка интерфейса: на сервере платформа 1С
    английская, а конфигурация русская, поэтому часть подписей на форме —
    русские ("Добавить отбор", "Закрыть и сформировать"), а всё, что рисует
    сама платформа (значения перечислений, стандартные кнопки выбора), —
    английское. Держим оба варианта и не гадаем, какой попадётся.
    """
    deadline = time.time() + timeout
    element = None
    while time.time() < deadline and element is None:
        for text in texts:
            element = _find_text(win, text, control_type=control_type, exact=exact)
            if element is not None:
                break
        if element is None:
            time.sleep(0.3)
    if element is None:
        raise TimeoutError(f"Не дождался ни одного из элементов: {list(texts)}")

    if double:
        _double_click_input_preserve_cursor(element)
    elif verify is not None:
        click_until(element, verify=verify, what=str(texts[0]))
    else:
        _click_input_preserve_cursor(element)
    return element


def _find_in_row(win, row_rect, control_type: str | None = None, text_predicate=None, tolerance: float = 6):
    """
    Ищет элемент В ТОЙ ЖЕ СТРОКЕ по вертикали, что и row_rect (таблица
    отборов — у каждой ячейки строки одна и та же высота/Y, независимо от
    того, что в ячейке: CheckBox, ComboBox, Custom-текст). Опционально
    сужает по control_type и/или text_predicate(text) -> bool.

    Живьём подтверждено (реальная форма "Сводная ведомость расчетов -
    копия 2.0", вкладка Отборы): accessible-текст ячеек НЕ совпадает с
    видимым на экране значением напрямую — 1С добавляет суффикс с именем
    колонки, напр. ячейка "Поле" для контрагента — 'Контрагент Поле' (не
    просто 'Контрагент'), ячейка "Условие" — 'Равно Условие' (плюс
    отдельный ComboBox с текстом 'Условие' в том же прямоугольнике),
    ячейка "★" (избранное) — просто ' ' (один пробел, без текста-состояния
    вообще, независимо от того "★"/"✕"/"-" видно на экране). Поэтому для
    таких ячеек надёжнее искать ПО ПОЗИЦИИ (та же строка) + control_type,
    а не по точному тексту.
    """
    row_mid_y = (row_rect.top + row_rect.bottom) / 2
    elements = win.descendants(control_type=control_type) if control_type else win.descendants()
    for el in elements:
        try:
            rect = el.rectangle()
        except Exception:
            continue
        mid_y = (rect.top + rect.bottom) / 2
        if abs(mid_y - row_mid_y) > tolerance:
            continue
        if text_predicate is not None:
            try:
                if not text_predicate(el.window_text()):
                    continue
            except Exception:
                continue
        return el
    return None


def _contractor_field_visible(win) -> bool:
    """
    True, если в шапке отчёта уже виден подписанный параметр "Контрагент:"
    (появляется, только когда фильтр "Контрагент" в настройках отчёта
    помечен избранным — см. configure_contractor_group_filter). Если да,
    настройку уже когда-то делали и 1С её запомнила — повторять не нужно.

    visible_only=True: живьём подтверждено, что при нескольких открытых
    вкладках отчётов (напр. "Ведомость по денежным средствам - копия" рядом
    со "Сводная ведомость расчетов - копия 2.0") 1С держит элементы НЕактивной
    вкладки в общем UIA-дереве — без этого фильтра поиск может зацепить
    чужую, неактивную вкладку вместо реально открытой сейчас.
    """
    return _find_text(win, "Контрагент:", visible_only=True) is not None


def _edit_value(el) -> str:
    """
    Текст Edit-элемента. ПОДТВЕРЖДЕНО ЖИВЬЁМ (живой дамп поля "Контрагент:"
    после реального выбора через "Подбор"): window_text() у ЭТОГО конкретного
    Edit-поля возвращает '' ВСЕГДА, даже когда на экране виден настоящий
    текст ('Opt-Example-Client-2') — 1С не заполняет accessible Name у этого
    поля. Настоящее значение лежит в UIA ValuePattern:
    legacy_properties()['Value'] / iface_value.CurrentValue — оба живьём
    подтверждены, содержат правильный текст, когда window_text() пуст.
    window_text() пробуем первым (работает для большинства обычных
    Edit-полей и не требует доп. интерфейса), ValuePattern — запасной путь
    именно для полей вроде этого.
    """
    try:
        t = el.window_text()
    except Exception:
        t = ""
    if t.strip():
        return t
    try:
        v = el.iface_value.CurrentValue
        if v:
            return v
    except Exception:
        pass
    try:
        v = el.legacy_properties().get("Value", "")
        if v:
            return v
    except Exception:
        pass
    return ""


def _contractor_field_text(win) -> str:
    """
    Текущий текст поля "Контрагент:" в шапке отчёта — используется как
    проверка по факту после диалога выбора (см. set_contractor): все
    клики могут пройти без единого исключения, а поле в итоге остаться
    пустым (напр. контрагент не нашёлся в справочнике 1С под этим именем
    и код кликнул не туда) — единственный надёжный способ узнать реальный
    результат — прочитать само поле, а не доверять "раз не упало, значит
    сработало".

    ПОДТВЕРЖДЕНО ЖИВЬЁМ: без фильтра видимости эта проверка может найти
    метку "Контрагент:" (и/или Edit-поле рядом с ней) на ДРУГОЙ, неактивной
    вкладке отчёта (в сессии одновременно открыты "Ведомость по денежным
    средствам - копия" и "Сводная ведомость расчетов - копия 2.0" — у обеих
    похожая структура шапки) — тогда результат читается с чужого, всегда
    пустого поля, а не с реально заполненного поля активной вкладки. И
    метку, и Edit-поле теперь ищем только среди видимых элементов.

    Читаем значение через _edit_value(), а не window_text() напрямую —
    см. его докстринг: именно ЭТО поле после выбора через "В группе из
    списка" всегда даёт пустой window_text(), несмотря на видимый на
    экране текст (живьём подтверждённая находка, а не гипотеза).
    """
    label = _find_text(win, "Контрагент:", control_type="CheckBox", exact=True, visible_only=True)
    if label is None:
        return ""
    rect = label.rectangle()
    mid_y = (rect.top + rect.bottom) / 2
    for el in win.descendants(control_type="Edit"):
        try:
            r2 = el.rectangle()
        except Exception:
            continue
        r2_mid_y = (r2.top + r2.bottom) / 2
        if abs(r2_mid_y - mid_y) <= 10 and r2.left >= rect.left:
            try:
                if not el.is_visible():
                    continue
            except Exception:
                pass
            return _edit_value(el)
    return ""


def configure_contractor_group_filter(win) -> bool:
    """
    Одноразовая (идемпотентная) настройка отчёта: включает фильтр
    "Контрагент" в режиме "В группе из списка" и помечает его избранным —
    тогда он появляется в шапке отчёта отдельным полем с кнопкой "...",
    через которую set_contractor() открывает диалог "Подбор". Без этой
    настройки set_contractor() работать не может (использует старый,
    менее надёжный способ ввода — см. _set_contractor_legacy).

    Шаги — по описанию и скриншотам реальной формы 1С пользователем:
      1. Кнопка "Настройки...".
      2. Переключить вид на "Расширенный" (radio-кнопка в группе "Вид:"),
         если форма открылась в простом виде — иначе вкладки/раздела
         отборов может не быть видно.
      3. Если строки "Контрагент" среди отборов ещё нет — "Добавить
         отбор", затем дважды кликнуть "Контрагент" в списке полей.
      4. Сменить условие строки "Контрагент" с "Равно" на "В группе из
         списка".
      5. Пометить строку "Контрагент" избранной (звёздочка) — тогда поле
         появится в шапке отчёта.
      6. "Закрыть и сформировать".

    ВАЖНО: шаги 4 и 5 (смена условия и звёздочка-избранное) — самое
    неопределённое место в этой функции: точный механизм клика (обычный
    клик по ячейке условия -> выпадающий список? двойной клик -> меню?
    отдельная колонка со звездой, куда нужно кликнуть?) не был подтверждён
    напрямую на форме ИМЕННО этого отчёта ("Сводная ведомость расчетов -
    копия 2.0" — при проверке в реальном сеансе оказался открыт другой
    отчёт, "Ведомость по денежным средствам - копия", с похожей, но не
    идентичной структурой настроек). Если эти два шага не сработают —
    функция не падает, а печатает предупреждение и продолжает: тогда их
    нужно один раз сделать вручную (как на скриншотах), после чего
    set_contractor() будет работать сам на уже настроенном фильтре.

    Возвращает True, если поле "Контрагент:" в итоге видно в шапке отчёта
    (настройка уже была или применилась сейчас), иначе False.
    """
    if _contractor_field_visible(win):
        print("  [1С] Фильтр 'Контрагент' уже настроен (виден в шапке отчёта) — пропускаю настройку.")
        return True

    print("  [1С] Настраиваю фильтр 'Контрагент' (Настройки -> Расширенный -> Отборы)...")
    # Эти два окна настройка открывает САМА — пока она идёт, сторож их не
    # трогает. В любое другое время такое окно висит зря и снимается.
    with expected_dialogs("Настройки отчета", "Выбор поля"):
        return _configure_contractor_group_filter_inner(win)


def _configure_contractor_group_filter_inner(win) -> bool:
    try:
        vi.focus(win)
        time.sleep(0.2)

        # Если диалог "Настройки" уже открыт (остался от предыдущего
        # незавершённого запуска — именно так и произошло на реальном
        # тесте) — НЕ нажимаем "Настройки..." заново: под открытым
        # модальным диалогом эта кнопка недоступна/перекрыта, и попытка её
        # найти впустую съедает весь таймаут ожидания (это и была причина
        # "долго делает" — не сама пауза после добавления поля, а лишний
        # клик по уже неактивной кнопке на предыдущем шаге).
        settings_already_open = _find_text(win, "Закрыть и сформировать", control_type="Button") is not None

        if settings_already_open:
            print("  [1С] Диалог настроек уже открыт (остался с прошлого раза) — использую как есть.")
        else:
            # verify: диалог настроек считается открытым, когда появилась
            # кнопка "Закрыть и сформировать" (она есть только в нём).
            _wait_click_text(
                win, "Настройки...", control_type="Button", timeout=8,
                verify=lambda: _find_text(win, "Закрыть и сформировать", control_type="Button") is not None,
            )
            time.sleep(1)

            # Простой вид часто не показывает раздел отборов — переключаемся
            # на "Расширенный" (radio-кнопка), если она есть и ещё не выбрана.
            try:
                _wait_click_text(win, "Расширенный", control_type="RadioButton", timeout=2)
                time.sleep(0.5)
                print("  [1С] Переключил вид настроек на 'Расширенный'.")
            except TimeoutError:
                print("  [1С] Radio-кнопка 'Расширенный' не найдена за 2 сек — возможно, уже в расширенном виде.")

            # Вкладка "Отборы", если она отдельная (не на всех отчётах есть —
            # у некоторых отборы показаны сразу в основном списке настроек).
            try:
                _wait_click_text(win, "Отборы", control_type="TabItem", timeout=2)
                time.sleep(0.5)
            except TimeoutError:
                print("  [1С] Отдельная вкладка 'Отборы' не найдена — ищу строку 'Контрагент' в текущем виде.")

        # Если строки "Контрагент" среди ОТБОРОВ ещё нет — добавляем.
        # ПОДТВЕРЖДЕНО ЖИВЬЁМ (реальная форма): строка рендерится как
        # CheckBox, но его accessible-текст — "Контрагент Поле" (с
        # суффиксом-именем колонки), НЕ просто "Контрагент" — точное
        # совпадение поэтому никогда не срабатывало на предыдущих
        # попытках. Ищем подстрокой ("Контрагент" содержится в тексте),
        # с control_type="CheckBox" (сужает от ложных совпадений типа
        # "Контрагент Заголовок" — той же строки, но другой колонки).
        if _find_text(win, "Контрагент", control_type="CheckBox", exact=False) is None:
            print("  [1С] Строки 'Контрагент' в отборах нет — добавляю через 'Добавить отбор'...")
            _wait_click_text(win, "Добавить отбор", control_type="Button", timeout=5)
            time.sleep(1)
            # Диалог "Выбор поля отчета" — дерево полей, где нужно ОДИН клик
            # по "Контрагент" (выделяет строку, подсвечивает оранжевым), а
            # затем отдельная кнопка "Выбрать" вверху слева — двойной клик
            # по строке дерева либо ничего не делает, либо разворачивает
            # узел вместо подтверждения выбора (см. скриншот пользователя).
            print("  [1С] Кликаю 'Контрагент' в списке полей и жму 'Выбрать'...")
            _wait_click_text(win, "Контрагент", timeout=8, double=False, exact=False)
            time.sleep(0.5)
            _wait_click_any_text(win, _SELECT_BUTTON_TITLES, control_type="Button", timeout=5)
            time.sleep(1)
            if _find_text(win, "Контрагент", control_type="CheckBox", exact=False) is None:
                print("  [1С] ⚠ После 'Выбрать' строка 'Контрагент' в отборах всё ещё не появилась.")
        else:
            print("  [1С] Строка 'Контрагент' уже есть среди отборов.")

        contractor_row = _wait_for(
            lambda: _find_text(win, "Контрагент", control_type="CheckBox", exact=False),
            "строка 'Контрагент' в отборах",
            timeout=10,
        )
        row_rect = contractor_row.rectangle()
        print(f"  [1С] Строка 'Контрагент' найдена: rect={row_rect}")

        # ─── Смена условия "Равно" -> "В группе из списка" ───
        # ПОДТВЕРЖДЕНО ЖИВЬЁМ (полная цепочка воспроизведена и проверена):
        # ячейка условия — Custom-элемент с текстом вида "Равно Условие"
        # (значение + суффикс-имя колонки). ДВОЙНОЙ клик по этой ячейке
        # переводит её в режим редактирования — только тогда рядом
        # появляются ComboBox и выпадающий список вариантов (до двойного
        # клика в окне вообще нет ни одного ComboBox — он не существует,
        # пока строка не "открыта" на редактирование). Одиночный клик
        # ничего не даёт. Сами варианты — ListItem с ЧИСТЫМ текстом без
        # суффиксов ("В группе из списка", "Равно", "Не равно" и т.д.).
        try:
            changed = False
            for attempt in range(1, 3):  # до 2 попыток — двойной клик иногда не сразу раскрывает список
                condition_cell = _find_in_row(
                    win, row_rect, control_type="Custom",
                    text_predicate=lambda t: any(
                        part in t for part in
                        ("Равно", "Условие", "группе", "Equal", "Condition", "group")
                    ),
                )
                if condition_cell is None:
                    print("  [1С] ⚠ Не нашёл ячейку условия в строке 'Контрагент' — потребуется сменить вручную.")
                    break
                _double_click_input_preserve_cursor(condition_cell)
                time.sleep(1.2)  # раскрытие списка на реальном прогоне занимало больше 0.8с
                if _click_in_list_group_option(win, timeout=6):
                    time.sleep(0.5)
                    changed = True
                    break
                _send_keys(win, "{ESC}")
                time.sleep(0.5)
                print(f"  [1С] Попытка {attempt}/2: список условий не раскрылся или пункт не найден — пробую снова...")
            if changed:
                print("  [1С] Условие 'Контрагент' изменено на 'В группе из списка'.")
            else:
                print("  [1С] ⚠ Не удалось сменить условие после 2 попыток — смени 'Контрагент' на 'В группе из списка' вручную один раз.")
        except Exception as e:
            print(f"  [1С] ⚠ Не удалось сменить условие 'Контрагент' на 'В группе из списка' автоматически: {e}. Смени вручную один раз.")

        # ─── Избранное (звёздочка) ───
        # ПОДТВЕРЖДЕНО ЖИВЬЁМ (полная цепочка воспроизведена и проверена
        # скриншотом — звезда реально встала, поле "Контрагент:" появилось
        # в шапке отчёта после "Закрыть и сформировать"): ячейка "★" не
        # несёт различающего текста — у ЛЮБОЙ строки (и избранной, и нет)
        # window_text() там просто " " (один пробел), состояние видно
        # только визуально ("★"/"✕"/"-"). ДВОЙНОЙ клик по этой ячейке (не
        # одиночный — одиночный клик визуально ничего не меняет) открывает
        # контекстное меню с вариантами: "В шапке отчета" (нужный нам —
        # обычно уже первый/выделенный по умолчанию), "Только флажок в
        # шапке отчета", "В настройках отчета", "Только флажок в
        # настройках отчета", "Не показывать". Кликаем по пункту меню
        # "В шапке отчета" (просто текстовый элемент где-то в дереве —
        # не привязываемся к control_type, он не подтверждён однозначно).
        try:
            star_cell = _find_in_row(
                win, row_rect, control_type="Custom", text_predicate=lambda t: t == " "
            )
            if star_cell is not None:
                starred = False
                for attempt in range(1, 3):  # до 2 попыток — двойной клик иногда не сразу раскрывает меню
                    _double_click_input_preserve_cursor(star_cell)
                    time.sleep(1.0)
                    try:
                        _wait_click_any_text(win, _IN_REPORT_HEADER_OPTIONS, timeout=5, exact=True)
                        time.sleep(0.5)
                        starred = True
                        break
                    except TimeoutError:
                        try:
                            _wait_click_any_text(win, _IN_REPORT_HEADER_OPTIONS, timeout=2, exact=False)
                            time.sleep(0.5)
                            starred = True
                            break
                        except TimeoutError:
                            _send_keys(win, "{ESC}")
                            time.sleep(0.5)
                            print(f"  [1С] Попытка {attempt}/2: меню избранного не раскрылось или пункт не найден — пробую снова...")
                if starred:
                    print("  [1С] Поставил избранное ('В шапке отчета') для строки 'Контрагент'.")
                else:
                    print("  [1С] ⚠ Не удалось поставить избранное после 2 попыток — поставь вручную один раз.")
            else:
                print("  [1С] ⚠ Не нашёл ячейку избранного в строке 'Контрагент' — поставь звёздочку вручную один раз.")
        except Exception as e:
            print(f"  [1С] ⚠ Не удалось поставить избранное для 'Контрагент' автоматически: {e}. Поставь вручную один раз.")

        _wait_click_text(win, "Закрыть и сформировать", control_type="Button", timeout=5)
        time.sleep(2)

        if _contractor_field_visible(win):
            print("  [1С] ✓ Фильтр 'Контрагент' настроен — поле появилось в шапке отчёта.")
            return True

        print(
            "  [1С] ⚠ Поле 'Контрагент:' в шапке отчёта так и не появилось. "
            "Похоже, не удалось поставить 'избранное' (звёздочку) автоматически — "
            "сделай это один раз вручную (Настройки -> Расширенный -> Отборы -> "
            "строка 'Контрагент' -> звёздочка), дальше скрипт справится сам."
        )
        return False
    except Exception as e:
        print(f"  [1С] ⚠ Не удалось автоматически настроить фильтр 'Контрагент': {e}")
        print(
            "  [1С] Настрой один раз вручную (Настройки -> Расширенный -> Отборы -> "
            "Добавить отбор -> Контрагент -> В группе из списка -> избранное -> "
            "Закрыть и сформировать), дальше скрипт справится сам."
        )
        return False


def _set_contractor_legacy(win, name: str) -> None:
    """
    Старый способ ввода контрагента — F4 -> заменить первую строку
    свободного списка. Работал, пока фильтр "Контрагент" был в режиме
    "Равно" (свободный текст, не строгая ссылка на справочник). Оставлен
    как запасной вариант на случай, если фильтр ещё не переведён в режим
    "В группе из списка" (см. configure_contractor_group_filter) — тогда
    основной способ (диалог "Подбор" через кнопку "...", см. set_contractor)
    не находит своих элементов и откатывается сюда.

    Вставка: двойной левый клик по строке (переводит её в режим редактирования
    и сразу выделяет весь текст) -> Home -> Shift+End (выделить всё) ->
    Shift+Insert (вставка поверх выделения). Правый клик -> "Вставить" из
    контекстного меню не используется — он ненадёжен (см. комментарий ниже).
    """
    vi.focus(win)
    time.sleep(0.5)

    # --- Находим поле контрагента с ожиданием загрузки формы ---
    print("  Ищу поле ввода контрагента...")
    edit = _contractor_edit(win)
    print("  Поле найдено, кликаю и жму F4...")

    _click_input_preserve_cursor(edit)
    time.sleep(0.5)
    _send_keys(edit, "{F4}")

    # --- Ждём реального появления диалога ---
    # Кнопка "Подбор" есть только в диалоге выбора контрагента.
    # Увеличиваем таймаут — F4 не всегда успевает открыть окно быстро.
    print("  Жду открытия диалога (кнопка 'Подбор')...")
    try:
        podbor = win.child_window(title="Подбор", control_type="Button")
        podbor.wait("exists visible", timeout=_DEFAULT_TIMEOUT)
    except Exception:
        # Если диалог не открылся или кнопка не найдена — вводим вручную.
        print("  Диалог не открылся, вставляю контрагента вручную...")
        _set_clipboard_text(name)
        _click_input_preserve_cursor(edit)
        time.sleep(0.3)
        _send_keys(None, "{HOME}")
        time.sleep(0.1)
        _send_keys(None, "+{END}")
        time.sleep(0.1)
        _send_keys(None, "+{INSERT}")
        time.sleep(0.5)
        _send_keys(None, "{ENTER}")
        return

    print("  Диалог открыт, вставляю контрагента...")

    _set_clipboard_text(name)

    row = _dialog_first_row(win)
    print(f"  Найдена строка: '{row.window_text()}'")

    # ─── Шаг 1: Двойной клик по строке ───
    # Двойной левый клик переводит строку в режим редактирования — появляется
    # inline Edit-поле.
    _double_click_input_preserve_cursor(row)
    time.sleep(1.0)

    # ─── Шаг 2: Выделяем ВЕСЬ текст в строке ───
    _send_keys(None, "{HOME}")
    time.sleep(0.2)
    _send_keys(None, "+{END}")      # Shift+End — выделить всё до конца строки
    time.sleep(0.3)
    print("  Текст выделен (Home + Shift+End)")

    # ─── Шаг 3: Вставка — заменяет выделенный текст ───
    # Правый клик → «Вставить» из контекстного меню убран: он не бросает
    # исключение даже когда реально ничего не подставляет (клик по пункту
    # меню «успешен» сам по себе), поэтому раньше маскировал случаи, когда
    # контрагент вставлялся неверно/пусто — отчёт формировался по плохому
    # фильтру → 0 строк → нечего писать в Sheets. Плюс это лишние ~1-6 сек
    # ожидания на каждого контрагента. Shift+Insert — единственный способ,
    # который проверенно работает стабильно.
    _send_keys(None, "+{INSERT}")    # Shift+Insert — вставка поверх выделения
    print("  Вставлено через Shift+Insert.")

    time.sleep(0.5)

    # Подтверждаем строку через Enter — без этого для некоторых контрагентов
    # (напр. у некоторых клиентов) строка остаётся в режиме редактирования и диалог не
    # реагирует на клик «ОК» (зависает открытым).
    _send_keys(None, "{ENTER}")
    print("  Нажал Enter для подтверждения контрагента.")

    time.sleep(0.5)

    # Закрываем подсказку автоподбора (если появилась)
    _send_keys(win, "{ESC}")
    time.sleep(0.5)

    _click_ok_button(_value_dialog() or win)
    time.sleep(0.5)


def _contractor_ellipsis_points(win, timeout: float = _DEFAULT_TIMEOUT) -> list:
    """
    Координаты-кандидаты [(x, y), ...] кнопки "..." рядом с полем "Контрагент:" в шапке
    отчёта. ПОДТВЕРЖДЕНО ЖИВЬЁМ: это НЕ отдельный UIA-элемент вообще — ни
    Button, ни Hyperlink, просто нарисованные пиксели внутри поля ввода
    (win.descendants() никогда её не находит, независимо от control_type
    — проверено точечным дампом всей области вокруг поля, там просто
    ничего нет). Единственный надёжный способ — клик по вычисленным
    координатам, а не по найденному элементу.

    Смещение +480px по X от левого края чекбокса "Контрагент:" (та же Y,
    что и у чекбокса) — измерено и подтверждено прямым кликом на реальной
    форме отчёта "Сводная ведомость расчетов - копия 2.0".

    visible_only=True — самое важное место для этого фильтра во всём
    модуле: если тут найдётся метка "Контрагент:" с ДРУГОЙ, неактивной
    вкладки (напр. "Ведомость по денежным средствам - копия", открытой
    рядом), клик по "..." попадёт на РЕАЛЬНЫЕ экранные координаты чужой,
    невидимой сейчас метки — то есть куда-то в пустоту текущей активной
    вкладки — и весь последующий диалог откроется не пойми где/что.
    """
    label = _wait_for(
        lambda: _find_text(win, "Контрагент:", control_type="CheckBox", exact=True, visible_only=True),
        "метка 'Контрагент:'",
        timeout=timeout,
    )
    rect = label.rectangle()
    y = (rect.top + rect.bottom) // 2

    # Кнопка привязана к ПРАВОМУ КРАЮ поля ввода, а не к фиксированному
    # отступу от метки. ЖИВЬЁМ ПОДТВЕРЖДЕНО на сервере: там окно шире
    # (1920), поле "Контрагент:" растянуто, и прежнее измеренное смещение
    # +480px попадало в середину пустого поля — диалог не открывался вовсе.
    # Ширина поля меняется вместе с окном, положение его правого края — нет.
    field = (_find_in_row(win, rect, control_type="Edit")
             or _find_in_row(win, rect, control_type="ComboBox"))

    points = []
    if field is not None:
        try:
            fr = field.rectangle()
            field_y = (fr.top + fr.bottom) // 2
            # "..." рисуется то внутри поля у самого края, то сразу за ним
            # (зависит от того, показан ли рядом крестик очистки) — пробуем
            # оба положения по очереди, проверяя результат.
            points += [(fr.right - 14, field_y), (fr.right + 14, field_y), (fr.right + 40, field_y)]
        except Exception:
            pass

    # Прежнее измеренное смещение — последним, как запасной вариант для
    # окна привычной ширины, если поле в строке почему-то не нашлось.
    points.append((rect.left + 480, y))
    return points


def _click_screen_coords(x: int, y: int, double: bool = False) -> None:
    """
    Клик по АБСОЛЮТНЫМ экранным координатам — там, где нет подходящего
    UIA-элемента: кнопка "..." рядом с "Контрагент:" (см.
    _contractor_ellipsis_points, диалог "Подбор") и строки грида отчёта
    (_row1_anchor в usluga_report).

    Есть активный рабочий стол -> НАСТОЯЩИЙ клик мышью с возвратом курсора
    на место. Это исходное, годами работающее поведение этого проекта;
    посылаемые сообщения по этим местам 1С не срабатывают (область
    отрисована самой 1С, отдельного hwnd под ней нет).

    Нет рабочего стола (сервер с отключённой сессией) -> посланное
    WM_LBUTTONDOWN/UP в hwnd под этой точкой.
    """
    if vi.ALLOW_REAL_INPUT:
        original_pos = win32api.GetCursorPos()
        try:
            win32api.SetCursorPos((x, y))
            time.sleep(0.1)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            time.sleep(0.05)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            if double:
                time.sleep(0.05)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                time.sleep(0.05)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            time.sleep(0.1)
        finally:
            win32api.SetCursorPos(original_pos)
        return

    vi.click_at(_default_hwnd(), x, y, double=double, settle=0.15)


def _search_in_progress(win) -> bool:
    return any(_find_text(win, text) is not None for text in _SEARCH_PLACEHOLDERS)


def _close_window_by_title(title: str, timeout: float = 5.0) -> bool:
    """
    Закрывает окно с заданным ЗАГОЛОВКОМ (напр. "Контрагенты" — каталог
    "Подбора") настоящим Windows-сообщением WM_CLOSE по его hwnd —
    НЕ кликом по вычисленным экранным координатам "✕".

    Раньше пытались найти "×"/"X"/"x" как текстовый элемент где-то в
    дереве — ПОДТВЕРЖДЕНО ЖИВЬЁМ, что это ловит не то: рядом со строкой
    поиска каталога есть СВОЙ "x" (очистить поиск), который находится
    раньше и кликается вместо настоящего закрытия окна. Тогда перешли на
    клик по координатам (right-15, top+15) — но это ТОЖЕ ловило не то:
    живьём подтверждено на контрагенте 'OPT-Example-Client' — окно
    "Контрагенты" осталось открытым, а клик, судя по всему, попал куда-то
    в другое место экрана (если хоть один из хендлов с таким заголовком —
    невидимый/устаревший remnant с другими координатами, или z-order в
    этот момент не тот, что предполагали координаты) — вместо надёжного
    закрытия получался непредсказуемый клик мимо. WM_CLOSE не зависит ни
    от пикселей, ни от z-order, ни от того, что сейчас поверх чего — это
    прямая команда конкретному hwnd, а не "клик туда, где по расчётам
    должен быть крестик".

    Возвращает True, если окно с таким заголовком нашлось (независимо от
    того, закрылось ли оно после WM_CLOSE — вызывающий код сам проверяет
    результат по своим признакам).

    Закрывает ВСЕ найденные окна с этим заголовком, а не только первое —
    от повторных тестовых прогонов на реальной машине остаются висящие
    дубли того же диалога ("Контрагенты").
    """
    deadline = time.time() + timeout
    hwnds = []
    while time.time() < deadline:
        found = []

        def _cb(h, _):
            if win32gui.IsWindowVisible(h) and win32gui.GetWindowText(h).strip() == title:
                found.append(h)

        win32gui.EnumWindows(_cb, None)
        if found:
            hwnds = found
            break
        time.sleep(0.3)

    if not hwnds:
        return False

    for hwnd in hwnds:
        try:
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        except Exception:
            continue
        time.sleep(0.3)
    return True


_CATALOG_WINDOW_TITLE = "Контрагенты"
_VALUE_DIALOG_TITLE = "Контрагент"


def _top_window_by_title(title: str):
    """
    Верхнеуровневое видимое окно с ТОЧНО таким заголовком, как обёртка
    pywinauto, или None. Совпадение точное, поэтому "Контрагент" (диалог
    списка значений фильтра) не путается с "Контрагенты" (каталог подбора).
    """
    found = []

    def _cb(hwnd, _):
        try:
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd).strip() == title:
                found.append(hwnd)
        except Exception:
            pass

    try:
        win32gui.EnumWindows(_cb, None)
    except Exception:
        return None
    if not found:
        return None
    try:
        _, pid = win32process.GetWindowThreadProcessId(found[-1])
        return Application(backend="uia").connect(process=pid).window(handle=found[-1])
    except Exception:
        return None


def _value_dialog():
    """
    Диалог списка значений фильтра "Контрагент" (тот, где кнопки "Подбор" и
    "ОК") — отдельное окно, а не часть окна отчёта.

    Нужен именно он, а не win: в окне отчёта живут чекбоксы "Период:" и
    "Контрагент:", включающие сами фильтры. Снятие отметок перед "Подбором"
    шло по win.descendants() и гасило ИХ — то есть выключало фильтр
    "Контрагент" прямо перед выбором контрагента. Этот путь выполняется
    только для ПЕРВОГО контрагента (дальше идёт перезапись на месте), что и
    давало "первый всегда с ошибкой, со второго всё хорошо".
    """
    return _top_window_by_title(_VALUE_DIALOG_TITLE)


def _catalog_window():
    """
    Окно каталога контрагентов ("Контрагенты"), которое открывает "Подбор".

    Отдельное ВЕРХНЕУРОВНЕВОЕ окно, а не часть окна отчёта. Это важно: и
    проверка "Подбор сработал", и поиск строки поиска раньше смотрели
    win.descendants(control_type="Edit") — то есть В ОКНЕ ОТЧЁТА, где Edit
    есть всегда (поля периода). Проверка поэтому проходила мгновенно и
    всегда, даже когда каталог не открылся вовсе: клик по "Подбор"
    считался успешным без единого повтора, и первый контрагент оставался
    необработанным.
    """
    return _top_window_by_title(_CATALOG_WINDOW_TITLE)


def _click_ok_button(scope, timeout: float = 5.0) -> None:
    """
    Жмёт "ОК" в диалоге. Заголовок проверяем и кириллицей, и латиницей: на
    англоязычной платформе кнопка называется латинским "OK", а для поиска
    по тексту "ОК" и "OK" — разные строки, хотя на экране неотличимы.
    """
    def _find():
        for title in ("ОК", "OK"):
            button = scope.child_window(title=title, control_type="Button")
            try:
                if button.exists():
                    return button
            except Exception:
                continue
        return None

    _click_input_preserve_cursor(_wait_for(_find, "кнопка 'ОК'", timeout=timeout))


def _set_contractor_via_podbor(win, name: str) -> None:
    """
    Полный путь через "Подбор" (каталог контрагентов) — нужен, когда в
    списке выбора ("Контрагент:" -> "...") ещё нет ни одной строки, т.е.
    для ПЕРВОГО контрагента синхронизации (см. set_contractor). Кнопка
    "..." должна уже быть нажата вызывающим кодом до этого вызова.
    """
    # Диалог списка значений — отдельное окно; всё дальнейшее делаем В НЁМ.
    dialog = _value_dialog() or win

    # Снимаем старую отметку, если она есть, ДО "Подбора" — иначе "В
    # группе из списка" накопит НЕСКОЛЬКО контрагентов сразу вместо
    # замены на текущего.
    #
    # ТОЛЬКО внутри диалога: в окне отчёта чекбоксы "Период:" и
    # "Контрагент:" включают сами фильтры, и обход по win.descendants()
    # снимал галочку с фильтра "Контрагент" прямо перед выбором
    # контрагента. Эта ветка выполняется лишь для ПЕРВОГО контрагента —
    # отсюда и было "первый с ошибкой, со второго нормально".
    if dialog is not win:
        print("  Снимаю прошлые отметки (если есть)...")
        for cb in dialog.descendants(control_type="CheckBox"):
            try:
                if cb.get_toggle_state() == 1:
                    _click_input_preserve_cursor(cb)
                    time.sleep(0.2)
            except Exception:
                continue
    else:
        print("  ⚠ Диалог 'Контрагент' как окно не нашёлся — прошлые отметки не трогаю "
              "(снимать их по всему окну отчёта нельзя: там чекбоксы самих фильтров).")

    print("  Открываю 'Подбор'...")
    # "Подбор" в этом диалоге может быть не Button, а Hyperlink (так было
    # в похожем диалоге другого отчёта, живьём подтверждено) — не
    # привязываемся к control_type, ищем просто по тексту.
    # verify: "Подбор" сработал, когда открылось окно каталога — его видно
    # по появившейся строке поиска (Edit). Без проверки при виртуальном
    # вводе клик мог "пройти" вхолостую, и дальше код искал строку поиска в
    # так и не открывшемся каталоге.
    _wait_click_any_text(
        dialog, _PICK_BUTTON_TITLES, timeout=8, exact=False,
        verify=lambda: _catalog_window() is not None,
    )

    catalog = _wait_for(_catalog_window, "окно каталога 'Контрагенты'", timeout=15)

    print(f"  Ищу '{name}' в каталоге контрагентов...")
    # Строку поиска берём ИЗ ОКНА КАТАЛОГА. В окне отчёта первый попавшийся
    # Edit — это поле периода, и раньше имя контрагента уезжало именно туда.
    search_box = _wait_for(
        lambda: next(iter(catalog.descendants(control_type="Edit")), None),
        "строка поиска в каталоге контрагентов",
        timeout=8,
    )
    _click_input_preserve_cursor(search_box)
    time.sleep(0.2)
    _send_keys(None, "^a")
    time.sleep(0.1)
    _set_clipboard_text(name)
    _send_keys(None, "+{INSERT}")
    time.sleep(0.3)
    _send_keys(None, "{ENTER}")

    _wait_for(lambda: True if not _search_in_progress(win) else None, "окончание поиска ('Поиск...' исчез)", timeout=10)
    time.sleep(0.3)

    print(f"  Выбираю '{name}' из результатов и жму 'Выбрать'...")
    # ПОДТВЕРЖДЕНО ЖИВЬЁМ: найденная строка результата поиска — control_type
    # "Text", а не "DataItem" (это была причина зависания именно тут).
    # Точную позиционную фильтрацию (относительно заголовка "Наименование")
    # пробовали добавить и это привело к обратному эффекту — на другом
    # прогоне отбросило РЕАЛЬНО существующую строку (видимо, из-за
    # нескольких одновременно открытых окон отчёта "Наименование" нашлось
    # не в том окне). Позиционную фильтрацию НЕ трогаем, но добавляем более
    # безопасный фильтр — is_visible(): элементы скрытых/неактивных вкладок
    # 1С держит в дереве даже когда они не отрисованы, а видимость (в
    # отличие от позиции относительно чужого заголовка) не зависит от
    # какого чужого окна — это и было заподозренной причиной пустого поля
    # у 'Opt-Example-Client-2' (клик попал в невидимый повтор из другой
    # вкладки/окна). Факт "всё равно кликнули не туда" всё равно ловит
    # проверка поля после "ОК" ниже (_contractor_field_text).
    def _find_row():
        # Сначала в окне каталога, потом в окне отчёта: каталог — отдельное
        # окно, и искать строку результата надо прежде всего там.
        for source in (catalog, win):
            if source is None:
                continue
            try:
                elements = source.descendants(control_type="Text")
            except Exception:
                continue
            for el in elements:
                try:
                    if not _names_match(name, el.window_text()):
                        continue
                    if not el.is_visible():
                        continue
                except Exception:
                    continue
                return el
        return None

    row = _wait_for(_find_row, f"строка '{name}' в результатах поиска", timeout=10)
    _click_input_preserve_cursor(row)
    time.sleep(0.3)
    try:
        _wait_click_any_text(catalog, _SELECT_BUTTON_TITLES, control_type="Button", timeout=5)
    except TimeoutError:
        _wait_click_any_text(win, _SELECT_BUTTON_TITLES, control_type="Button", timeout=5)
    time.sleep(1)

    # ПОДТВЕРЖДЕНО ЖИВЬЁМ: "Выбрать" добавляет элемент (всплывает "Элемент
    # добавлен в ..."), но НЕ закрывает окно каталога "Контрагенты" само.
    # Искать "×"/"X" как текстовый элемент НЕ работает — рядом со строкой
    # поиска каталога есть СВОЙ "x" (очистить поиск), он находится первым
    # и кликается вместо настоящего закрытия окна (окно не закрывается).
    # Закрываем именно окно "Контрагенты" по его собственной кнопке "✕" в
    # заголовке (вычисленной от границ окна, см. _close_window_by_title) —
    # если это конкретное окно не нашлось, Escape как запасной вариант.
    if not _close_window_by_title("Контрагенты", timeout=3):
        _send_keys(win, "{ESC}")
    time.sleep(0.8)

    _click_ok_button(dialog)
    time.sleep(0.5)


def _set_contractor_in_place(win, name: str) -> bool:
    """
    "Как обычно" — перезаписывает текст УЖЕ существующей строки в списке
    выбора (открытом по "..."), минуя "Подбор": двойной клик по строке ->
    Home -> Shift+End -> Shift+Insert (name) -> Enter -> Escape -> "ОК".
    Тот же приём, что и в _set_contractor_legacy. Нужен для ВТОРОГО и
    последующих контрагентов синхронизации — после первого через "Подбор"
    в списке уже есть строка, которую быстрее переписать на месте, чем
    открывать каталог заново.

    Возвращает False (без исключения), если строка для перезаписи не
    нашлась за таймаут — вызывающий код тогда сам откатывается на
    _set_contractor_via_podbor. Собственно совпадение итогового значения
    поля "Контрагент:" с ожидаемым проверяет уже вызывающий код.

    ЖИВЬЁМ ЗАМЕЧЕНО: на одном из контрагентов ("OPT-Example-Client") список
    выбора после клика "..." иногда прогружается дольше 3 сек (тот же
    класс задержки, что и у других диалогов этого отчёта) — со старым
    таймаутом 3 сек это ложно читалось как "строки нет" и откатывалось на
    "Подбор", хотя строка была. Увеличен до 6 сек.
    """
    try:
        row = _dialog_first_row(win, timeout=6)
    except TimeoutError:
        return False

    print(f"  Строка уже есть ('{row.window_text()}') — переписываю на месте...")
    _double_click_input_preserve_cursor(row)
    time.sleep(1.0)
    _send_keys(None, "{HOME}")
    time.sleep(0.2)
    _send_keys(None, "+{END}")
    time.sleep(0.3)
    _set_clipboard_text(name)
    _send_keys(None, "+{INSERT}")
    time.sleep(0.5)
    _send_keys(None, "{ENTER}")
    time.sleep(0.5)
    _send_keys(win, "{ESC}")
    time.sleep(0.5)

    _click_ok_button(_value_dialog() or win)
    time.sleep(0.5)
    return True


def _contractor_list_is_open(win) -> bool:
    """
    Диалог списка значений фильтра "Контрагент" открыт? Признак — ссылка
    "Подбор" ("Pick"), она есть в этом диалоге и в первый раз, и в
    последующие. Проверка по СОДЕРЖИМОМУ, а не по тому, что клик "прошёл":
    при виртуальном вводе клик по нарисованной кнопке легко уходит мимо,
    ничего не ломая, и дальше код работает вслепую.
    """
    return any(
        _find_text(win, text, exact=False) is not None for text in _PICK_BUTTON_TITLES
    )


def _open_contractor_list(win, points) -> bool:
    """
    Открывает список фильтра "Контрагент", пробуя точки-кандидаты кнопки
    "..." по очереди (см. _contractor_ellipsis_points) и проверяя после
    каждой, что диалог реально открылся. True — открылся.

    Если уже открыт (остался с прошлого контрагента) — ничего не делаем.
    """
    if _contractor_list_is_open(win):
        return True
    for index, (x, y) in enumerate(points, start=1):
        _click_screen_coords(x, y)
        time.sleep(1)
        if _contractor_list_is_open(win):
            if index > 1:
                print(f"    (кнопка '...' сработала с {index}-й точки)")
            return True
    return False


def set_contractor(win, name: str) -> None:
    """
    Основной способ выставить контрагента: фильтр "Контрагент" настроен
    как "В группе из списка" (см. configure_contractor_group_filter).

    По явному уточнению пользователя, полный "Подбор" (поиск по каталогу)
    нужен только ПЕРВЫЙ раз за синхронизацию — когда список выбора совсем
    пуст и нечего редактировать. Начиная со второго контрагента там уже
    есть строка от предыдущего вызова — её просто ПЕРЕЗАПИСЫВАЕМ на месте
    (как обычно, см. _set_contractor_in_place), а не открываем "Подбор"
    заново.

    "Первый раз или не первый" определяем НЕ через поиск строки в списке
    (так и было раньше — ловило иконку тулбара диалога вместо настоящей
    строки, из-за чего "Подбор" не нажимался вообще ни для одного
    контрагента), а через текущий текст самого поля "Контрагент:" в
    шапке отчёта (_contractor_field_text) ДО открытия списка — пусто
    значит первый раз, что-то есть значит не первый. Это то же самое
    поле, которое ниже служит проверкой результата, так что источник
    истины один и тот же в обеих проверках.

    Если "в списке уже есть строка" определили неверно (перезапись на
    месте не нашла строку или дала не тот результат) — откатывается на
    полный "Подбор", а не падает сразу: не тот же самый диалог заново
    открывать не нужно, он уже открыт.

    Если поле "Контрагент:" / кнопка "..." не найдены за timeout (фильтр
    ещё не переведён в "В группе из списка") — откатывается на
    _set_contractor_legacy(win, name) (старый способ через F4), чтобы не
    падать всей синхронизацией из-за не до конца настроенного фильтра.
    """
    vi.focus(win)
    time.sleep(0.3)

    try:
        ellipsis_points = _contractor_ellipsis_points(win, timeout=5)
    except TimeoutError:
        print("  [1С] Поле 'Контрагент:' не найдено — использую старый способ ввода (F4).")
        _set_contractor_legacy(win, name)
        return

    previous = _contractor_field_text(win)
    is_first_time = not previous.strip()

    print("  Открываю список 'Контрагент' (...)...")
    if not _open_contractor_list(win, ellipsis_points):
        print("  [1С] ⚠ Список 'Контрагент' не открылся ни по одной из точек — "
              "использую старый способ ввода (F4).")
        _set_contractor_legacy(win, name)
        return

    used_podbor = False
    if is_first_time:
        _set_contractor_via_podbor(win, name)
        used_podbor = True
    else:
        if not _set_contractor_in_place(win, name):
            print("  [1С] ⚠ Строка для перезаписи на месте не нашлась — иду через 'Подбор'.")
            _set_contractor_via_podbor(win, name)
            used_podbor = True

    # Проверка по факту, а не по тому, что все клики "прошли без ошибок":
    # живьём поймали случай, когда весь диалог отработал без единого
    # исключения, а поле "Контрагент:" в итоге осталось ПУСТЫМ (контрагента
    # не нашлось в справочнике 1С под этим именем, и код молча кликнул не
    # туда). Явно сверяем итоговый текст поля — если не совпадает, это
    # реальная проблема с ЭТИМ контрагентом, а не "всё ок".
    actual = _contractor_field_text(win)
    if not _names_match(name, actual):
        if not used_podbor:
            # Перезапись на месте дала не тот результат — последний шанс:
            # полный "Подбор". В отличие от ветки выше, тут диалог "Контрагент"
            # уже ЗАКРЫТ (успела нажаться "ОК" в конце _set_contractor_in_place)
            # — открываем его заново кнопкой "...", прежде чем звать "Подбор".
            print(f"  [1С] ⚠ После перезаписи на месте поле показывает {actual!r} — пробую 'Подбор'...")
            _open_contractor_list(win, ellipsis_points)
            _set_contractor_via_podbor(win, name)
            actual = _contractor_field_text(win)
        if not _names_match(name, actual):
            raise RuntimeError(
                f"После выбора контрагента поле 'Контрагент:' показывает {actual!r}, а не '{name}' — "
                f"похоже, контрагент не найден в справочнике 1С под этим именем."
            )
    способ = "Подбор" if used_podbor else "перезапись на месте"
    print(f"  Контрагент '{name}' выбран через {способ} (проверено: поле показывает '{actual}').")


def _date_parts(text: str) -> tuple:
    """
    Числа из даты как кортеж: "8/24/2026" и "08/24/2026" дают одно и то же
    (8, 24, 2026).

    Сравнивать сырые строки или сплошную последовательность цифр нельзя:
    маска поля ПЕРЕРИСОВЫВАЕТ введённое по-своему — вставленное "08/24/2026"
    поле показывает как "8/24/2026", и посимвольная сверка объявила бы
    правильно введённую дату неверной.
    """
    return tuple(int(part) for part in re.findall(r"\d+", text or ""))


def _edit_text(edit) -> str:
    """
    Текущее содержимое Edit, не путая его с ЭХОМ ПОДПИСИ: у полей 1С
    window_text() нередко отдаёт подпись ("Период:"), а не значение —
    настоящее лежит в legacy Value / ValuePattern (та же ловушка, что на
    полях документов, см. usluga_report._field_value).
    """
    for getter in (
        lambda: edit.window_text(),
        lambda: edit.legacy_properties().get("Value", ""),
        lambda: edit.iface_value.CurrentValue,
    ):
        try:
            value = (getter() or "").strip()
        except Exception:
            continue
        if value and not value.endswith(":"):
            return value
    return ""


def _format_date_like_field(current: str, day: int, month: int, year: int) -> str:
    """
    Дата в том формате, который ждёт САМО поле.

    ЖИВЬЁМ ПОДТВЕРЖДЕНО на сервере: там региональные настройки английские и
    маска поля периода — M/d/yyyy ("8/24/2026"), а мы вставляли "24.08.2026".
    Мало того что разделитель другой — переставлены местами день и месяц,
    поэтому даже вставка одних цифр дала бы не ту дату.

    Формат определяем по разделителю в ТЕКУЩЕМ содержимом поля: "/" —
    американский порядок (месяц первым), иначе привычный ДД.ММ.ГГГГ. Это
    надёжнее, чем смотреть на локаль своего процесса: важна локаль 1С.
    """
    if "/" in current:
        return f"{month:02d}/{day:02d}/{year:04d}"
    return f"{day:02d}.{month:02d}.{year:04d}"


def set_period(win, date_from: str, date_to: str) -> None:
    """
    Устанавливает период в полях «от» и «до».
    На вход — ДД.ММ.ГГГГ (например, 01.05.2026); в поле уходит формат,
    который это поле само показывает (см. _format_date_like_field).
    Поля периода — первые два Edit на форме (edits[0] и edits[1]).

    После каждой вставки ПРОВЕРЯЕМ, что в поле действительно нужная дата, и
    повторяем до трёх раз. Причина та же, что в emit_transactions
    (lib1c.type_digits): клик может "успешно" пройти через UIA-слой, не
    переведя туда реальный фокус, и вставка молча уходит в никуда — а
    дальше отчёт формируется за старый период, и это никак не заметно.
    Сравниваем ТОЛЬКО числа даты: маска сама расставляет разделители.
    """
    vi.focus(win)
    time.sleep(0.3)

    print(f"  Ставлю период: {date_from} – {date_to}")

    def _find_edits():
        edits = win.descendants(control_type="Edit")
        return edits if len(edits) >= 2 else None

    edits = _wait_for(_find_edits, "поля периода (2+ Edit)")

    for idx, date_val in enumerate([date_from, date_to]):
        edit = edits[idx]
        day, month, year = (int(part) for part in date_val.split("."))

        placed = False
        actual = ""
        for attempt in range(1, 4):
            wanted = _format_date_like_field(_edit_text(edit), day, month, year)

            _click_input_preserve_cursor(edit)
            time.sleep(0.3)
            _send_keys(None, "{HOME}")
            time.sleep(0.1)
            _send_keys(None, "+{END}")
            time.sleep(0.1)
            _set_clipboard_text(wanted)
            _send_keys(None, "+{INSERT}")
            time.sleep(0.5)

            actual = _edit_text(edit)
            if _date_parts(actual) == _date_parts(wanted):
                placed = True
                if attempt > 1 or wanted != date_val:
                    print(f"    поле {idx + 1}: {actual}")
                break
            print(f"    попытка {attempt}/3: в поле {actual!r}, а нужно {wanted!r} — повторяю...")

        if not placed:
            raise RuntimeError(
                f"Не смог поставить дату {date_val} в поле периода: там осталось {actual!r}. "
                f"Отчёт сформировался бы за чужой период, поэтому дальше не иду."
            )

    print("  Период установлен.")


def generate(win, wait_seconds: int = 0.2, timeout: float = 20.0) -> None:
    """
    Формирует отчёт (Ctrl+Enter — эквивалент кнопки 'Сформировать').

    Сначала ждём по индикатору "Отчёт формируется..." (появился -> исчез —
    надёжный прямой сигнал, не зависящий от того, быстро или долго считается
    отчёт), затем ещё раз проверяем стабильность строк отчёта как
    дополнительную подстраховку на случай, если сетка дорисовывается уже
    после исчезновения индикатора.
    """
    old_snapshot = _report_snapshot(win)
    _send_keys(win, "^{ENTER}")
    _wait_for_generation_indicator(win, appear_timeout=2.0, disappear_timeout=max(timeout, 120.0), poll=0.1)
    # К этому моменту индикатор "Отчёт формируется..." уже подтвердил, что
    # отчёт готов — эта проверка снимков строк теперь просто короткая
    # подстраховка (сетка иногда дорисовывается чуть после исчезновения
    # индикатора), а не основной механизм ожидания, поэтому держим её
    # быстрой.
    _wait_for_report_refresh(win, old_snapshot, timeout=5.0, poll=0.2, stable_time=0.4)
    time.sleep(wait_seconds)


_SAVE_DIALOG_CLASS = "#32770"

# Заголовок диалога сохранения — НА ОБОИХ ЯЗЫКАХ. ЖИВЬЁМ ПОДТВЕРЖДЕНО на
# сервере: там интерфейс 1С английский и окно называется "Save As", а не
# "Сохранение" — поиск по русскому заголовку его не находил, и шаг BNPL
# падал с "Диалог сохранения не появился", хотя окно было прямо на экране.
# Сравниваем в нижнем регистре, поэтому хватает кусков без первой буквы.
_SAVE_DIALOG_TITLE_SUBSTRS = ("охранен", "save as", "сохранение")

# Пункт списка "Тип файла" — ищем по РАСШИРЕНИЮ, а не по названию: оно
# переводится ("Лист Excel2007-... (*.xlsx)" против "Excel 2007-... sheet
# (*.xlsx)", причём в английском ещё и пробел после Excel), а "(*.xlsx)"
# одинаково в любой локали.
_XLSX_FILE_TYPE_SUBSTR = "(*.xlsx)"


def _dialog_hwnds():
    """Все видимые окна класса #32770 с заголовком 'Сохранение'."""
    found = []

    def _cb(hwnd, _):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            if win32gui.GetClassName(hwnd) != _SAVE_DIALOG_CLASS:
                return True
            title = (win32gui.GetWindowText(hwnd) or "").lower()
            if any(part in title for part in _SAVE_DIALOG_TITLE_SUBSTRS):
                found.append(hwnd)
        except Exception:
            pass
        return True

    win32gui.EnumWindows(_cb, None)
    return found


def _has_filename_field(hwnd) -> bool:
    """True, если это НАСТОЯЩИЙ диалог сохранения, а не окно-ошибка."""
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        dlg = Application(backend="uia").connect(process=pid).window(handle=hwnd)
        return dlg.child_window(auto_id="FileNameControlHost", control_type="ComboBox").exists(timeout=0.5)
    except Exception:
        return False


def _find_save_dialog_hwnd():
    """
    hwnd НАСТОЯЩЕГО диалога сохранения 1С или None.

    ЖИВЬЁМ ПОДТВЕРЖДЕНО (дважды):
      * это ОТДЕЛЬНОЕ top-level окно — EnumWindows его видит, а pywinauto
        не отдаёт ни через Desktop().windows(), ни через
        Application.windows() (владелец окна — главный фрейм 1С, поэтому в
        UIA-дереве диалог висит его ребёнком). Ищем через win32gui;
      * окно ОШИБКИ 1С ("Этот путь не существует") имеет ТОТ ЖЕ класс
        #32770 и ТОТ ЖЕ заголовок "Сохранение" — по заголовку их не
        различить, и код мог начать работать с окном-ошибкой. Настоящий
        диалог отличаем по наличию поля "Имя файла"
        (auto_id=FileNameControlHost).
    """
    candidates = _dialog_hwnds()
    for hwnd in candidates:
        if _has_filename_field(hwnd):
            return hwnd
    return None


def _is_exit_confirmation(hwnd) -> bool:
    """Это окно подтверждения выхода из 1С ("Завершить работу?")?"""
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        dlg = Application(backend="uia").connect(process=pid).window(handle=hwnd)
        for btn in dlg.descendants(control_type="Button"):
            caption = (btn.window_text() or "").strip().lower()
            if (caption.startswith("заверш") or caption == "exit"
                    or caption.startswith("продолжить") or caption == "continue"):
                return True
    except Exception:
        pass
    return False


_watchdog_thread = None
_watchdog_stop = threading.Event()


def start_message_box_watchdog(interval: float = 1.5) -> None:
    """
    Фоновый сторож: сам закрывает всплывшие модальные сообщения 1С, пока
    идёт работа.

    Зачем: ЖИВЬЁМ подтверждено — одно случайное окно ("Поле объекта
    недоступно для чтения (Расшифровка)" и т.п.) модально блокирует
    главное окно, и прогон, который до этого шёл идеально, встаёт намертво:
    все следующие поиски элементов падают, будто окно опустело. Разовые
    вызовы dismiss_1c_message_boxes() в отдельных местах этого не ловят —
    окно может выскочить в любой момент, в том числе посреди длинной
    операции. Сторож проверяет это раз в interval секунд всё время работы.

    Поток демонический — умирает вместе с процессом, отдельно останавливать
    не обязательно (но см. stop_message_box_watchdog перед закрытием 1С).
    """
    global _watchdog_thread
    if _watchdog_thread is not None and _watchdog_thread.is_alive():
        return

    def _loop():
        seen_before: set = set()
        while not _watchdog_stop.wait(interval):
            try:
                dismiss_1c_message_boxes(max_rounds=1)
            except Exception:
                pass
            # Зависшие ФОРМЫ (напр. "Выбор поля") сообщениями не считаются,
            # но блокируют главное окно так же намертво. Снимаем их, если
            # окно провисело ДВА тика подряд: за это время наш собственный
            # код успел бы с ним поработать и закрыть, значит оно ничьё.
            try:
                now = _stray_window_handles()
                for order, hwnd in enumerate(sorted(now & seen_before), start=1):
                    print(f"  [1С] Зависшее окно {_window_title_head(hwnd)!r} — закрываю "
                          f"({'Escape' if order == 1 else 'Отмена'}).")
                    _escape_window(hwnd, order=order)
                    time.sleep(0.4)
                    if win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd):
                        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                seen_before = now
            except Exception:
                pass

    _watchdog_stop.clear()
    _watchdog_thread = threading.Thread(target=_loop, name="1c-msgbox-watchdog", daemon=True)
    _watchdog_thread.start()


def stop_message_box_watchdog() -> None:
    """Останавливает сторожа — нужно перед штатным закрытием 1С."""
    _watchdog_stop.set()


# Рабочие диалоги, которые НЕЛЬЗЯ принимать за случайное сообщение 1С.
# Они того же класса (V8NewLocalFrameBaseWnd) и с тем же "1С:Предприятие"
# в заголовке, что и настоящие сообщения, и у них есть кнопка "ОК" —
# сторож модальных окон нажимал её и закрывал диалог прямо посреди работы.
_WORKING_DIALOG_TITLES = ("Контрагент", "Контрагенты")

# Управляющие элементы, которых у настоящего сообщения не бывает: если в
# окне есть поле ввода, таблица или список — это рабочая форма, а не
# сообщение, и трогать её нельзя.
_NON_MESSAGE_CONTROLS = ("Edit", "Table", "List", "DataGrid", "Tree", "CheckBox", "ComboBox")


def _is_working_dialog(hwnd, dlg=None) -> bool:
    """
    Это рабочий диалог (а не всплывшее сообщение 1С)?

    Два признака, любого достаточно: знакомый заголовок и наличие
    управляющих элементов. Заголовок 1С пишет как "Контрагент
    (1С:Предприятие)", поэтому берём часть до скобки.
    """
    try:
        title = (win32gui.GetWindowText(hwnd) or "").split("(")[0].strip()
    except Exception:
        title = ""
    if title in _WORKING_DIALOG_TITLES:
        return True
    if dlg is None:
        return False
    return not _looks_like_message_box(dlg)


def _looks_like_message_box(dlg) -> bool:
    """
    Окно — настоящее сообщение (текст + кнопка), а не форма с полями и
    списками? Сторож в фоне закрывает ТОЛЬКО такие: форму он трогать не
    должен, иначе закроет диалог, с которым прямо сейчас идёт работа.
    """
    for control_type in _NON_MESSAGE_CONTROLS:
        try:
            if dlg.descendants(control_type=control_type):
                return False
        except Exception:
            continue
    return True


# Диалоги, которые наш же код открывает НАМЕРЕННО на время настройки
# фильтра. Пока идёт эта настройка, сторож их не трогает; в любое другое
# время такое окно — зависший мусор, и его надо снять.
_EXPECTED_DIALOG_TITLES: set = set()


class expected_dialogs:
    """
    Контекст "сейчас мы сами открываем вот эти окна — не закрывай их".

    Нужен, потому что одно и то же окно бывает и рабочим, и зависшим:
    "Выбор поля" законно открывается при настройке фильтра, а в остальное
    время просто висит и блокирует главное окно (живьём: после сообщения
    1С оно осталось на экране, и весь остаток прогона шёл вхолостую).
    """

    def __init__(self, *titles):
        self.titles = set(titles)

    def __enter__(self):
        _EXPECTED_DIALOG_TITLES.update(self.titles)
        return self

    def __exit__(self, *exc):
        _EXPECTED_DIALOG_TITLES.difference_update(self.titles)
        return False


def _window_title_head(hwnd) -> str:
    """Начало заголовка окна — до " - " и до скобки (для сообщений в лог)."""
    try:
        title = win32gui.GetWindowText(hwnd) or ""
    except Exception:
        return ""
    return title.split(" - ")[0].split("(")[0].strip()


def _is_stray_1c_window(hwnd, main_hwnd) -> bool:
    """
    Это чужое окно 1С, которое висит зря?

    Признаки: видимое, принадлежит тому же процессу, что и главное окно,
    самим главным окном не является, и его заголовок не в списке тех, что
    мы открываем осознанно. По классу окна НЕ фильтруем: "Выбор поля"
    оказалось другого класса, чем сообщения, и прежняя проверка его молча
    пропускала. GetParent тоже НЕ проверяем: у ВЛАДЕЕМОГО диалога он
    возвращает владельца (главное окно 1С), и проверка отбрасывала как раз
    те окна, ради которых всё и делается.
    """
    try:
        if not hwnd or hwnd == main_hwnd or not win32gui.IsWindowVisible(hwnd):
            return False
        if not main_hwnd:
            return False
        _, main_pid = win32process.GetWindowThreadProcessId(main_hwnd)
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        if pid != main_pid:
            return False
    except Exception:
        return False

    try:
        title = (win32gui.GetWindowText(hwnd) or "").strip()
    except Exception:
        return False
    if not title:
        return False

    # Сверяем по НАЧАЛУ заголовка: 1С дописывает имя конфигурации и
    # "(1C:Enterprise)", а у окна настроек внутри заголовка ещё и своё
    # " - ", так что резать по разделителю нельзя.
    protected = set(_WORKING_DIALOG_TITLES) | _EXPECTED_DIALOG_TITLES
    return not any(title.startswith(prefix) for prefix in protected)


def _stray_window_handles() -> set:
    """Хендлы всех сейчас висящих чужих окон 1С (см. _is_stray_1c_window)."""
    main_hwnd = _ACTIVE_HWND
    found = set()

    def _cb(hwnd, _):
        try:
            if _is_stray_1c_window(hwnd, main_hwnd):
                found.add(hwnd)
        except Exception:
            pass
        return True

    try:
        win32gui.EnumWindows(_cb, None)
    except Exception:
        pass
    return found


def escape_stray_dialogs(max_windows: int = 6) -> int:
    """
    Закрывает ВСЕ чужие окна 1С, оставшиеся висеть от предыдущего шага, —
    по одному нажатию на окно (два окна = два нажатия). Главное окно 1С не
    трогает. Возвращает число закрытых окон.
    """
    main_hwnd = _ACTIVE_HWND
    closed = 0

    for order in range(1, max_windows + 1):
        strays = []

        def _cb(hwnd, _):
            try:
                if not _is_stray_1c_window(hwnd, main_hwnd):
                    return True
                strays.append((hwnd, win32gui.GetWindowText(hwnd) or ""))
            except Exception:
                pass
            return True

        try:
            win32gui.EnumWindows(_cb, None)
        except Exception:
            break
        if not strays:
            break

        hwnd, title = strays[-1]  # верхнее окно первым — оно поверх остальных
        print(f"  [1С] Закрываю зависшее окно ({'Escape' if order == 1 else 'Отмена'}): {title!r}")
        _escape_window(hwnd, order=order)
        time.sleep(0.6)
        if win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd):
            try:
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                time.sleep(0.4)
            except Exception:
                pass
        closed += 1

    return closed


# Кнопки, которыми окно закрывается БЕЗ последствий (ничего не применяя).
_CANCEL_BUTTON_TITLES = ("отмена", "cancel", "закрыть", "close")


def _click_cancel_button(hwnd) -> bool:
    """
    Жмёт в окне "Отмена" ("Cancel"), если такая кнопка есть. True — нажали.

    Это надёжнее Escape: живьём окно "Выбор поля" оставалось висеть и
    после Escape, и после WM_CLOSE, а кнопка "Отмена" у него прямо на
    форме. Берём именно отменяющую кнопку, а не "Выбрать"/"ОК" — окно
    чужое, ничего применять в нём нельзя.
    """
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        dlg = Application(backend="uia").connect(process=pid).window(handle=hwnd)
        for btn in dlg.descendants(control_type="Button"):
            try:
                caption = (btn.window_text() or "").strip().lower()
            except Exception:
                continue
            if caption in _CANCEL_BUTTON_TITLES:
                vi.click(btn, settle=0.4)
                return True
    except Exception:
        pass
    return False


def _press_escape(hwnd) -> None:
    """Escape прямо в окно."""
    try:
        vi.send_keys(hwnd, "{ESC}", settle=0.3, required=False)
    except Exception:
        pass


def _window_alive(hwnd) -> bool:
    try:
        return bool(win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd))
    except Exception:
        return False


def _escape_window(hwnd, order: int = 1) -> None:
    """
    Закрывает окно, ничего в нём не применяя. ПОРЯДОК — по прямому
    указанию пользователя:

      ПЕРВОЕ окно  -> Escape (это всплывшее сообщение 1С, Escape его гасит);
      ВТОРОЕ и дальше -> сразу кнопка "Отмена" (за сообщением вылезает
      форма вроде "Выбор поля", на Escape она живьём не реагировала, а
      кнопка "Отмена" у неё прямо на форме).

    order — номер окна в текущем проходе (1, 2, 3...). Если выбранный
    способ не сработал, пробуем второй, и только потом WM_CLOSE: цель —
    чтобы окно точно ушло, иначе весь дальнейший прогон идёт вхолостую.
    """
    if order <= 1:
        _press_escape(hwnd)
        if _window_alive(hwnd):
            _click_cancel_button(hwnd)
    else:
        if not _click_cancel_button(hwnd):
            _press_escape(hwnd)


def dismiss_1c_message_boxes(max_rounds: int = 3) -> int:
    """
    Закрывает модальные сообщения самой 1С ("1С:Предприятие" с кнопкой ОК).

    ЖИВЬЁМ ПОДТВЕРЖДЕНО: одно такое окно парализует ВЕСЬ дальнейший прогон —
    главное окно модально заблокировано, и следующий шаг не находит уже
    ничего (ни вкладок разделов, ни пунктов меню), падая с
    ElementNotFoundError на ровном месте. Поймано на "Поле объекта
    недоступно для чтения (Расшифровка)" — 1С показала его после лишнего
    Enter, ушедшего в таблицу отчёта.

    Класс окна — V8NewLocalFrameBaseWnd с заголовком "1С:Предприятие"
    (это НЕ стандартный #32770, поэтому обычный поиск диалогов его не
    видит). Возвращает число закрытых окон.
    """
    closed = 0
    for _ in range(max_rounds):
        boxes = []

        def _cb(hwnd, _):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                if win32gui.GetClassName(hwnd) != "V8NewLocalFrameBaseWnd":
                    return True
                # ЖИВЬЁМ ПОДТВЕРЖДЕНО: на сервере платформа англоязычная и
                # заголовок у этих окон "1C:Enterprise", а не
                # "1С:Предприятие" — проверка на русское слово там не
                # срабатывала НИ РАЗУ, поэтому сообщения не закрывались
                # вообще, а прогон продолжался вслепую.
                title = win32gui.GetWindowText(hwnd) or ""
                if "Предприятие" in title or "Enterprise" in title:
                    boxes.append(hwnd)
            except Exception:
                pass
            return True

        win32gui.EnumWindows(_cb, None)
        if not boxes:
            break

        for hwnd in boxes:
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                dlg = Application(backend="uia").connect(process=pid).window(handle=hwnd)

                # Рабочие диалоги ("Контрагент", "Контрагенты" и любое окно
                # с полями/таблицами) — НЕ ТРОГАЕМ. У них тоже есть "ОК", и
                # раньше сторож нажимал её, закрывая диалог посреди выбора
                # контрагента.
                if _is_working_dialog(hwnd, dlg):
                    continue

                texts = [t.window_text().strip() for t in dlg.descendants(control_type="Text")]
                message = " | ".join(t for t in texts if t)
                if message:
                    print(f"  [1С] Сообщение 1С: {message} — закрываю.")
                clicked = False
                for btn in dlg.descendants(control_type="Button"):
                    if (btn.window_text() or "").strip().lower() in ("ок", "ok", "да", "yes", "закрыть"):
                        vi.click(btn, settle=0.4)
                        clicked = True
                        break
                if not clicked:
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            except Exception as e:
                print(f"  ⚠ Не смог закрыть сообщение 1С: {e}")
                try:
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                except Exception:
                    pass

            # Кнопка и WM_CLOSE срабатывают не всегда — тогда "Отмена"/Escape
            # прямо в это окно. По ОДНОМУ нажатию на окно: два висящих
            # сообщения закрываются двумя нажатиями, как и должно быть.
            time.sleep(0.4)
            try:
                if win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd):
                    _escape_window(hwnd)
            except Exception:
                pass

            closed += 1
            time.sleep(0.5)
    return closed


def _error_box_text(hwnd) -> str:
    """Текст окна-ошибки — чтобы в логе было видно, на что именно ругнулась 1С."""
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        dlg = Application(backend="uia").connect(process=pid).window(handle=hwnd)
        parts = [t.window_text().strip() for t in dlg.descendants(control_type="Text")]
        return " | ".join(p for p in parts if p)
    except Exception:
        return ""


def _dismiss_error_boxes(rounds: int = 3) -> int:
    """
    Закрывает окна-ошибки 1С (те же #32770 с заголовком "Сохранение", но
    БЕЗ поля "Имя файла"). Они модальные: пока такое висит, следующий
    Ctrl+S просто ничего не откроет, а диалог сохранения не даст нажать
    "Сохранить".

    Сначала жмём кнопку ("ОК"/"OK"/"Да") — WM_CLOSE на message box иногда
    не срабатывает; он же остаётся запасным вариантом. Возвращает, сколько
    окон закрыли.
    """
    closed = 0
    for _ in range(rounds):
        boxes = [h for h in _dialog_hwnds() if not _has_filename_field(h)]
        if not boxes:
            break
        for hwnd in boxes:
            # Подтверждение выхода из 1С ("Завершить работу?") — НЕ трогаем:
            # его закрывает close_1c() сам, когда действительно закрывает
            # программу. Иначе сторож гасил бы его и мешал штатному выходу.
            if _is_exit_confirmation(hwnd):
                continue
            # Рабочие диалоги не трогаем — см. _is_working_dialog.
            if _is_working_dialog(hwnd):
                continue
            text = _error_box_text(hwnd)
            if text:
                print(f"  [1С] Окно ошибки: {text}")
            done = False
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                dlg = Application(backend="uia").connect(process=pid).window(handle=hwnd)
                for btn in dlg.descendants(control_type="Button"):
                    if (btn.window_text() or "").strip().lower() in ("ок", "ok", "да", "yes", "закрыть"):
                        vi.click(btn, settle=0.4)
                        done = True
                        break
            except Exception:
                pass
            if not done:
                try:
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                except Exception:
                    pass
            closed += 1
            time.sleep(0.5)
    if closed:
        print(f"  [1С] Закрыл окно(а) ошибки: {closed}.")
    return closed


def _wait_save_dialog(timeout: float = 20.0, poll: float = 0.3):
    """Ждёт появления диалога сохранения, возвращает обёртку окна или None."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        hwnd = _find_save_dialog_hwnd()
        if hwnd:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            app = Application(backend="uia").connect(process=pid)
            return app.window(handle=hwnd)
        time.sleep(poll)
    return None


def _close_stray_save_dialog() -> None:
    """
    Убирает всё, что могло остаться с прошлого (в т.ч. прерванного) раза:
    сначала окна-ошибки, потом сам диалог сохранения — иначе новый Ctrl+S
    упрётся в модальное окно.
    """
    _dismiss_error_boxes()
    hwnd = _find_save_dialog_hwnd()
    if not hwnd:
        return
    print("  [1С] Диалог сохранения уже открыт — закрываю перед новым сохранением.")
    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
    time.sleep(1.0)
    _dismiss_error_boxes()


def _wait_file_written(path: Path, timeout: float = 300.0, poll: float = 0.5, stable_checks: int = 3) -> None:
    """
    Ждёт, пока 1С реально допишет файл: он должен появиться и перестать
    расти (stable_checks одинаковых замеров размера подряд). Большой отчёт
    (десятки тысяч строк) пишется заметное время после закрытия диалога.
    """
    deadline = time.time() + timeout
    last_size = -1
    stable = 0
    while time.time() < deadline:
        if path.exists():
            size = path.stat().st_size
            if size > 0 and size == last_size:
                stable += 1
                if stable >= stable_checks:
                    return
            else:
                stable = 0
            last_size = size
        time.sleep(poll)
    raise TimeoutError(f"Файл '{path}' так и не был дописан за {timeout} сек.")


def save_report_as_xlsx(win, path, timeout: float = 300.0) -> Path:
    """
    Сохраняет текущий отчёт 1С в файл .xlsx через Ctrl+S — вместо
    Ctrl+A/Ctrl+C.

    ПОЧЕМУ так, а не буфер обмена (по прямому указанию пользователя
    2026-09-19): на больших отчётах ("Отчет по должникам (копия)" — десятки
    тысяч строк) 1С НЕ вытягивает копирование — буфер обмена приходит
    пустым/обрывается. Сохранение в файл тот же объём отдаёт стабильно.

    Шаги (все живьём подтверждены на реальном диалоге):
      Ctrl+S -> окно "Сохранение" (класс #32770, см. _find_save_dialog_hwnd)
      -> ComboBox "Тип файла:" (auto_id="FileTypeControlHost") -> пункт
      списка с "Excel2007" ("Лист Excel2007-... (*.xlsx)")
      -> Edit "Имя файла:" (auto_id="1001") — полный путь
      -> кнопка "Сохранить" (auto_id="1")
      -> ждём, пока файл появится и перестанет расти.
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        path.unlink()

    _close_stray_save_dialog()

    vi.focus(win)
    time.sleep(0.3)
    _force_foreground(win.handle)
    time.sleep(0.3)

    # Клик в область отчёта — чтобы Ctrl+S ушёл именно в табличный
    # документ, а не в шапку формы (та же причина, что и в copy_report_data).
    try:
        rect = win.rectangle()
        x = (rect.right - rect.left) // 2
        y = int((rect.bottom - rect.top) * 0.65)
        _click_input_preserve_cursor(win, coords=(x, y))
        time.sleep(0.4)
    except Exception as e:
        print(f"  Не удалось кликнуть в область отчёта: {e}")

    _force_foreground(win.handle)
    time.sleep(0.3)

    print("  [1С] Ctrl+S — сохраняю отчёт в файл...")
    _send_ctrl_key(ord("S"))

    dlg = _wait_save_dialog(timeout=20.0)
    if dlg is None:
        raise RuntimeError(
            "Диалог сохранения ('Сохранение' / 'Save As') не появился после Ctrl+S."
        )

    vi.focus(dlg)
    time.sleep(0.3)

    # 1) Тип файла -> "Лист Excel2007-... (*.xlsx)"
    combo = dlg.child_window(auto_id="FileTypeControlHost", control_type="ComboBox")
    combo.expand()
    time.sleep(0.6)
    target = None
    for item in dlg.descendants(control_type="ListItem"):
        try:
            text = item.window_text() or ""
        except Exception:
            continue
        if _XLSX_FILE_TYPE_SUBSTR in text:
            target = item
            break
    if target is None:
        raise RuntimeError(f"В списке 'Тип файла' не нашёл пункт с '{_XLSX_FILE_TYPE_SUBSTR}'.")
    print(f"  [1С] Тип файла: {target.window_text()}")
    _click_input_preserve_cursor(target)
    time.sleep(0.5)

    # 2) Имя файла — полный путь (обычный Windows-диалог так понимает папку+имя).
    #
    # ЖИВЬЁМ ПОДТВЕРЖДЁННЫЙ БАГ: диалог ПОДСТАВЛЯЕТ в это поле путь из
    # прошлого раза. Если наш текст туда не лёг, 1С сохраняет по СТАРОМУ
    # пути — а если той папки уже нет, показывает модальное "Этот путь не
    # существует", которое блокирует и следующий запуск. Поэтому: пишем
    # значение напрямую (UIA ValuePattern — вообще без клавиатуры), с
    # запасным вариантом через буфер, и ОБЯЗАТЕЛЬНО перечитываем поле —
    # не совпало, дальше не идём.
    edit = dlg.child_window(auto_id="1001", control_type="Edit")
    target_text = str(path)

    def _read_field() -> str:
        """
        Значение поля, не путая его с ЭХОМ ПОДПИСИ: у этого Edit
        window_text() отдаёт "Имя файла:" (свою подпись), а не содержимое —
        ровно та же ловушка, что уже ловилась на полях документов 1С
        (см. usluga_report._field_value). Настоящее значение — в
        legacy Value / ValuePattern.
        """
        for getter in (
            lambda: edit.window_text(),
            lambda: edit.legacy_properties().get("Value", ""),
            lambda: edit.iface_value.CurrentValue,
        ):
            try:
                value = (getter() or "").strip().strip('"')
            except Exception:
                continue
            if value and not value.endswith(":"):
                return value
        return ""

    def _field_ok() -> bool:
        return _read_field() in (target_text, path.name, path.stem)

    # ВАЖНО: путь вводим именно как в поле ввода (выделить -> удалить ->
    # вставить из буфера), а НЕ через UIA SetValue. ЖИВЬЁМ ПОДТВЕРЖДЕНО:
    # SetValue меняет ВИДИМЫЙ текст поля, но сам диалог этого не замечает и
    # по "Сохранить" сохраняет по СТАРОМУ запомненному пути (в логе была
    # ошибка про давно удалённую временную папку, хотя в поле уже стоял
    # новый путь). Правка через буфер идёт обычным путём обработки ввода,
    # и диалог её принимает.
    _click_input_preserve_cursor(edit)
    time.sleep(0.2)
    _send_ctrl_key(ord("A"), target=edit)
    time.sleep(0.1)
    _send_keys(edit, "{DELETE}", settle=0.1)
    _set_clipboard_text(target_text)
    _send_ctrl_key(ord("V"), target=edit)
    time.sleep(0.4)

    if not _field_ok():
        # Запасной вариант — Shift+Insert (в 1С-полях этого проекта он
        # исторически надёжнее Ctrl+V).
        _send_keys(edit, "+{INSERT}")
        time.sleep(0.4)

    if not _field_ok():
        raise RuntimeError(
            "Не удалось вписать путь в поле 'Имя файла' диалога сохранения — там осталось "
            f"{_read_field()!r} вместо {target_text!r}."
        )
    print(f"  [1С] Имя файла: {target_text}")

    # 3) Сохранить. Сначала Enter прямо в поле имени — для стандартного
    # диалога это и есть "принять введённое имя" (кнопка "Сохранить" при
    # программной правке поля может сработать по старому состоянию). Если
    # диалог после Enter остался — дожимаем кнопкой.
    _send_keys(edit, "{ENTER}", settle=0.8)
    if _find_save_dialog_hwnd():
        _click_input_preserve_cursor(dlg.child_window(auto_id="1", control_type="Button"))

    # Если 1С ответила модальной ошибкой ("Этот путь не существует" и т.п.) —
    # ждать файл бессмысленно, он не появится: сразу гасим окно и падаем с
    # понятным текстом, а не через таймаут в несколько минут.
    time.sleep(1.0)
    stuck = [h for h in _dialog_hwnds() if not _has_filename_field(h)]
    if stuck:
        message = _error_box_text(stuck[0])
        _dismiss_error_boxes()
        _close_stray_save_dialog()
        raise RuntimeError(f"1С не смогла сохранить файл: {message or 'модальное окно ошибки'}")

    print("  [1С] Жду, пока файл допишется...")
    _wait_file_written(path, timeout=timeout)
    print(f"  [1С] Файл сохранён: {path} ({path.stat().st_size} байт)")
    return path


def _cell_to_text(value) -> str:
    """
    Значение ячейки .xlsx -> строка в том же виде, в каком её раньше
    отдавал буфер обмена 1С (чтобы существующие парсеры не менять).
    """
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, _dt.datetime):
        if value.hour or value.minute or value.second:
            return value.strftime("%d.%m.%Y %H:%M:%S")
        return value.strftime("%d.%m.%Y")
    if isinstance(value, _dt.date):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def read_xlsx_as_tsv(path) -> str:
    """
    Читает сохранённый отчёт .xlsx и отдаёт TSV-текст РОВНО в том формате,
    который раньше приходил из буфера обмена — так все существующие
    парсеры (parse_clipboard_report, parse_debtors_report и т.д.)
    продолжают работать без изменений.

    Пишем через csv.writer с тем же диалектом (tab + кавычки), которым
    парсеры читают, — поэтому многострочные ячейки (в "Отчет по должникам
    (копия)" столбец "Номенклатура" реально содержит переносы строк внутри
    одной ячейки) корректно переживают круг "xlsx -> текст -> csv.reader".

    ЖИВЬЁМ ПОДТВЕРЖДЕНО (реальный файл "Отчет по должникам (копия)",
    20739 строк): в .xlsx из 1С колонок БОЛЬШЕ, чем логических столбцов
    отчёта — ширина колонки табличного документа разложена на несколько
    физических колонок, между данными появляются колонки-разделители
    (здесь: 14 колонок вместо 11 — пустые №1 и №3, плюс №2, в которой
    лежит ТОЛЬКО "Период отчета: ..." из шапки отчёта). В буфере обмена
    их не было, поэтому парсеры съехали бы по столбцам.

    Настоящие столбцы определяем по строке-образцу — строке с
    НАИБОЛЬШИМ числом заполненных ячеек (это строка заголовка отчёта:
    "Партнер | Контакт | Заказ | ..."), она заполнена во всех настоящих
    столбцах и пуста в разделителях. Дополнительно оставляем столбцы,
    заполненные хотя бы в 30% строк, — страховка для отчётов, где
    заголовок разбит на две строки и часть подписей уходит во вторую.
    """
    import openpyxl  # локальный импорт — reports.py грузится и там, где openpyxl не нужен

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.active
        rows = [[_cell_to_text(v) for v in row] for row in ws.iter_rows(values_only=True)]
    finally:
        wb.close()

    width = max((len(r) for r in rows), default=0)
    filled = [[bool(i < len(r) and r[i].strip()) for i in range(width)] for r in rows]

    template = max(filled, default=[], key=lambda f: sum(f)) if filled else []
    threshold = max(2, int(len(rows) * 0.3))
    counts = [sum(f[i] for f in filled) for i in range(width)]
    keep = [i for i in range(width) if (template and template[i]) or counts[i] >= threshold]

    dropped = width - len(keep)
    if dropped:
        print(f"  [1С] В файле {width} колонок, из них {dropped} колонок-разделителей — выбрасываю.")

    buf = io.StringIO()
    writer = csv.writer(buf, delimiter="\t", quotechar='"', lineterminator="\n",
                        quoting=csv.QUOTE_MINIMAL)
    for r in rows:
        writer.writerow([r[i] if i < len(r) else "" for i in keep])
    return buf.getvalue()


def export_report_data(win, timeout: float = 300.0) -> str:
    """
    Замена copy_report_data() для больших отчётов: сохраняет отчёт во
    временный .xlsx (Ctrl+S), читает его и УДАЛЯЕТ файл, возвращая тот же
    TSV-текст, что раньше давал буфер обмена.

    Файл кладётся в ПОСТОЯННУЮ папку проекта (tools/_1c_export) с уникальным
    именем, и после чтения удаляется только САМ ФАЙЛ — папка остаётся.

    ЖИВЬЁМ ПОДТВЕРЖДЁННЫЙ БАГ (из-за которого так и сделано): раньше папка
    создавалась через tempfile.mkdtemp() и удалялась целиком в finally.
    Если прогон прерывали (Ctrl+C, убитый процесс) ровно в момент ожидания
    файла, finally успевал снести папку, пока диалог сохранения в 1С был
    ещё открыт — и 1С по нажатию "Сохранить" выдавала модальное
    "Этот путь не существует", которое блокировало и следующий запуск.
    Постоянная папка убирает эту гонку полностью.
    """
    out_dir = Path(__file__).resolve().parent.parent / "tools" / "_1c_export"
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / f"report_{int(time.time())}.xlsx"
    try:
        save_report_as_xlsx(win, path, timeout=timeout)
        text = read_xlsx_as_tsv(path)
        print(f"  [1С] Прочитано из файла: {len(text)} символов.")
        return text
    finally:
        try:
            if path.exists():
                path.unlink()
        except Exception as e:
            print(f"  ⚠ Не удалось удалить временный файл {path}: {e}")


def copy_report_data(win) -> str:
    """
    Копирует все данные из сформированного отчёта 1С через Ctrl+A → Ctrl+C.
    Возвращает текст из буфера обмена (TSV).

    Ctrl+A/Ctrl+C отправляются через keybd_event (аппаратный уровень) после
    принудительного вывода окна 1С в реальный OS-фокус (_force_foreground) —
    pywinauto type_keys/send_keys до табличного документа отчёта не
    доходили, хотя клик по отчёту фокус там ставит нормально (обычный
    Ctrl+A вручную с клавиатуры у пользователя работает).

    Буфер очищается перед копированием: если Ctrl+C не сработает, в буфере
    останется "" (а не мусор от предыдущего шага — имя контрагента/дата
    периода), и вызывающий код корректно распознает это как "буфер пуст",
    а не как 0 строк, полученных при разборе чужого текста.
    """
    vi.focus(win)
    time.sleep(0.5)

    _set_clipboard_text("")

    try:
        rect = win.rectangle()
        x = (rect.right - rect.left) // 2
        y = int((rect.bottom - rect.top) * 0.65)
        _click_input_preserve_cursor(win, coords=(x, y))
        time.sleep(0.5)
    except Exception as e:
        print(f"  Не удалось кликнуть в область отчёта: {e}")

    try:
        _force_foreground(win.handle)
        time.sleep(0.3)
    except Exception as e:
        print(f"  Не удалось вывести окно на передний план: {e}")

    def _read_clipboard() -> str:
        try:
            _open_clipboard_with_retry()
        except Exception as e:
            print(f"  Буфер обмена занят: {e}")
            return None
        try:
            return win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
        except (TypeError, Exception):
            return ""
        finally:
            win32clipboard.CloseClipboard()

    # Пробуем скопировать данные несколько раз подряд — отчёт иногда ещё не
    # успевает сформироваться к первой попытке (иногда формируется быстро,
    # иногда долго — размер отчёта разный у разных контрагентов). Без
    # больших пауз, но с большим числом попыток, чтобы всё равно дождаться
    # данных.
    #
    # Для больших отчётов (десятки-сотни строк) 1С может не успеть
    # полностью дописать TSV в буфер обмена к моменту первого чтения —
    # получаем непустой, но НЕПОЛНЫЙ текст (например, обрывается на
    # середине последней строки). Поэтому после первого непустого чтения
    # ещё раз перечитываем буфер без нового Ctrl+C — если текст за это
    # время не изменился, буфер стабилен и скопирован полностью; если
    # изменился (дозаписался), ждём и сверяем снова.
    data = ""
    for attempt in range(1, 51):
        _send_ctrl_key(ord("A"))
        time.sleep(0.05)
        _send_ctrl_key(ord("C"))
        time.sleep(0.15)

        first = _read_clipboard()
        if first is None:
            print(f"  Буфер обмена занят (попытка {attempt}/50) — пробую снова...")
            continue
        if not first:
            print(f"  Буфер пуст, попытка {attempt}/50 — пробую снова...")
            continue

        # Непустой текст есть — проверяем, что он стабилен (два чтения
        # подряд дают ОДИНАКОВЫЙ результат), прежде чем ему доверять.
        stable = False
        for _ in range(6):
            time.sleep(0.15)
            second = _read_clipboard()
            if second == first:
                stable = True
                break
            print(f"  Буфер обмена ещё дозаписывается (попытка {attempt}/50) — жду стабилизации...")
            first = second

        if stable:
            data = first
            break

        # Не стабилизировался за отведённое время — пробуем скопировать
        # заново с начала (Ctrl+A → Ctrl+C), а не просто ждать ещё.
        print(f"  Копия не стабилизировалась, попытка {attempt}/50 — копирую заново...")

    if not data:
        # Ctrl+C в табличном документе 1С — не единственный способ. На
        # сервере (виртуальный ввод, отключённая сессия) он часто не
        # доносит данные вовсе, а сохранение отчёта в .xlsx тот же объём
        # отдаёт стабильно — это уже проверено на BNPL, где буфер не
        # тянул из-за размера.
        print("  Буфер пуст после копирования — пробую выгрузить отчёт в файл (Ctrl+S)...")
        try:
            data = export_report_data(win)
            if data:
                print("  Данные получены через выгрузку в файл.")
        except Exception as e:  # noqa: BLE001 — запасной путь, его отказ не важнее основного
            print(f"  Выгрузка в файл тоже не удалась: {e}")

    if not data:
        print("  Данных нет — ни через буфер обмена, ни через выгрузку в файл.")

    return data
