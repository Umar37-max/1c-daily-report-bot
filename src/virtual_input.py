# -*- coding: utf-8 -*-
"""
Cursor-free ("virtual") replacements for pywinauto's real-input calls, so this
project's 1C automation can run on a server with nobody watching - no RDP/VNC
session attached, no physical mouse, screen possibly locked.

WHY THIS EXISTS
---------------
pywinauto's `click_input()` / `pywinauto.mouse.*` / `keyboard.send_keys()` all
drive the REAL system cursor and the REAL keyboard queue (SetCursorPos +
SendInput underneath). Windows refuses those when the session has no active
desktop - exactly what happens the moment an RDP session is disconnected on
Windows Server:

    pywintypes.error: (0, 'SetCursorPos', 'No error message is available')
    RuntimeError: There is no active desktop required for moving mouse cursor!

That is a Windows restriction, not a pywinauto bug, so no other library
(pyautogui, autoclicker apps, ...) gets around it - they all call those same
two APIs. Installing a virtual display driver did not help either (confirmed
live on this project's server).

What DOES work without an active desktop is talking to the application
programmatically instead of through the OS input queue:

  1. UI Automation control patterns - Invoke / SelectionItem.Select /
     ExpandCollapse.Expand / Toggle / Value.SetValue. These ask the app to
     perform the action ("press yourself"), so no cursor is involved at all.
     1C exposes real UIA controls for buttons / tabs / list items / edits,
     which covers most of what this project clicks.
  2. LegacyIAccessible.DoDefaultAction() / .SetValue() - the MSAA bridge, for
     controls exposing no modern pattern. Known to work in 1C: this project
     already reads a combo's value through `legacy_properties()["Value"]`.
  3. Window messages - WM_LBUTTONDOWN/UP and WM_KEYDOWN/UP posted straight to
     the window handle (pywinauto's `HwndWrapper.click()` and
     `.send_keystrokes()`, both documented as working on inactive/hidden
     windows). This is the fallback for 1C's custom-drawn grids, which expose
     no UIA children at all and so have no pattern to invoke.

Every helper tries those in order and only touches real input if explicitly
allowed (see ALLOW_REAL_INPUT) - by default it never does, so a run can
neither crash on SetCursorPos nor steal a human's cursor if someone happens
to be logged in at the time.

HONEST LIMITS
-------------
Layers 1 and 2 are reliable: they are direct calls into the app.

Layer 3 is best-effort by nature. pywinauto's own docs say of
`send_keystrokes`: "each application handles input differently ... it may work
only partially depending on the target app", and it warns specifically that
key COMBINATIONS (^a, ^c, ^{ENTER}) may or may not be honoured. In this
project that matters for exactly three things, all in the custom-drawn grids:
Ctrl+A/Ctrl+C to copy a grid to the clipboard, arrow/Ctrl+Home navigation
between rows, and Ctrl+Enter to post a document. If 1C ignores posted key
messages there, those steps need a different approach entirely - the code
raises instead of silently doing nothing, and `report_capabilities()` prints
what is actually available on the machine in front of it.
"""
import ctypes
import os
import re
import time
import warnings

import win32con
import win32gui
from pywinauto import uia_defines
from pywinauto.controls.hwndwrapper import HwndWrapper

_user32 = ctypes.windll.user32
_kernel32 = ctypes.windll.kernel32   # GetCurrentThreadId lives here, not in user32
_PBYTE256 = ctypes.c_ubyte * 256

# В cashflow1c значение по умолчанию ПЕРЕВЁРНУТО по сравнению с оригиналом
# из emit_transactions: реальный ввод РАЗРЕШЁН, если он вообще возможен.
#
# Почему (живьём подтверждено на этом проекте): 1С рисует свои разделы и
# отчёты сама, дочерних hwnd под ними нет, и посланный WM_LBUTTONDOWN в
# главный фрейм часто либо не срабатывает (клик "прошёл", раздел не
# открылся), либо попадает не туда и открывает лишние окна ("Выбор поля",
# "Расшифровка"), которые модально блокируют всё остальное. Ровно к тому же
# выводу пришёл и соседний проект emit_transactions: "реальный ввод в разы
# надёжнее и проще" — его сервер работает с подключённой RDP-сессией.
#
# Поэтому: есть активный рабочий стол -> работаем реальным вводом (как
# раньше, проверено на всех отчётах этого проекта); нет (сервер без
# подключённой сессии) -> сами, без правок кода, уходим на курсор-
# независимый путь. Принудительно выключить реальный ввод:
# VIRTUAL_INPUT_ALLOW_REAL=0.
def _real_input_possible() -> bool:
    try:
        import win32api
        win32api.SetCursorPos(win32api.GetCursorPos())  # холостое движение на месте
        return True
    except Exception:
        return False


# По прямому указанию пользователя проект работает ПОЛНОСТЬЮ виртуально:
# ни мышь, ни клавиатура реально не дёргаются — всё идёт UIA-паттернами и
# посланными окну сообщениями. Это то, что нужно для сервера (там рабочего
# стола может не быть вовсе) и не мешает человеку, если он за компьютером.
# Включить реальный ввод обратно: VIRTUAL_INPUT_ALLOW_REAL=1.
ALLOW_REAL_INPUT = os.environ.get("VIRTUAL_INPUT_ALLOW_REAL", "") == "1"

# Set VIRTUAL_INPUT_DEBUG=1 to print which layer served each action - useful
# when something silently does nothing on a machine you cannot watch.
DEBUG = os.environ.get("VIRTUAL_INPUT_DEBUG", "") == "1"


class VirtualInputError(RuntimeError):
    """Raised when no cursor-free path could perform the action."""


def _log(msg):
    if DEBUG:
        print(f"   [vinput] {msg}")


# --------------------------------------------------------------- introspection

def desktop_available():
    """True if real input could work right now (i.e. this session has an
    active desktop). Diagnostics only - the virtual paths below do not care
    either way."""
    try:
        import win32api
        win32api.SetCursorPos(win32api.GetCursorPos())  # no-op move to current position
        return True
    except Exception:
        return False


def _hwnd_of(target):
    """Accept a UIA wrapper, an HwndWrapper, or a raw hwnd int."""
    if isinstance(target, int):
        return target
    hwnd = getattr(target, "handle", None)
    if hwnd:
        return hwnd
    # A UIA element that is not a top-level window has no handle of its own -
    # walk up to the nearest ancestor that does.
    try:
        parent = target.parent()
    except Exception:
        parent = None
    while parent is not None:
        hwnd = getattr(parent, "handle", None)
        if hwnd:
            return hwnd
        try:
            parent = parent.parent()
        except Exception:
            break
    return None


def _hwnd_wrapper(target):
    hwnd = _hwnd_of(target)
    if not hwnd:
        return None
    try:
        return HwndWrapper(hwnd)
    except Exception:
        return None


def _legacy_iface(ctrl):
    try:
        return uia_defines.get_elem_interface(ctrl.element_info.element, "LegacyIAccessible")
    except Exception:
        return None


def _raise_no_legacy():
    raise RuntimeError("no LegacyIAccessible interface")


def _client_center(ctrl, hwnd):
    """The control's centre point in hwnd's client coordinates, which is what
    WM_LBUTTONDOWN expects."""
    try:
        r = ctrl.rectangle()
        sx, sy = (r.left + r.right) // 2, (r.top + r.bottom) // 2
        return win32gui.ScreenToClient(hwnd, (sx, sy))
    except Exception:
        return None


def hwnd_at_point(x, y, fallback=None):
    """The DEEPEST child window at a screen point.

    Mouse messages are delivered per-HWND, so posting WM_LBUTTONDOWN to a
    top-level window when the thing under the cursor is actually one of its
    child windows means the click lands nowhere - the parent has no handler
    for that point, and the child never hears about it. That produces
    exactly the symptom seen live: the click "succeeds" (PostMessage
    returns fine) and nothing happens.

    Under a UIA element this matters because element.rectangle() gives
    SCREEN coordinates of the visual, which says nothing about which HWND
    actually owns those pixels. WindowFromPoint + ChildWindowFromPointEx
    answer that, so the click can be posted where it will actually be
    handled."""
    try:
        hwnd = win32gui.WindowFromPoint((x, y))
    except Exception:
        return fallback
    if not hwnd:
        return fallback

    # WindowFromPoint stops at the top-level window for some hit-test
    # results, so keep descending explicitly.
    flags = win32con.CWP_SKIPINVISIBLE | win32con.CWP_SKIPTRANSPARENT
    for _ in range(16):   # bounded: a pathological tree must not spin forever
        try:
            cx, cy = win32gui.ScreenToClient(hwnd, (x, y))
            child = win32gui.ChildWindowFromPointEx(hwnd, (cx, cy), flags)
        except Exception:
            break
        if not child or child == hwnd:
            break
        hwnd = child
    return hwnd


def _describe(ctrl):
    try:
        return ctrl.window_text()
    except Exception:
        return "<unreadable>"


# ---------------------------------------------------------------------- clicks

def _layers(ctrl, hwnd=None):
    """The cursor-free ways to press a control, cheapest/most reliable first.
    Yields (name, callable) pairs; the callable raises if that layer isn't
    available on this control."""

    def do_invoke():
        ctrl.invoke()

    def do_select():
        ctrl.select()

    def do_expand():
        ctrl.expand()

    def do_legacy():
        iface = _legacy_iface(ctrl)
        if iface is None:
            raise RuntimeError("no LegacyIAccessible interface")
        iface.DoDefaultAction()

    def _message_click_at_element(use_hit_test):
        r = ctrl.rectangle()
        sx, sy = (r.left + r.right) // 2, (r.top + r.bottom) // 2
        if use_hit_test:
            target_hwnd = hwnd_at_point(sx, sy, fallback=hwnd or _hwnd_of(ctrl))
        else:
            target_hwnd = hwnd or _hwnd_of(ctrl)
        if not target_hwnd:
            raise RuntimeError("no window handle to post to")
        wrapper = _hwnd_wrapper(target_hwnd)
        if wrapper is None:
            raise RuntimeError(f"could not wrap hwnd {target_hwnd}")
        coords = win32gui.ScreenToClient(target_hwnd, (sx, sy))
        _log(f"message click on hwnd {target_hwnd} "
             f"({win32gui.GetClassName(target_hwnd)!r}) at client {coords}")
        wrapper.click(coords=coords)

    def do_message_hit_test():
        # Post to whichever CHILD window actually owns the pixels under the
        # element - see hwnd_at_point. Tried before the plain variant
        # because posting to the top-level window when a child owns the
        # point is silently dropped.
        _message_click_at_element(use_hit_test=True)

    def do_message():
        _message_click_at_element(use_hit_test=False)

    def do_real():
        ctrl.click_input()

    # Порядок в cashflow1c: сначала то, что ближе всего к НАСТОЯЩЕМУ клику.
    #
    # Живьём подтверждено на 1С: UIA-паттерны (invoke/select) у её кнопок
    # регулярно отрабатывают БЕЗ ОШИБКИ, ничего при этом не нажимая — и
    # клик, у которого нет проверки результата, на них и останавливался
    # ("Настройки...", "Расширенный" и т.п. считались нажатыми, а диалог не
    # открывался). Посланное WM_LBUTTONDOWN/UP в тот hwnd, которому реально
    # принадлежат пиксели элемента, ведёт себя как обычный клик мышью,
    # поэтому идёт первым; паттерны остаются запасными (для части
    # контролов работают именно они).
    if ALLOW_REAL_INPUT:
        yield "real", do_real
    yield "message-child", do_message_hit_test   # posted to the child HWND under the point
    yield "message", do_message    # posted to the top-level window
    yield "invoke", do_invoke      # InvokePattern: buttons, hyperlinks, menu items
    yield "select", do_select      # SelectionItemPattern: tab items, list items
    yield "expand", do_expand      # ExpandCollapsePattern: tab/combo controls that open a flyout submenu
                                    # rather than directly invoking or merely selecting themselves
    yield "legacy", do_legacy      # MSAA bridge - the only pattern 1C's buttons expose


def click(ctrl, settle=0.3, hwnd=None, verify=None, verify_timeout=3.0, verify_poll=0.25):
    """Press/activate a UIA control without moving the cursor.

    Walks the cursor-free layers in order (Invoke -> SelectionItem ->
    LegacyIAccessible.DoDefaultAction -> posted WM_LBUTTONDOWN/UP) and returns
    the name of the one that worked.

    `verify` is a zero-argument callable returning truthy once the click has
    actually taken effect (e.g. "the report tab is now open"). Pass it
    wherever the effect is checkable, because a layer succeeding is NOT proof
    the click did anything: confirmed live that 1C's buttons expose no
    InvokePattern at all, and that LegacyIAccessible.DoDefaultAction() can
    return cleanly while doing nothing. With `verify`, such a silent no-op
    falls through to the next layer instead of being mistaken for success.

    Without `verify` the old behaviour stands - first layer that doesn't raise
    wins - which is the best that can be done when the effect isn't
    observable."""
    errors = []

    for name, action in _layers(ctrl, hwnd):
        try:
            action()
        except Exception as e:
            errors.append(f"{name}: {e}")
            continue

        if verify is None:
            _log(f"{name} -> {_describe(ctrl)!r}")
            time.sleep(settle)
            return name

        if _wait_for(verify, verify_timeout, verify_poll):
            _log(f"{name} -> {_describe(ctrl)!r} (effect confirmed)")
            return name

        errors.append(f"{name}: call succeeded but had no visible effect")
        _log(f"{name} on {_describe(ctrl)!r} did nothing - trying the next layer")

    raise VirtualInputError(
        f"Could not click {_describe(ctrl)!r} without a cursor - tried: {'; '.join(errors)}"
    )


def _wait_for(predicate, timeout, poll):
    end = time.time() + timeout
    while True:
        try:
            if predicate():
                return True
        except Exception:
            pass
        if time.time() >= end:
            return False
        time.sleep(poll)


def click_at(target, x, y, double=False, settle=0.3, hit_test=True):
    """Click at absolute SCREEN coordinates via posted messages (no cursor).

    `target` is the window that nominally owns that point; with
    hit_test=True (the default) the click is actually posted to the deepest
    CHILD window at those coordinates instead, since that's the HWND that
    will really handle the mouse message (see hwnd_at_point)."""
    fallback = _hwnd_of(target)
    hwnd = hwnd_at_point(x, y, fallback=fallback) if hit_test else fallback
    if not hwnd:
        raise VirtualInputError("click_at: no window handle to post the click to")
    wrapper = _hwnd_wrapper(hwnd)
    if wrapper is None:
        raise VirtualInputError(f"click_at: could not wrap hwnd {hwnd}")
    try:
        coords = win32gui.ScreenToClient(hwnd, (x, y))
    except Exception as e:
        raise VirtualInputError(f"click_at: ScreenToClient failed: {e}")
    wrapper.click(coords=coords, double=double)
    _log(f"WM_LBUTTON* {'double ' if double else ''}on hwnd {hwnd} at client {coords}")
    time.sleep(settle)
    return "message"


def double_click(ctrl, settle=0.4, hwnd=None):
    """Double-click without a cursor. Grid rows in 1C open on double click and
    there is no UIA pattern for "double invoke", so this is messages-first,
    falling back to a plain click() (for a real button a single Invoke is the
    intended action anyway)."""
    target_hwnd = hwnd or _hwnd_of(ctrl)
    if target_hwnd:
        coords = _client_center(ctrl, target_hwnd)
        wrapper = _hwnd_wrapper(target_hwnd)
        if coords and wrapper is not None:
            try:
                wrapper.click(coords=coords, double=True)
                _log(f"WM_LBUTTON* double at client {coords}")
                time.sleep(settle)
                return "message"
            except Exception as e:
                _log(f"double click via messages failed: {e}")
    return click(ctrl, settle=settle, hwnd=hwnd)


# ---------------------------------------------------------------------- clipboard

_CLIPBOARD_SENTINEL = "\x01__virtual_input_sentinel__\x01"


def copy_via_ctrl_c(target, set_clipboard_text, get_clipboard_text, settle=0.3):
    """Send Ctrl+C to `target` and return what actually landed on the
    clipboard - raising if nothing did.

    Confirmed live: a posted Ctrl+C that never actually reaches the intended
    control (wrong focus, a dialog that already closed, ...) fails
    completely silently - get_clipboard_text() then just returns whatever
    was on the clipboard from something unrelated (an OS-level file copy in
    Explorer, in one observed case), which reads exactly like a real result
    and was mistaken for one. Priming the clipboard with a sentinel value
    first turns that into a loud, immediate failure: if the sentinel is
    still there afterwards, the Ctrl+C plainly did not do anything."""
    set_clipboard_text(_CLIPBOARD_SENTINEL)
    hwnd = _hwnd_of(target)
    if not hwnd:
        raise VirtualInputError("copy_via_ctrl_c: no window handle to send Ctrl+C to")
    send_keys(hwnd, "^c", settle=settle)
    result = get_clipboard_text()
    if result == _CLIPBOARD_SENTINEL:
        raise VirtualInputError(
            "Ctrl+C did not change the clipboard - it likely never reached the intended "
            "control (wrong focus, or the target window already closed)."
        )
    return result


def read_selected_cell_text(grid):
    """Read the currently-selected item's text via UIA SelectionPattern -
    no keyboard, no clipboard, no click. Returns None if the grid exposes no
    SelectionPattern or nothing is currently selected (the grids this
    project drives are documented as exposing NO UIA children at all, which
    would likely make this unavailable too, but it costs nothing to check
    before falling back to a clipboard-based read)."""
    try:
        selected = grid.get_selection()
    except Exception as e:
        _log(f"read_selected_cell_text: SelectionPattern unavailable ({e})")
        return None
    if not selected:
        _log("read_selected_cell_text: SelectionPattern available but nothing is selected")
        return None
    try:
        return selected[0].name
    except Exception as e:
        _log(f"read_selected_cell_text: could not read selected item's name ({e})")
        return None


# ------------------------------------------------------------------------ text

def get_text(ctrl):
    """Read a field's current value, whichever interface exposes it."""
    for getter in (
        lambda: ctrl.iface_value.CurrentValue,
        lambda: ctrl.get_value(),
        lambda: (_legacy_iface(ctrl) or _raise_no_legacy()).CurrentValue,
        lambda: ctrl.window_text(),
    ):
        try:
            v = getter()
            if v is not None:
                return str(v)
        except Exception:
            continue
    return ""


def set_text(ctrl, text, verify=True, settle=0.25):
    """Replace a field's contents without cursor or keyboard.

    ValuePattern.SetValue first, then LegacyIAccessible.SetValue. Both hand
    the string straight to the control, which also sidesteps the Cyrillic
    problem the clipboard dance existed to work around.

    verify=True re-reads the field and rejects the attempt if it did not take,
    so a field that silently ignored SetValue (some masked/formatted 1C fields
    do) cannot pass for success."""
    errors = []

    for label, setter in (
        ("value pattern", lambda: ctrl.iface_value.SetValue(text)),
        ("legacy", lambda: (_legacy_iface(ctrl) or _raise_no_legacy()).SetValue(text)),
    ):
        try:
            setter()
            time.sleep(settle)
            if verify:
                actual = get_text(ctrl)
                if actual.strip() != text.strip():
                    errors.append(f"{label}: wrote {text!r} but field reads {actual!r}")
                    continue
            _log(f"{label} SetValue({text!r})")
            return label
        except Exception as e:
            errors.append(f"{label}: {e}")

    raise VirtualInputError(
        f"Could not set text {text!r} without keyboard input - tried: {'; '.join(errors)}"
    )


# -------------------------------------------------------------------- keyboard

# Combos this project uses, as (modifier VKs, main VK) - resolved from
# explicit virtual-key codes so the active keyboard layout is irrelevant.
_VK_CTRL = win32con.VK_CONTROL
_NAMED_VKS = {
    "ENTER": win32con.VK_RETURN,
    "RETURN": win32con.VK_RETURN,
    "TAB": win32con.VK_TAB,
    "ESC": win32con.VK_ESCAPE,
    "ESCAPE": win32con.VK_ESCAPE,
    "HOME": win32con.VK_HOME,
    "END": win32con.VK_END,
    "DELETE": win32con.VK_DELETE,
    "DEL": win32con.VK_DELETE,
    "BACKSPACE": win32con.VK_BACK,
    "UP": win32con.VK_UP,
    "DOWN": win32con.VK_DOWN,
    "LEFT": win32con.VK_LEFT,
    "RIGHT": win32con.VK_RIGHT,
    "PGUP": win32con.VK_PRIOR,
    "PGDN": win32con.VK_NEXT,
    "SPACE": win32con.VK_SPACE,
    # Добавлено в копии для cashflow1c: этот проект держится на Shift+Insert
    # (вставка из буфера в поля 1С — единственный надёжный способ ввода там)
    # и на F4 (открыть список в поле контрагента). Без них строки уходили бы
    # в pywinauto-парсер send_keystrokes, который переводит символы через
    # активную раскладку — ровно тот путь, который в 1С уже подводил.
    "INSERT": win32con.VK_INSERT,
    "INS": win32con.VK_INSERT,
    **{f"F{n}": getattr(win32con, f"VK_F{n}") for n in range(1, 13)},
}

# "^a", "^{ENTER}", "{DOWN 3}", "{ENTER}", "01062026" - the shapes actually
# used in this project. Модификаторы: "^" = Ctrl, "+" = Shift (в cashflow1c
# нужен Shift+Insert / Shift+End), в любом порядке и количестве.
_COMBO_RE = re.compile(r"^([\^+]*)(?:\{([A-Za-z]+\d*)(?:\s+(\d+))?\}|([A-Za-z0-9]))$")


# VK_OEM_COMMA / VK_OEM_PERIOD - not exposed by win32con, but standard,
# documented Win32 virtual-key codes for the physical ","/"." keys, in the
# same "identifies a key position, not a layout-translated character" spirit
# as VK_A..VK_Z. Needed for amounts like "139643,00" (Сумма fields use a
# comma decimal separator) - confirmed live that without this, such a string
# doesn't match the plain-run branch below, falls back to pywinauto's
# layout-translating send_keystrokes, and the field's value never actually
# changes (caught by type_digits()'s own verification, which is what
# surfaced this rather than it silently posting a wrong amount).
_VK_OEM_COMMA = 0xBC
_VK_OEM_PERIOD = 0xBE
_PUNCT_VKS = {",": _VK_OEM_COMMA, ".": _VK_OEM_PERIOD}


def _vk_for(token):
    """Virtual key for a single token, independent of keyboard layout."""
    if len(token) == 1:
        if token in _PUNCT_VKS:
            return _PUNCT_VKS[token]
        ch = token.upper()
        if "A" <= ch <= "Z" or "0" <= ch <= "9":
            return ord(ch)          # VK_A..VK_Z / VK_0..VK_9 == ord of the ASCII char
        return None
    return _NAMED_VKS.get(token.upper())


def _post_vk(hwnd, vk, down, extended=False):
    scan = _user32.MapVirtualKeyW(vk, 0)
    lparam = 1 | (scan << 16)
    if extended:
        lparam |= 1 << 24
    if not down:
        lparam |= (1 << 30) | (1 << 31)   # previous-state + transition bits
    msg = win32con.WM_KEYDOWN if down else win32con.WM_KEYUP
    # PostMessage, not SendMessage: SendMessage calls the target's WndProc
    # synchronously and makes InSendMessage() true for the duration - some
    # apps deliberately skip/short-circuit key handling in that case
    # specifically to dodge cross-thread re-entrancy, which would make a
    # synthetic keystroke silently do nothing even though delivery
    # "succeeded" (confirmed live: Ctrl+C posted this way to 1C's ledger
    # grid never actually copied anything). PostMessage queues it instead,
    # so it's picked up through the target's normal GetMessage loop like a
    # real keystroke would be.
    win32gui.PostMessage(hwnd, msg, vk, lparam)


def send_vk_combo(hwnd, vk, ctrl=False, repeat=1, settle=0.0, pump_delay=0.05, shift=False):
    """Post a keystroke (optionally with Ctrl held) to a window using explicit
    virtual-key codes - no dependence on the active keyboard layout, and no
    global input queue.

    Also mirrors the modifier into the target thread's keyboard state, because
    apps typically ask GetKeyState(VK_CONTROL) rather than tracking the
    Ctrl key-down message themselves; without this, Ctrl+C arrives looking
    like a plain 'C'.

    `pump_delay` matters more than it looks: PostMessage only QUEUES the
    keystrokes - the target thread's own message loop has to actually run to
    process them, on its own schedule, not ours. Confirmed live: restoring
    the keyboard state and detaching input immediately after posting (no
    delay at all) reliably beat the target's message loop to it, so by the
    time it actually handled the letter keydown, Ctrl had already been
    "released" again - Ctrl+V landed as a plain V (translated through the
    active layout: 'фмфмфмм' accumulating in a 1C search field instead of
    ever pasting "EMIT PINFL"). This delay gives the target a chance to pump
    and process the posted messages while Ctrl is still held in the shared
    state, before we tear that state back down."""
    target_tid = _user32.GetWindowThreadProcessId(hwnd, None)
    current_tid = _kernel32.GetCurrentThreadId()
    attached = bool(_user32.AttachThreadInput(target_tid, current_tid, True))

    saved = _PBYTE256()
    _user32.GetKeyboardState(ctypes.byref(saved))
    try:
        if ctrl or shift:
            state = _PBYTE256.from_buffer_copy(saved)
            if ctrl:
                state[_VK_CTRL] |= 128
            if shift:
                state[win32con.VK_SHIFT] |= 128
            _user32.SetKeyboardState(ctypes.byref(state))
            if ctrl:
                _post_vk(hwnd, _VK_CTRL, down=True)
            if shift:
                _post_vk(hwnd, win32con.VK_SHIFT, down=True)

        for _ in range(repeat):
            _post_vk(hwnd, vk, down=True)
            _post_vk(hwnd, vk, down=False)

        if shift:
            _post_vk(hwnd, win32con.VK_SHIFT, down=False)
        if ctrl:
            _post_vk(hwnd, _VK_CTRL, down=False)

        if pump_delay:
            time.sleep(pump_delay)
    finally:
        _user32.SetKeyboardState(ctypes.byref(saved))
        if attached:
            _user32.AttachThreadInput(target_tid, current_tid, False)

    if settle:
        time.sleep(settle)
    return True


def _send_via_vks(hwnd, keys):
    """Try to express `keys` purely as explicit virtual keys. Returns True if
    it was sent that way, False if the string is something this can't model
    (caller then falls back to pywinauto's parser)."""
    m = _COMBO_RE.match(keys.strip())
    if m:
        mods, named, count, single = m.groups()
        token = named or single
        vk = _vk_for(token)
        if vk is not None:
            send_vk_combo(hwnd, vk, ctrl="^" in (mods or ""), shift="+" in (mods or ""),
                          repeat=int(count or 1))
            return True
        return False

    # A plain run of ASCII letters/digits/,./  (e.g. a date typed as
    # "06012026", or an amount like "139643,00").
    if keys and all(("0" <= c <= "9") or ("a" <= c.lower() <= "z") or c in _PUNCT_VKS for c in keys):
        for ch in keys:
            send_vk_combo(hwnd, _vk_for(ch))
        return True
    return False


def _force_foreground(hwnd, retries=10, poll=0.15):
    """Reliably make hwnd the real foreground window - plain
    SetForegroundWindow is often silently refused by Windows when the
    calling process isn't the one currently "owning" input attribution
    (confirmed live: a bare SetForegroundWindow call here still left real
    keyboard input going wherever the actual foreground window already
    was, not to 1C - the ALLOW_REAL_INPUT fallback appeared to run but had
    no effect for exactly this reason).

    The documented workaround is AttachThreadInput: temporarily attach this
    thread's input queue to both the current foreground thread's and the
    target's, which grants permission to actually change it. Duplicated
    from lib1c.force_foreground() rather than imported from it - lib1c
    already imports this module, and importing back would be circular."""
    if _user32.IsIconic(hwnd):
        _user32.ShowWindow(hwnd, 9)  # SW_RESTORE
    for _ in range(retries):
        fg = _user32.GetForegroundWindow()
        if fg == hwnd:
            return True
        our_tid = _kernel32.GetCurrentThreadId()
        fg_tid = _user32.GetWindowThreadProcessId(fg, None) if fg else 0
        target_tid = _user32.GetWindowThreadProcessId(hwnd, None)
        attached_fg = attached_target = False
        try:
            if fg_tid and fg_tid != our_tid:
                attached_fg = bool(_user32.AttachThreadInput(our_tid, fg_tid, True))
            if target_tid and target_tid != our_tid and target_tid != fg_tid:
                attached_target = bool(_user32.AttachThreadInput(our_tid, target_tid, True))
            _user32.SetForegroundWindow(hwnd)
        finally:
            if attached_fg:
                _user32.AttachThreadInput(our_tid, fg_tid, False)
            if attached_target:
                _user32.AttachThreadInput(our_tid, target_tid, False)
        if _user32.GetForegroundWindow() == hwnd:
            return True
        time.sleep(poll)
    return _user32.GetForegroundWindow() == hwnd


def _real_send_keys(hwnd, keys):
    """Real keyboard input (SendInput, via pywinauto.keyboard) - only used
    when ALLOW_REAL_INPUT is set. `keys` uses the exact same syntax
    (pywinauto's own), so callers never need to know which path ran.

    Real input goes to whatever window has actual OS focus, not to a
    specific hwnd, so this brings the target forward first - via
    _force_foreground, not a bare SetForegroundWindow (see its docstring
    for why the bare call isn't enough on its own). Confirmed live that
    message-posted keys (PostMessage, correct VK codes, correct
    modifier-key state) do not reliably trigger 1C's own keyboard handling
    for at least Ctrl+C and RIGHT/Enter grid navigation - clicks (see
    _layers' hit-tested "message-child") work fine posted, but these
    specific keyboard interactions apparently need the real input path.
    Needs an active desktop, same as click()'s "real" layer."""
    if not _force_foreground(hwnd):
        raise RuntimeError(f"could not bring hwnd {hwnd} to the real foreground")
    time.sleep(0.05)
    from pywinauto.keyboard import send_keys as _pywinauto_real_send_keys
    _pywinauto_real_send_keys(keys)


def send_keys(target, keys, settle=0.2, required=True):
    """Send keystrokes to a window.

    With ALLOW_REAL_INPUT set, tries REAL keyboard input first (see
    _real_send_keys) - message-posted keys have proven unreliable for
    specific 1C interactions (Ctrl+C, grid navigation) even with correct
    VK codes and modifier state, so when an active desktop is available and
    explicitly allowed, prefer the path known to actually work. Falls back
    to the message-based paths below if real input itself fails for some
    reason (e.g. SetForegroundWindow refused).

    Message-based paths, in order:

    1. Explicit virtual-key codes (`_send_via_vks`) for the shapes this
       project uses - "^a", "^c", "^v", "^{ENTER}", "^{HOME}", "{ENTER}",
       "{TAB}", "{ESC}", "{DELETE}", "{DOWN 3}", "{RIGHT 3}", and plain
       digit/letter runs. Layout-independent, which matters: see this
       module's note on VkKeyScanExW returning -1 for Latin letters under a
       Russian layout, which silently turns Ctrl+A into Ctrl+<vk 255>.
    2. pywinauto's `send_keystrokes` for anything else - also message-based
       ("Silently send keystrokes to the control in an inactive window"), but
       it does translate characters through the active layout.

    required=False turns a failure into a False return instead of an
    exception, for keystrokes that are only a nice-to-have."""
    hwnd = _hwnd_of(target)
    if not hwnd:
        if required:
            raise VirtualInputError(f"send_keys({keys!r}): no window handle to post to")
        return False

    if ALLOW_REAL_INPUT:
        try:
            _real_send_keys(hwnd, keys)
            _log(f"real send_keys({keys!r}) -> hwnd {hwnd}")
            time.sleep(settle)
            return True
        except Exception as e:
            _log(f"real send_keys({keys!r}) failed ({e}) - falling back to message-based")

    try:
        if _send_via_vks(hwnd, keys):
            _log(f"send_vk({keys!r}) -> hwnd {hwnd} (explicit VKs, layout-independent)")
            time.sleep(settle)
            return True
    except Exception as e:
        if required:
            raise VirtualInputError(f"send_keys({keys!r}) failed: {e}")
        _log(f"send_vk({keys!r}) failed: {e}")
        return False

    wrapper = _hwnd_wrapper(hwnd)
    if wrapper is None:
        if required:
            raise VirtualInputError(f"send_keys({keys!r}): could not wrap hwnd {hwnd}")
        return False
    try:
        with warnings.catch_warnings():
            # send_keystrokes warns on every key combination and on thread
            # attach; both are expected here and would otherwise bury the real
            # log output.
            warnings.simplefilter("ignore")
            wrapper.send_keystrokes(keys)
        _log(f"send_keystrokes({keys!r}) -> hwnd {hwnd} (pywinauto parser)")
        time.sleep(settle)
        return True
    except Exception as e:
        if required:
            raise VirtualInputError(f"send_keys({keys!r}) failed: {e}")
        _log(f"send_keystrokes({keys!r}) failed: {e}")
        return False


def focus(ctrl):
    """Give a control keyboard focus programmatically (UIA SetFocus - no
    cursor). Needed before send_keys for anything that routes keys by focus."""
    try:
        ctrl.set_focus()
        time.sleep(0.15)
        return True
    except Exception as e:
        _log(f"set_focus failed: {e}")
        return False


# ------------------------------------------------------------------ diagnostics

def report_capabilities(win):
    """Print which virtual paths are actually available on the window in front
    of us. Cheap, read-only, and worth logging once per run: on a machine
    nobody is watching, "the click did nothing" is otherwise indistinguishable
    from "the click never happened"."""
    lines = ["=== virtual input capabilities ==="]
    lines.append(f"active desktop present: {desktop_available()} "
                 "(False is fine - virtual input does not need one)")
    lines.append(f"real-input fallback allowed: {ALLOW_REAL_INPUT}")
    lines.append(f"target hwnd: {_hwnd_of(win)}")

    counts = {}
    for ct in ("Button", "TabItem", "ListItem", "Edit", "ComboBox", "DataGrid", "Hyperlink"):
        try:
            counts[ct] = len(win.descendants(control_type=ct))
        except Exception:
            counts[ct] = "?"
    lines.append(f"controls visible to UIA: {counts}")

    # Which screen 1C is actually on, and what can be navigated from it -
    # the first thing worth knowing when a click "worked" but nothing moved.
    try:
        panes = [p.window_text().strip() for p in win.descendants(control_type="Pane")
                 if p.window_text().strip()]
        lines.append(f"active pane: {panes[0]!r}" if panes else "active pane: <none>")
    except Exception as e:
        lines.append(f"active pane: unreadable ({e})")
    try:
        tabs = sorted({t.window_text().strip() for t in win.descendants(control_type="TabItem")
                       if t.window_text().strip()})
        lines.append(f"section tabs: {tabs}")
    except Exception as e:
        lines.append(f"section tabs: unreadable ({e})")

    probe = None
    try:
        buttons = win.descendants(control_type="Button")
        probe = buttons[0] if buttons else None
    except Exception:
        pass
    if probe is not None:
        try:
            lines.append(f"InvokePattern on a sample button: {probe.iface_invoke is not None}")
        except Exception as e:
            lines.append(f"InvokePattern on a sample button: no ({e})")
        lines.append(f"LegacyIAccessible on a sample button: {_legacy_iface(probe) is not None}")
    print("\n".join(lines))
