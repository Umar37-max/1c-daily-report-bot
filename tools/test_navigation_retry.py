"""
Проверка повторов навигации к отчёту — без 1С, на подставных функциях.

Движок навигации (src/reports.py) не меняется; проверяем обвязку вокруг:

  * gross_profit.open_report — второй заход в том же окне, и если ссылка
    отчёта уже на экране (клик сработал с опозданием) — сразу по ней;
  * daily_report.fetch_chunk — до двух перезапусков 1С, и перед ними
    перевод сессии на консоль, если клики прошли вхолостую;
  * desktop_session.ensure_interactive_desktop(force_console=True) — не
    трогает сессию на консоли и переводит RDP-сессию.

    python tools/test_navigation_retry.py
"""
import sys
from pathlib import Path
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(errors="replace")

import daily_report
from src import desktop_session, gross_profit

FAILED = 0
NO_EFFECT = RuntimeError("Не удалось нажать пункт 'Отчеты по продажам': Could not click "
                         "'Отчеты по продажам' without a cursor - tried: message: call "
                         "succeeded but had no visible effect")


def check(name, got, expected):
    global FAILED
    ok = got == expected
    FAILED += not ok
    print(f"  {'ok ' if ok else 'ОШИБКА'} {name}" + ("" if ok else f" (получил {got!r}, ждал {expected!r})"))


class FakeWin:
    def __init__(self, link_visible):
        self.link_visible = link_visible

    def child_window(self, **kwargs):
        return SimpleNamespace(exists=lambda timeout=0: self.link_visible())


def run_open_report(section_results, link_visible=lambda: False):
    """section_results — по одному на вызов движка: None (успех) или исключение."""
    calls = {"section": 0, "link": 0}
    results = list(section_results)

    def fake_section(*args, **kwargs):
        calls["section"] += 1
        outcome = results.pop(0)
        if outcome is not None:
            raise outcome

    def fake_link(*args, **kwargs):
        calls["link"] += 1

    gross_profit.open_section_report = fake_section
    gross_profit.open_report_link = fake_link
    gross_profit.dismiss_1c_message_boxes = lambda: None
    gross_profit.time.sleep = lambda s: None
    try:
        gross_profit.open_report(FakeWin(link_visible))
        calls["raised"] = False
    except Exception:  # noqa: BLE001
        calls["raised"] = True
    return calls


def run_fetch_chunk(outcomes):
    """outcomes — по одному на попытку fetch_report: строка (текст) или исключение."""
    log = {"restarts": 0, "force": []}
    results = list(outcomes)

    def fake_fetch(win, d1, d2):
        outcome = results.pop(0)
        if isinstance(outcome, Exception):
            raise outcome
        return outcome

    daily_report.fetch_report = fake_fetch
    daily_report.ensure_interactive_desktop = lambda force_console=False: log["force"].append(force_console)
    session = SimpleNamespace(win=None, restart=lambda: log.__setitem__("restarts", log["restarts"] + 1))
    try:
        log["text"] = daily_report.fetch_chunk(session, None, None, None)
    except Exception:  # noqa: BLE001
        log["text"] = "УПАЛО"
    return log


def main() -> int:
    print("── open_report: заходы в одном окне 1С ──")
    calls = run_open_report([None])
    check("всё с первого раза — один заход", (calls["section"], calls["link"], calls["raised"]), (1, 0, False))

    calls = run_open_report([NO_EFFECT, None])
    check("первый заход сорвался — второй прошёл", (calls["section"], calls["raised"]), (2, False))

    calls = run_open_report([NO_EFFECT], link_visible=lambda: True)
    check("ссылка появилась с опозданием — открыта без нового клика по меню",
          (calls["section"], calls["link"], calls["raised"]), (1, 1, False))

    calls = run_open_report([NO_EFFECT, NO_EFFECT])
    check("оба захода сорвались — ошибка наружу (fetch_chunk перезапустит 1С)",
          (calls["section"], calls["raised"]), (2, True))

    print("\n── fetch_chunk: перезапуски 1С ──")
    log = run_fetch_chunk(["TSV"])
    check("с первого раза — без перезапуска", (log["text"], log["restarts"]), ("TSV", 0))

    log = run_fetch_chunk([NO_EFFECT, "TSV"])
    check("после «no visible effect» — один перезапуск", (log["text"], log["restarts"]), ("TSV", 1))
    check("и сессия переводится на консоль", log["force"], [True])

    log = run_fetch_chunk([NO_EFFECT, NO_EFFECT, "TSV"])
    check("два сбоя подряд — третья попытка", (log["text"], log["restarts"]), ("TSV", 2))

    log = run_fetch_chunk([NO_EFFECT, NO_EFFECT, NO_EFFECT])
    check("три сбоя — ошибка наружу", (log["text"], log["restarts"]), ("УПАЛО", 2))

    log = run_fetch_chunk([ValueError("Итого не сошлось"), "TSV"])
    check("другая ошибка — перезапуск без перевода на консоль", log["force"], [False])

    print("\n── перевод сессии на консоль ──")
    tscon = []
    desktop_session._tscon = lambda sid: (tscon.append(sid), (True, ""))[1]
    desktop_session.time.sleep = lambda s: None
    desktop_session.desktop_alive = lambda sid=None: True

    desktop_session.via_rdp = lambda sid=None: False
    check("сессия на консоли — не трогаем",
          (desktop_session.ensure_interactive_desktop(force_console=True), len(tscon)), (True, 0))

    rdp = {"on": True}
    desktop_session.via_rdp = lambda sid=None: rdp["on"]

    def tscon_moves(sid):
        tscon.append(sid)
        rdp["on"] = False  # tscon перевёл сессию на консоль
        return True, ""

    desktop_session._tscon = tscon_moves
    check("RDP-сессия, клики вхолостую — переводим на консоль",
          (desktop_session.ensure_interactive_desktop(force_console=True), len(tscon)), (True, 1))

    tscon.clear()
    rdp["on"] = True
    check("RDP-сессия без сбоя кликов — не трогаем",
          (desktop_session.ensure_interactive_desktop(), len(tscon)), (True, 0))

    print("\n" + ("Все проверки прошли." if not FAILED else f"ПРОВАЛОВ: {FAILED}"))
    return 1 if FAILED else 0


if __name__ == "__main__":
    sys.exit(main())
