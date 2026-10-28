"""
Проверка прав в боте: кто что может.

  * администратор (BOT_ADMIN_IDS) — запускает заливку и всё остальное;
  * наблюдатель (BOT_VIEWER_IDS, по умолчанию получатели рассылки) —
    только «Скриншот» и «Статус»;
  * посторонний — ничего.

    python tools/test_roles.py
"""
import sys
from pathlib import Path
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(errors="replace")

import bot


def _update(user_id):
    return SimpleNamespace(effective_user=SimpleNamespace(id=user_id, full_name="Тест"))


def main() -> int:
    admin = bot.CFG.admin_ids[0] if bot.CFG.admin_ids else 1
    viewer = next((v for v in bot.CFG.viewer_ids if v not in bot.CFG.admin_ids), None)
    stranger = 999_999_999

    failed = 0
    print(f"админ={admin}  наблюдатель={viewer}  посторонний={stranger}\n")

    cases = [
        ("админ запускает заливку", bot.is_admin(_update(admin)), True),
        ("админ видит отчёт", bot.is_viewer(_update(admin)), True),
        ("посторонний не запускает", bot.is_admin(_update(stranger)), False),
        ("посторонний не видит отчёт", bot.is_viewer(_update(stranger)), False),
    ]
    if viewer is not None:
        cases += [
            ("наблюдатель НЕ запускает заливку", bot.is_admin(_update(viewer)), False),
            ("наблюдатель видит отчёт", bot.is_viewer(_update(viewer)), True),
        ]
    else:
        print("  (в .env нет отдельного наблюдателя — проверяю только админа и постороннего)\n")

    for name, got, expected in cases:
        ok = got is expected
        failed += not ok
        print(f"  {'ok ' if ok else 'ОШИБКА'} {name}: {got}")

    print("\n── меню по ролям ──")
    admin_menu = [b.text for row in bot.menu_for(_update(admin)).keyboard for b in row]
    ok = bot.BTN_RUN in admin_menu
    failed += not ok
    print(f"  {'ok ' if ok else 'ОШИБКА'} у админа есть «{bot.BTN_RUN}»")

    if viewer is not None:
        viewer_menu = [b.text for row in bot.menu_for(_update(viewer)).keyboard for b in row]
        checks = [
            ("нет кнопки заливки", bot.BTN_RUN not in viewer_menu),
            ("нет перезалива недели", bot.BTN_WEEK not in viewer_menu),
            ("нет периода", bot.BTN_RANGE not in viewer_menu),
            ("есть скриншот", bot.BTN_SHOT in viewer_menu),
            ("есть статус", bot.BTN_STATUS in viewer_menu),
        ]
        for name, ok in checks:
            failed += not ok
            print(f"  {'ok ' if ok else 'ОШИБКА'} у наблюдателя {name}")

    ok = bot.menu_for(_update(stranger)) is None
    failed += not ok
    print(f"  {'ok ' if ok else 'ОШИБКА'} у постороннего меню нет")

    print("\n" + ("Все проверки прошли." if not failed else f"ПРОВАЛОВ: {failed}"))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
