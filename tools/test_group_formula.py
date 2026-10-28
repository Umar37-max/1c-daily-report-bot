"""
Проверка правки формулы «Группа» (J) — на формуле той же структуры, что в книге
(названия магазинов и партнёров заменены на условные).

Формула определяет канал продаж; всё, чего она не знает, остаётся пустым,
и прогон удаляет такую строку. Проверяем, что:

  * новое подразделение попадает в ОБА перечня розницы (BNPL и обычный);
  * старые магазины и прочие ветки остаются нетронутыми;
  * повторный прогон ничего не меняет (можно звать каждый раз);
  * ветка маркетплейса D добавляется, если её нет, и не дублируется, если есть;
  * формула с запятыми-разделителями (другая локаль) тоже чинится.

    python tools/test_group_formula.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(errors="replace")

from src.sheets_register import add_marketplace_branch, add_retail_divisions

ROW = 18489

# Структура формулы из книги-регистра, строка 18489 (значение на 06.10.2026).
FORMULA = f'''=IF(K{ROW}="BNPL";
    IF(OR(D{ROW}="Магазин 1";D{ROW}="Магазин 2";D{ROW}="Магазин 3";D{ROW}="Магазин 4";D{ROW}="Магазин 5";D{ROW}="Магазин 6");
        "Розница";
        IF(OR(D{ROW}="Интернет магазин";D{ROW}="Онлайн рассрочка";D{ROW}="Свойская рассрочка");
            "Интернет Магазин";
            ""
        )
    );
    IF(H{ROW}="Partner A Marketplace";
        "Партнёр A";
        IF(H{ROW}="Partner B marketplace";
            "Партнёр B";
            IF(H{ROW}="Partner C Marketplace";
                "Партнёр C";
                IF(H{ROW}="UZS 30 days";
                    "Optom 2";
                    IF(H{ROW}="Marketplace";
                        IF(A{ROW}="MARKETPLACE A";
                            "Маркетплейс A";
                            IF(OR(A{ROW}="MARKETPLACE B";A{ROW}="MARKETPLACE B Partner";A{ROW}="MARKETPLACE B Retail");
                                "Маркетплейс B";
                                IF(A{ROW}="MARKETPLACE C";
                                    "Маркетплейс C";IF(A{ROW}="MARKETPLACE D";"Маркетплейс D";"")
                                )
                            )
                        );
                        IF(OR(D{ROW}="Магазин 1";D{ROW}="Магазин 2";D{ROW}="Магазин 3";D{ROW}="Магазин 4";D{ROW}="Магазин 5";D{ROW}="Магазин 6");
                            "Розница";
                            IF(OR(D{ROW}="Интернет магазин";D{ROW}="Онлайн рассрочка";D{ROW}="Свойская рассрочка");
                                "Интернет Магазин";
                                ""
                            )
                        )
                    )
                )
            )
        )
    )
)'''

NEW_STORE = "Магазин 7"
OTHER_NEW_STORE = "Магазин 8"

FAILED = 0


def check(name, got, expected=True):
    global FAILED
    ok = got == expected
    FAILED += not ok
    print(f"  {'ok ' if ok else 'ОШИБКА'} {name}" + ("" if ok else f" (получил {got!r})"))


def main() -> int:
    print(f"── {NEW_STORE} в перечне розницы ──")
    patched = add_retail_divisions(FORMULA, ROW, [NEW_STORE])
    check("добавлен в оба списка розницы", patched.count(f'D{ROW}="{NEW_STORE}"'), 2)
    check("списков розницы по-прежнему два", patched.count('"Розница"'), 2)
    check("Магазин 1 на месте", patched.count(f'D{ROW}="Магазин 1"'), 2)
    check("Магазин 6 на месте", patched.count(f'D{ROW}="Магазин 6"'), 2)
    check("в «Интернет Магазин» не залез", patched.count(f'D{ROW}="Интернет магазин"'), 2)
    check("ветка маркетплейса D не тронута", patched.count('"Маркетплейс D"'), 1)
    check("скобки сошлись", patched.count("(") == patched.count(")"))
    check("длина выросла ровно на две вставки",
          len(patched) - len(FORMULA), 2 * len(f';D{ROW}="{NEW_STORE}"'))

    print("\n── повторный вызов ничего не портит ──")
    twice = add_retail_divisions(patched, ROW, [NEW_STORE])
    check("формула не изменилась", twice == patched)

    print("\n── несколько подразделений сразу ──")
    multi = add_retail_divisions(FORMULA, ROW, [NEW_STORE, OTHER_NEW_STORE])
    check(f"{NEW_STORE} добавлен", multi.count(f'D{ROW}="{NEW_STORE}"'), 2)
    check(f"{OTHER_NEW_STORE} добавлен", multi.count(f'D{ROW}="{OTHER_NEW_STORE}"'), 2)

    print("\n── пустой список ──")
    check("формула без изменений", add_retail_divisions(FORMULA, ROW, []) == FORMULA)
    check("пробелы-пустышки игнорируются",
          add_retail_divisions(FORMULA, ROW, ["", "  "]) == FORMULA)

    print("\n── локаль с запятыми ──")
    comma = FORMULA.replace(";", ",")
    comma_patched = add_retail_divisions(comma, ROW, [NEW_STORE])
    check("разделитель взят из формулы", comma_patched.count(f',D{ROW}="{NEW_STORE}"'), 2)
    check("точку с запятой не принёс", ";" not in comma_patched)

    print("\n── ветка маркетплейса D ──")
    without_d = FORMULA.replace(f'"Маркетплейс C";IF(A{ROW}="MARKETPLACE D";"Маркетплейс D";"")',
                                '"Маркетплейс C";""')
    check("в заготовке маркетплейса D действительно нет", '"Маркетплейс D"' not in without_d)
    with_d = add_marketplace_branch(without_d, ROW)
    check("маркетплейс D добавлен", with_d.count(f'A{ROW}="MARKETPLACE D"'), 1)
    check("скобки сошлись", with_d.count("(") == with_d.count(")"))
    check("на готовой формуле ничего не меняет", add_marketplace_branch(FORMULA, ROW) == FORMULA)

    print("\n── вместе, как в прогоне ──")
    both = add_retail_divisions(add_marketplace_branch(without_d, ROW), ROW, [NEW_STORE])
    check(f"есть и маркетплейс D, и {NEW_STORE}",
          '"Маркетплейс D"' in both and both.count(f'D{ROW}="{NEW_STORE}"') == 2)
    check("скобки сошлись", both.count("(") == both.count(")"))

    print("\n" + ("Все проверки прошли." if not FAILED else f"ПРОВАЛОВ: {FAILED}"))
    return 1 if FAILED else 0


if __name__ == "__main__":
    sys.exit(main())
