"""
Отчёт «Валовая прибыль организаций - копия» (раздел Продажи → Отчеты по
продажам) — источник строк для вкладки «Регистр» книги Sales by day.

Навигация ровно та же, что руками: TabItem «Продажи» → MenuItem «Отчеты по
продажам» → Hyperlink с ТОЧНЫМ названием отчёта. Точность названия здесь
критична: в списке отчётов рядом лежат «Валовая прибыль организаций -»,
«... - Cross_Sale», «... - <сотрудник>», «... - Marketing»,
«... - Рассрочка» — открыть не тот вариант значит молча взять чужие
настройки отбора.

Настройки отбора в самом отчёте уже сохранены пользователем (галки
«Подразделение» и «Клиент.Группа доступа» сняты) — их НЕ трогаем, меняем
только период.

Структура выгрузки (TSV, 20 колонок) — ровно порядок столбцов листа
«Регистр», где A..I идут из отчёта, J/K — формулы книги, L..V снова из
отчёта:
  0  Клиент                                → A
  1  Менеджер                              → B
  2  Серия                                 → C
  3  Подразделение                         → D
  4  Заказ клиента / Реализация.Номер      → E
  5  Имей                                  → F
  6  Номенклатура.Марка (бренд)            → G
  7  Заказ клиента / Реализация.Соглашение → H
  8  Период                                → I
  9  Клиент.Группа доступа                 → L
  10 Склад                                 → M
  11 Номенклатура                          → N
  12 Вид номенклатуры                      → O
  13 Количество                            → P
  14 Стоимость закупки                     → Q
  15 Доп.расходы                           → R
  16 Выручка                               → S
  17 Сумма авт. скидки                     → T
  18 Валовая прибыль                       → U
  19 Рентабельность %                      → V
"""
import re
import time
from datetime import date

from src.reports import (
    copy_report_data,
    dismiss_1c_message_boxes,
    export_report_data,
    generate,
    open_report_link,
    open_section_report,
    set_period,
)

REPORT_SECTION = "Продажи"
REPORT_MENU_ITEM = "Отчеты по продажам"
REPORT_TITLE = "Валовая прибыль организаций - копия"

COLUMNS = [
    "Клиент", "Менеджер", "Серия", "Подразделение",
    "Заказ клиента / Реализация.Номер", "Имей", "Номенклатура.Марка (бренд)",
    "Заказ клиента / Реализация.Соглашение", "Период",
    "Клиент.Группа доступа", "Склад", "Номенклатура", "Вид номенклатуры",
    "Количество", "Стоимость закупки", "Доп.расходы", "Выручка",
    "Сумма авт. скидки", "Валовая прибыль", "Рентабельность %",
]
COLUMN_COUNT = len(COLUMNS)

# Индекс колонки «Период» в выгрузке — по ней опознаём строки данных.
PERIOD_IDX = 8

# Колонки-числа (Количество .. Рентабельность %).
NUMERIC_IDX = {13, 14, 15, 16, 17, 18, 19}

_DATE_RE = re.compile(r"^\s*(\d{1,2})[./](\d{1,2})[./](\d{4})")
_PARAMS_PERIOD_RE = re.compile(
    r"Период:\s*(\d{1,2})[./](\d{1,2})[./](\d{4})\s*-\s*(\d{1,2})[./](\d{1,2})[./](\d{4})"
)
_SPACES_RE = re.compile(r"[\s\u00a0\u202f']+")


def fmt(d: date) -> str:
    return d.strftime("%d.%m.%Y")


OPEN_ROUNDS = 2          # заходов навигации в одном окне 1С
OPEN_ROUND_PAUSE = 6.0   # секунд между заходами — дать 1С догрузить раздел


def open_report(win, rounds: int = OPEN_ROUNDS) -> None:
    """
    Открывает отчёт заново (каждый вызов — новая вкладка, старые не мешают).

    Сама навигация — движок src/reports.py (open_section_report), он не
    меняется. Здесь только второй заход поверх него.

    ЖИВЬЁМ 28.10.2026: клик по «Отчеты по продажам» иногда проходит
    вхолостую все три попытки движка, а иногда срабатывает со второй.
    Раздел 1С подгружается с сервера (ws://erp1…), и ссылки отчётов
    появляются с опозданием — уже ПОСЛЕ того, как движок сдался. Поэтому:
    пауза, и если ссылка отчёта к этому времени уже на экране — открываем
    её сразу, не кликая по меню ещё раз; если нет — проходим навигацию
    заново.
    """
    link = win.child_window(title=REPORT_TITLE, control_type="Hyperlink")
    for round_no in range(1, rounds + 1):
        try:
            if round_no > 1 and link.exists(timeout=1):
                print(f"  [1С] Ссылка «{REPORT_TITLE}» уже на экране — открываю её.")
                open_report_link(win, link, wait_after=4)
            else:
                open_section_report(win, REPORT_SECTION, REPORT_MENU_ITEM, REPORT_TITLE, wait_after=4)
            return
        except Exception as e:  # noqa: BLE001 — ещё один заход в том же окне
            if round_no == rounds:
                raise
            print(f"  [1С] Отчёт не открылся (заход {round_no}/{rounds}): {e}")
            print(f"  [1С] Жду {OPEN_ROUND_PAUSE:.0f} c, пока 1С догрузит раздел, и захожу ещё раз...")
            dismiss_1c_message_boxes()
            time.sleep(OPEN_ROUND_PAUSE)


def fetch_report(win, date_from: date, date_to: date, prefer_file: bool = False) -> str:
    """
    Открывает отчёт, ставит период, формирует и отдаёт TSV-текст.

    Данные забираем буфером обмена (Ctrl+A/Ctrl+C): copy_report_data сам
    дожидается, пока буфер перестанет дозаписываться, и сам падает в
    выгрузку файлом, если буфер так и остался пустым.

    prefer_file=True пробует сначала Ctrl+S → .xlsx. ЖИВЬЁМ (20.09.2026) на
    этой машине диалог «Сохранение» виртуальным вводом не нажимается
    («Событие не смогло вызвать ни одного из абонентов»), окно закрывается
    по Escape, и весь путь бесполезно тратит ~40 секунд — поэтому по
    умолчанию выключено. Оставлено для отчётов на тысячи строк, где буфер
    может не вытянуть.
    """
    open_report(win)
    set_period(win, fmt(date_from), fmt(date_to))
    generate(win, wait_seconds=1.0)

    text = ""
    if prefer_file:
        try:
            text = export_report_data(win)
        except Exception as e:  # noqa: BLE001 — есть запасной путь через буфер
            print(f"  ! Выгрузка в файл не удалась ({e}) — пробую буфер обмена.")
    if not text.strip():
        text = copy_report_data(win)
    if not text.strip():
        raise RuntimeError("Отчёт не удалось получить ни файлом, ни через буфер обмена.")

    verify_period(text, date_from, date_to)
    return text


def verify_period(text: str, date_from: date, date_to: date) -> None:
    """
    Сверяет период в шапке САМОГО отчёта («Параметры: Период: ... - ...») с
    тем, который мы просили.

    Это единственная по-настоящему надёжная проверка: поле периода умеет
    показывать введённую дату, а отчёт при этом сформироваться за старый
    период (живьём подтверждено в соседнем проекте emit_transactions). Если
    шапку не нашли — только предупреждаем: у отчёта её можно отключить, и
    падать из-за этого нельзя.
    """
    m = _PARAMS_PERIOD_RE.search(text)
    if not m:
        print("  ! В шапке отчёта нет строки «Период:» — сверить период не с чем.")
        return
    d1, m1, y1, d2, m2, y2 = (int(g) for g in m.groups())
    want = (date_from, date_to)
    variants = []
    for day1, mon1, day2, mon2 in ((d1, m1, d2, m2), (m1, d1, m2, d2)):
        try:
            variants.append((date(y1, mon1, day1), date(y2, mon2, day2)))
        except ValueError:
            continue
    if want not in variants:
        raise RuntimeError(
            f"Отчёт сформирован за период {variants[0][0]}–{variants[0][1]}, "
            f"а просили {want[0]}–{want[1]}. Дальше не иду, чтобы не записать "
            "в регистр чужие дни."
        )
    print(f"  Период в шапке отчёта совпал с запрошенным: {fmt(date_from)} – {fmt(date_to)}")


def detect_decimal_sep(text: str) -> str:
    """
    Какой символ в ЭТОЙ выгрузке десятичный: ',' или '.'.

    Формат чисел задаёт не наш код, а региональные настройки машины, где
    работает 1С. На рабочем ПК это «1 213,97», а на сервере (английская
    локаль) — «1,213.97». ЖИВЬЁМ 21.09.2026: жёстко зашитая замена «,» на
    «.» превратила там суммы вида «1,213.97» в «1.213.97», float() их не
    взял, и в регистр ушли ПУСТЫЕ выручки — итог дня упал в несколько
    раз, причём ВП (числа меньше тысячи, без разделителя) совпала
    точно, из-за чего расхождение и не бросалось в глаза.

    Считаем только однозначные свидетельства:
      * разделитель, после которого НЕ три цифры, тысячным быть не может;
      * число, где есть и группировка, и дробная часть («1 213,97»,
        «1,213.97»), сразу показывает оба символа.
    Количество «1,000» (три знака после запятой в русской локали) и его
    английский двойник «1.000» голосов не дают — они неотличимы от
    тысячной группы, и именно на них ошибались наивные проверки.
    """
    ru = 5 * len(re.findall(r"\d[.\s\u00a0\u202f]\d{3},\d", text))
    en = 5 * len(re.findall(r"\d[,\s\u00a0\u202f]\d{3}\.\d", text))
    ru += len(re.findall(r"\d,\d{1,2}(?!\d)", text))
    en += len(re.findall(r"\d\.\d{1,2}(?!\d)", text))
    if ru != en:
        return "," if ru > en else "."
    # Ничего однозначного: если запятая встречается только как группа из
    # трёх цифр — она тысячная, значит дробная точка.
    if re.search(r"\d,\d{3}(?!\d)", text):
        return "."
    return ","


def clean_number(text: str, decimal_sep: str = ","):
    """
    Число из 1С → float. Убираем пробелы-разделители тысяч (обычный,
    неразрывный, узкий неразрывный) и апостроф, лишний разделитель
    выбрасываем, десятичный приводим к точке.

    Это и есть «пробелы между цифр убираешь, чтоб Google Sheets читал» —
    плюс защита от второго формата чисел (см. detect_decimal_sep).
    """
    text = (text or "").strip()
    if not text:
        return ""
    text = _SPACES_RE.sub("", text)
    text = text.replace("−", "-")  # юникодный минус
    neg = text.startswith("(") and text.endswith(")")
    if neg:
        text = text[1:-1]

    other = "." if decimal_sep == "," else ","
    text = text.replace(other, "")  # разделитель тысяч нам не нужен
    parts = text.split(decimal_sep)
    if len(parts) > 1:
        # Все вхождения, кроме последнего, — тоже группировка тысяч.
        text = "".join(parts[:-1]) + "." + parts[-1]

    try:
        value = float(text)
    except ValueError:
        return ""
    return -value if neg else value


# Как узнать нашу колонку в шапке отчёта: (позиция, признаки названия).
# Проверяется по нормализованному имени (нижний регистр, без пробелов).
_HEADER_RULES = (
    (0, ("клиент",), ("клиент.",)),
    (1, ("менеджер",), ()),
    (2, ("серия",), ()),
    (3, ("подразделение",), ()),
    (4, ("заказклиента",), ("соглашен",)),
    (5, ("имей",), ()),
    (6, ("номенклатура.марка",), ()),
    (7, ("заказклиента",), ("номер",)),
    (8, ("период",), ()),
    (9, ("клиент.группадоступа",), ()),
    (10, ("склад",), ()),
    (11, ("номенклатура",), (".марка", "видномен")),
    (12, ("видноменклатуры",), ()),
    (13, ("количество",), ()),
    (14, ("стоимостьзакупки",), ()),
    (15, ("доп",), ()),
    (16, ("выручка",), ()),
    (17, ("суммаавт",), ()),
    (18, ("валоваяприбыль",), ()),
    (19, ("рентабел",), ()),
)


def _norm(name: str) -> str:
    return (name or "").strip().lower().replace(" ", "").replace(" ", "")


def column_map(header: list) -> list:
    """
    Соответствие «наша колонка → индекс в выгрузке», найденное ПО НАЗВАНИЯМ
    в шапке отчёта.

    ЖИВЬЁМ 22.09.2026: у отчёта на сервере набор колонок оказался другим, и
    жёсткие номера попали не туда — выручка дня выросла в 18 раз, при том что валовая прибыль совпадала до копейки и
    ошибка выглядела как «данные просто такие». Имя колонки надёжнее номера:
    добавили колонку в отчёт — разбор не поедет.

    Возвращает список из 20 индексов; None означает «такой колонки в этом
    отчёте нет» (её значение станет пустым).
    """
    mapping = [None] * COLUMN_COUNT
    used = set()
    # Сначала — точные совпадения названий: если в отчёте есть и «Выручка»,
    # и, скажем, «Выручка (вал.)», брать надо первую, а не ту, что стоит
    # левее (живьём 22.09.2026 из-за этого читалась чужая колонка).
    exact = {_norm(name): idx for idx, name in enumerate(header) if name}
    for slot, canonical in enumerate(COLUMNS):
        idx = exact.get(_norm(canonical))
        if idx is not None and idx not in used:
            mapping[slot] = idx
            used.add(idx)
    for slot, includes, extra in _HEADER_RULES:
        if mapping[slot] is not None:
            continue
        for idx, name in enumerate(header):
            if idx in used:
                continue
            clean = _norm(name)
            if not clean or not all(token in clean for token in includes):
                continue
            if slot == 4 and "номер" not in clean:      # «Заказ клиента ... Номер»
                continue
            if slot == 7 and "соглашен" not in clean:   # «Заказ клиента ... Соглашение»
                continue
            if slot == 11 and any(bad in clean for bad in extra):
                continue
            if slot == 0 and clean != "клиент":
                continue
            mapping[slot] = idx
            used.add(idx)
            break
    return mapping


def column_decimal_seps(samples: list, mapping: list, fallback: str) -> dict:
    """
    Десятичный разделитель ОТДЕЛЬНО для каждой числовой колонки.

    ЖИВЬЁМ 22.09.2026 (сырая выгрузка с сервера): в одной и той же строке
    отчёта форматы разные — «Стоимость закупки» 700 100,50 (запятая
    десятичная), «Выручка» 707,300.75 (запятая — разделитель тысяч),
    «Доп. расходы» -50.25, «Валовая прибыль» 6 800,25. Один разделитель на
    весь файл здесь в принципе не годится: при выборе запятой «216.06»
    превращалось в 21606 — отсюда завышение выручки в десятки раз при
    верной валовой прибыли.

    По каждой колонке голосуем теми же однозначными признаками, что и
    detect_decimal_sep; если в колонке нет ничего однозначного — берём
    общий разделитель по файлу.
    """
    seps = {}
    for idx in sorted(NUMERIC_IDX):
        at = mapping[idx] if idx < len(mapping) else None
        if at is None:
            continue
        values = [cols[at] for cols in samples if at < len(cols) and cols[at].strip()]
        seps[idx] = detect_decimal_sep("\n".join(values)) if values else fallback
    return seps


def _score_revenue_column(samples: list, idx: int, cost_at: int, profit_at: int,
                          seps: dict, fallback: str) -> tuple:
    """
    Насколько колонка idx похожа на «Выручку»: доля строк, где выполняется
    равенство Выручка = Стоимость закупки + Валовая прибыль.

    У каждой колонки свой формат чисел (см. column_decimal_seps), поэтому
    и разбираем каждую своим разделителем — иначе проверка сравнивала бы
    по-разному прочитанные числа и всегда «не сходилась».
    """
    cost_sep = seps.get(14, fallback)
    profit_sep = seps.get(18, fallback)
    value_sep = detect_decimal_sep("\n".join(
        cols[idx] for cols in samples if idx < len(cols) and cols[idx].strip()
    )) or fallback

    ok = total = 0
    for cols in samples:
        if cost_at >= len(cols) or profit_at >= len(cols) or idx >= len(cols):
            continue
        cost = clean_number(cols[cost_at], cost_sep)
        profit = clean_number(cols[profit_at], profit_sep)
        value = clean_number(cols[idx], value_sep)
        if not isinstance(cost, float) or not isinstance(profit, float) or not isinstance(value, float):
            continue
        total += 1
        expected = cost + profit
        if abs(value - expected) <= max(0.02, abs(expected) * 0.001):
            ok += 1
    return (ok / total if total else 0.0), total


def calibrate_revenue(samples: list, mapping: list, seps: dict, fallback: str) -> list:
    """
    Проверяет, что колонка «Выручка» найдена верно, и при необходимости
    находит настоящую — по равенству «Выручка = Стоимость закупки + Валовая
    прибыль», которое в этом отчёте выполняется построчно.

    ЖИВЬЁМ 22.09.2026: на сервере под другим пользователем 1С у отчёта своя
    раскладка, и по имени «Выручка» попадалась не та колонка — выручка дня
    выходила в ~49 раз больше при верной валовой прибыли, то есть выглядела
    правдоподобно и молча портила регистр. Имена колонок у каждого
    пользователя свои, а арифметика — одна на всех, поэтому опираемся на
    неё.
    """
    cost_at, profit_at, revenue_at = mapping[14], mapping[18], mapping[16]
    if cost_at is None or profit_at is None or not samples:
        return mapping

    current, checked = (_score_revenue_column(samples, revenue_at, cost_at, profit_at, seps, fallback)
                        if revenue_at is not None else (0.0, 0))
    if checked and current >= 0.8:
        return mapping

    width = max(len(cols) for cols in samples)
    best_idx, best_score, best_checked = None, 0.0, 0
    for idx in range(width):
        if idx in (cost_at, profit_at):
            continue
        score, total = _score_revenue_column(samples, idx, cost_at, profit_at, seps, fallback)
        if total >= max(3, len(samples) // 4) and score > best_score:
            best_idx, best_score, best_checked = idx, score, total

    if best_idx is None or best_score < 0.8:
        print("  ! Не удалось подтвердить колонку «Выручка» равенством "
              "«закупка + валовая прибыль» — оставляю найденную по шапке.")
        return mapping
    if best_idx != revenue_at:
        print(f"  ! Колонка «Выручка» по шапке ({revenue_at}) не сходится с арифметикой "
              f"(совпало {current:.0%}); беру колонку {best_idx} — там сходится "
              f"{best_score:.0%} из {best_checked} строк.")
        mapping = list(mapping)
        mapping[16] = best_idx
    return mapping


def find_header(text: str) -> list:
    """Строка-шапка отчёта («Клиент», «Менеджер», …) или None."""
    for line in text.replace("\r\n", "\n").split("\n"):
        cols = line.split("\t")
        if len(cols) >= 10 and _norm(cols[0]) == "клиент":
            return cols
    return None


def parse_period(text: str, order: str = "dmy"):
    """
    «19.09.2026» / «9/19/2026» / «19.09.2026 0:00:00» → date, иначе None.

    order — порядок чисел в дате: "dmy" (день первым) или "mdy" (месяц
    первым). Однозначные случаи (одно из чисел больше 12) разбираются
    правильно при любом order, остальные — по нему.
    """
    m = _DATE_RE.match(text or "")
    if not m:
        return None
    a, b, y = (int(g) for g in m.groups())
    if a > 12:
        day, month = a, b
    elif b > 12:
        day, month = b, a
    elif order == "mdy":
        day, month = b, a
    else:
        day, month = a, b
    try:
        return date(y, month, day)
    except ValueError:
        return None


def detect_date_order(text: str, expect_from: date = None, expect_to: date = None) -> str:
    """
    В каком порядке 1С печатает даты в ЭТОЙ выгрузке: "dmy" или "mdy".

    ЖИВЬЁМ 24.09.2026: на сервере 1С печатает «9/1/2026» (месяц первым).
    Разбор читал это как 9 января, строки уходили «вне периода», и день
    молча терялся — именно так пропали 01–12.09, а числа больше 12
    (13–31) при этом разбирались верно, поэтому проблема выглядела как
    «обрезанная выгрузка».

    Порядок берём из строки «Параметры: Период: A - B»: мы знаем, какой
    период запрашивали, и просто смотрим, какое прочтение с ним сходится.
    Если шапки нет — ищем в данных дату, где одно из чисел больше 12:
    она однозначна. Если и таких нет (например, период 01–07 сентября),
    остаётся значение по умолчанию.
    """
    m = _PARAMS_PERIOD_RE.search(text)
    if m and expect_from and expect_to:
        d1, m1, y1, d2, m2, y2 = (int(g) for g in m.groups())
        for order, (day1, mon1, day2, mon2) in (("dmy", (d1, m1, d2, m2)),
                                                ("mdy", (m1, d1, m2, d2))):
            try:
                if (date(y1, mon1, day1), date(y2, mon2, day2)) == (expect_from, expect_to):
                    return order
            except ValueError:
                continue

    for line in text.replace("\r\n", "\n").split("\n"):
        cols = line.split("\t")
        if len(cols) <= PERIOD_IDX:
            continue
        found = _DATE_RE.match(cols[PERIOD_IDX].strip())
        if not found:
            continue
        a, b, _ = (int(g) for g in found.groups())
        if a > 12:
            return "dmy"
        if b > 12:
            return "mdy"
    return "dmy"


def parse_report(text: str, expect_from: date = None, expect_to: date = None,
                 failures: list = None) -> list:
    """
    TSV отчёта → список строк по 20 значений (числа уже как float).

    Строкой данных считаем ту, где в колонке «Период» стоит дата. Это разом
    отсекает заголовок, «Параметры:», итоги и пустые строки — тот же приём,
    что в соседнем проекте (src/parse_report.py там).
    """
    rows = []
    out_of_range = []
    decimal_sep = detect_decimal_sep(text)
    order = detect_date_order(text, expect_from, expect_to)
    if order != "dmy":
        print(f"  Даты в выгрузке в формате «{'месяц/день' if order == 'mdy' else 'день.месяц'}»")
    print(f"  Формат чисел в выгрузке: десятичный разделитель «{decimal_sep}»")

    # Колонки ищем по названиям в шапке отчёта; номера — только если шапки
    # нет вовсе (см. column_map).
    header = find_header(text)
    mapping = column_map(header) if header else None
    if mapping is None:
        print("  ! Шапку отчёта не нашёл — читаю колонки по их обычным номерам.")
        mapping = list(range(COLUMN_COUNT))
    else:
        missing = [COLUMNS[i] for i, idx in enumerate(mapping) if idx is None]
        if missing:
            print(f"  ! В отчёте нет колонок: {', '.join(missing)} — они останутся пустыми.")
        if mapping[:COLUMN_COUNT] != list(range(COLUMN_COUNT)):
            print(f"  Колонки отчёта разложены по шапке: {mapping}")
    # Колонку «Выручка» дополнительно проверяем арифметикой строк —
    # названия у разных пользователей 1С свои (см. calibrate_revenue).
    data_cols = [
        line.split("\t")
        for line in text.replace("\r\n", "\n").split("\n")
        if line.count("\t") >= 10
    ]
    # Формат чисел определяем ПО КАЖДОЙ КОЛОНКЕ (см. column_decimal_seps):
    # в одной строке отчёта они могут отличаться.
    seps = column_decimal_seps(data_cols, mapping, decimal_sep)
    if header and mapping[14] is not None and mapping[18] is not None:
        mapping = calibrate_revenue(data_cols[:400], mapping, seps, decimal_sep)
        seps = column_decimal_seps(data_cols, mapping, decimal_sep)
    mixed = {COLUMNS[i]: s for i, s in seps.items() if s != decimal_sep}
    if mixed:
        print("  ! В колонках свой формат чисел: "
              + ", ".join(f"{name} «{s}»" for name, s in mixed.items()))

    if mapping[PERIOD_IDX] is None:
        # Шапка нашлась, но опознать в ней «Период» не удалось (обрезанная
        # шапка, другое название). Падать нельзя — возвращаемся к обычным
        # номерам колонок, как было до разбора по шапке.
        print("  ! В шапке не нашёл колонку «Период» — читаю колонки по номерам.")
        mapping = list(range(COLUMN_COUNT))
    period_at = mapping[PERIOD_IDX]

    for line in text.replace("\r\n", "\n").split("\n"):
        if not line.strip():
            continue
        cols = line.split("\t")
        if len(cols) <= period_at:
            continue
        day = parse_period(cols[period_at], order)
        if day is None:
            continue
        if expect_from and expect_to and not (expect_from <= day <= expect_to):
            out_of_range.append(day)
            continue

        cols = [c.strip() for c in cols]
        row = []
        for idx in range(COLUMN_COUNT):
            at = mapping[idx]
            value = cols[at] if at is not None and at < len(cols) else ""
            if idx == PERIOD_IDX:
                row.append(day)
            elif idx in NUMERIC_IDX:
                number = clean_number(value, seps.get(idx, decimal_sep))
                # Непустой текст, который НЕ стал числом — ровно та поломка,
                # из-за которой 21.09 в регистр ушли пустые выручки. Ловим
                # её здесь: это не зависит ни от разметки строки «Итого»,
                # ни от набора колонок в конкретном варианте отчёта.
                if number == "" and value.strip():
                    if failures is not None:
                        failures.append((COLUMNS[idx], value))
                row.append(number)
            else:
                row.append(value)
        rows.append(row)

    if out_of_range:
        uniq = sorted(set(out_of_range))
        print(f"  ! Пропущено {len(out_of_range)} строк с датами вне периода: {uniq[:5]}")
    return rows


def report_totals(text: str, decimal_sep: str = None) -> dict:
    """
    Строка «Итого» из самого отчёта — контрольная сумма выгрузки.

    ЖИВЬЁМ 21.09.2026: на сервере часть сумм выручки молча не разобралась и
    ушла в регистр пустыми — итог дня упал в несколько раз, и заметил это
    человек, а не программа. Отчёт печатает собственный «Итого» — сверка с
    ним ловит любой сбой разбора чисел сразу, независимо от того, какой
    формат чисел придёт от очередной машины.
    """
    if decimal_sep is None:
        decimal_sep = detect_decimal_sep(text)

    lines = text.replace("\r\n", "\n").split("\n")
    header = None
    total = None
    for line in lines:
        cols = line.split("\t")
        if header is None and len(cols) >= COLUMN_COUNT and cols[0].strip() == "Клиент":
            header = cols
        elif cols and cols[0].strip().lower() == "итого":
            total = cols
    if header is None or total is None:
        return None

    def _index(*names):
        for idx, name in enumerate(header):
            clean = name.strip().lower().replace(" ", "")
            if any(clean.startswith(n) for n in names):
                return idx
        return None

    # Строка «Итого» бывает выровнена не так, как строки данных: у неё
    # своё число ведущих пустых ячеек. ЖИВЬЁМ 22.09.2026 на сервере из-за
    # этого по индексу «Выручки» читались «Доп. расходы» (копейки вместо
    # выручки), сверка объявила ложное расхождение и остановила
    # прогон — данных за 21.09 в таблице не появилось. Поэтому индексы
    # берём из шапки отчёта и поправляем на сдвиг самой строки итога.
    shift = len(total) - len(header)
    out = {}
    for key, names in (
        ("quantity", ("количество",)),
        ("revenue", ("выручка",)),
        ("profit", ("валоваяприбыль",)),
    ):
        idx = _index(*names)
        if idx is None:
            continue
        pos = idx + shift
        if 0 <= pos < len(total):
            column_sep = detect_decimal_sep(total[pos]) if total[pos].strip() else decimal_sep
            out[key] = clean_number(total[pos], column_sep)

    # Все числа строки итога — запасной якорь для сверки: если выравнивание
    # окажется ещё каким-то третьим, сумма строк всё равно должна найтись
    # среди них (см. verify_totals).
    if out:
        out["_numbers"] = [
            value for value in (clean_number(cell, decimal_sep) for cell in total)
            if isinstance(value, float)
        ]
    return out or None


def verify_parsed(rows: list, failures: list) -> None:
    """
    Главная защита от порчи данных: ни одно непустое число отчёта не должно
    остаться неразобранным.

    ЖИВЬЁМ 21.09.2026: суммы выручки молча не распознались и ушли в регистр
    пустыми — день показал в несколько раз меньше. Проверка по строке «Итого»
    от этого не спасла: её разметка у разных вариантов отчёта разная (на
    сервере там вообще оказались подытоги групп, и сверка стала ложно
    падать, теряя день). А вот «текст был, числа не вышло» — признак
    однозначный и от разметки не зависит.
    """
    if not rows:
        return
    if failures:
        examples = "; ".join(f"{name}={raw!r}" for name, raw in failures[:5])
        raise RuntimeError(
            f"Не разобрано чисел: {len(failures)}. В таблицу ничего не пишу, "
            f"иначе туда уйдут пустые суммы. Примеры: {examples}"
        )
    print(f"  Все числа разобраны ({len(rows)} строк).")


def verify_totals(rows: list, totals: dict, tolerance: float = 0.05) -> None:
    """
    Сверка суммы разобранных строк с «Итого» самого отчёта — стоп-кран
    против ОБРЕЗАННОЙ выгрузки.

    1С отдаёт отчёт через буфер обмена, и на длинном периоде он приходит
    неполным: ЖИВЬЁМ 23.09.2026 при заливке 01–22.09 строки за 01–12.09 не
    пришли вовсе, а прогон уже удалил прежние. Итог отчёта считает сама 1С
    по всем строкам, поэтому расхождение с ним — верный признак, что до нас
    доехало не всё.

    Если «Итого» размечен иначе (подытоги групп) и по индексу колонки лежит
    чужое число, сумма ищется среди ВСЕХ чисел строки итога — ложную
    тревогу это снимает, а обрезанную выгрузку всё равно поймает.
    """
    if not totals:
        print("  ! В отчёте нет строки «Итого» — сверить суммы не с чем.")
        return
    checks = (
        ("Выручка", sum(r[16] for r in rows if isinstance(r[16], float)), totals.get("revenue")),
        ("Валовая прибыль", sum(r[18] for r in rows if isinstance(r[18], float)), totals.get("profit")),
    )
    numbers = totals.get("_numbers") or []
    problems = []
    for name, got, expected in checks:
        if not isinstance(expected, float):
            continue
        if abs(got - expected) <= tolerance:
            continue
        # Значение не совпало с тем, что лежит по индексу колонки. Прежде
        # чем останавливать прогон, ищем нашу сумму среди ВСЕХ чисел строки
        # итога: если она там есть, то сошлось всё, а разъехалось только
        # выравнивание самой строки — это не повод терять день.
        if any(abs(got - value) <= tolerance for value in numbers):
            print(f"  ! «Итого» отчёта выровнен иначе, чем строки данных; "
                  f"{name} нашлась в нём по другому столбцу — считаю сверку пройденной.")
            continue
        problems.append(f"{name}: разобрано {got:,.2f}, в «Итого» отчёта {expected:,.2f}")
    if problems:
        reason = (
            "Ни одна строка отчёта не попала в запрошенный период — похоже, "
            "даты в выгрузке прочитаны не в том порядке (день/месяц)."
            if not rows else
            "Похоже на обрезанную выгрузку: 1С отдала в буфер обмена не все строки."
        )
        raise RuntimeError(
            f"Суммы разобранных строк не сходятся с «Итого» самого отчёта. {reason} "
            "В таблицу ничего не пишу, иначе день потеряет часть продаж.\n  "
            + "\n  ".join(problems)
        )
    print("  Суммы сходятся с «Итого» отчёта.")


# Колонка «Клиент.Группа доступа» в разобранной строке.
CLIENT_GROUP_IDX = 9


def drop_excluded_groups(rows: list, excluded) -> tuple:
    """
    Убирает строки оптовых клиентов — то же, что даёт галка «Клиент.Группа
    доступа (не в списке): B2B; B2B new; optom» в самом отчёте.

    ЖИВЬЁМ 22.09.2026: человек формирует отчёт С этой галкой, а автомат её
    не ставил — и в регистр попадали оптовые строки. Сверка с ручной
    таблицей коллег показала это в цифрах: без фильтра 19.09 давало
    заметно больше, с фильтром — ровно как у них; так же сошлись
    16.09, 18.09 и 20.09.

    Фильтруем в коде, а не галкой в 1С: одна галка в чужой форме — лишний
    шаг UI-автоматизации, который может не прожаться и молча изменить
    выборку, а здесь условие явное и проверяемое.

    Возвращает (оставшиеся строки, сколько убрано, по группам).
    """
    wanted = {str(g).strip().lower() for g in excluded if str(g).strip()}
    if not wanted:
        return rows, 0, {}
    kept, by_group = [], {}
    for row in rows:
        group = str(row[CLIENT_GROUP_IDX]).strip().lower()
        if group in wanted:
            by_group[group] = by_group.get(group, 0) + 1
            continue
        kept.append(row)
    return kept, len(rows) - len(kept), by_group


def rows_by_date(rows: list) -> dict:
    """Группировка разобранных строк по дате — для логов и проверок."""
    out = {}
    for row in rows:
        out.setdefault(row[PERIOD_IDX], []).append(row)
    return out
