"""
Запись строк отчёта в лист «Регистр» Google-таблицы — ровно то, что
делается руками:

  1. строки за нужные даты, если они уже есть, удаляются целиком
     (чтобы повторный прогон не задваивал день);
  2. новые строки дописываются в конец: A..I и L..V — данные из 1С,
     J/K не трогаем;
  3. формулы J/K «протягиваются» вниз копированием из последней строки
     (Google сам пересчитывает ссылки на свою строку) — вместе с форматом
     всей строки A..V, чтобы дата в I и числа в P..V выглядели как у
     соседей;
  4. строки, у которых формула в J вернула пусто (канал продаж не
     определился), удаляются целиком.

Дата в колонку I пишется ЧИСЛОМ-СЕРИЙНИКОМ (дней от 30.12.1899), а не
текстом: текстовую дату Google разбирает по локали таблицы, и «19.09.2026»
в англоязычной локали молча превратилось бы в текст, после чего SUMIFS на
листе «Sales by day» (сравнивает I с датой) перестал бы видеть эти строки.
Серийник + формат даты в скопированном формате строки дают один и тот же
результат в любой локали.
"""
import re
import time
from datetime import date

import gspread
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

# Ноль дат Google Sheets/Excel.
SERIAL_EPOCH = date(1899, 12, 30)

PERIOD_IDX = 8          # индекс «Период» в строке из gross_profit
FIRST_BLOCK = (0, 9)    # колонки отчёта → A..I
SECOND_BLOCK = (9, 20)  # колонки отчёта → L..V
FORMULA_COLS = ("J", "K")

# Предохранитель: сколько строк за раз разрешено удалять. Дневной объём —
# десятки строк, недельный перезалив — сотни; тысячи означают, что что-то
# пошло не так (например, в колонке дат оказался мусор), и лучше упасть.
MAX_DELETE_ROWS = 5000


def to_serial(day: date) -> int:
    return (day - SERIAL_EPOCH).days


def from_serial(value) -> date:
    try:
        return SERIAL_EPOCH.fromordinal(SERIAL_EPOCH.toordinal() + int(value))
    except (TypeError, ValueError):
        return None


def get_client(credentials_path: str, retries: int = 5) -> gspread.Client:
    last = None
    for attempt in range(1, retries + 1):
        try:
            creds = Credentials.from_service_account_file(credentials_path, scopes=SCOPES)
            return gspread.authorize(creds)
        except Exception as e:  # noqa: BLE001 — сеть/DNS, имеет смысл повторить
            last = e
            delay = min(5 * 2 ** (attempt - 1), 60)
            print(f"  ! Google: ошибка авторизации ({e}); повтор через {delay} c")
            time.sleep(delay)
    raise last


def open_register(client: gspread.Client, spreadsheet_id: str, sheet_name: str):
    try:
        ss = client.open_by_key(spreadsheet_id)
    except gspread.exceptions.APIError as e:
        if "must not be an Office file" in str(e):
            raise RuntimeError(
                f"Таблица {spreadsheet_id} — это .xlsx-файл на Google Диске, "
                "Sheets API такие файлы не пишет. Нужен ID НАТИВНОЙ Google-таблицы "
                "(Файл → Сохранить как Google Таблицы), пропиши его в .env "
                "(REGISTER_SPREADSHEET_ID)."
            ) from e
        raise
    return ss, ss.worksheet(sheet_name)


def read_period_column(ws, header_row: int):
    """
    Колонка I целиком → список (номер строки, date|None) начиная со строки
    после заголовка. Читаем UNFORMATTED_VALUE, чтобы получить серийники, а
    не подписи в локальном формате.
    """
    first = header_row + 1
    values = ws.get(
        f"I{first}:I{ws.row_count}",
        value_render_option="UNFORMATTED_VALUE",
    )
    out = []
    for offset, row in enumerate(values):
        value = row[0] if row else ""
        out.append((first + offset, from_serial(value) if value != "" else None))
    return out


def read_day_totals(ws, header_row: int) -> dict:
    """
    Выручка по дням прямо из листа: {дата: сумма по колонке S}.

    Нужна, чтобы прогон мог сказать, ЧТО изменилось при перезаливе прошлых
    дней. 1С правит документы задним числом (за 19.09 через два дня
    отдавала уже другую сумму), и без такого отчёта итог месяца «сам собой»
    прыгает, а человек видит только конечную цифру и не понимает, почему.
    """
    first = header_row + 1
    values = ws.get(f"I{first}:S{ws.row_count}", value_render_option="UNFORMATTED_VALUE")
    totals = {}
    for raw in values:
        raw = (list(raw) + [""] * 11)[:11]
        day = from_serial(raw[0]) if raw[0] != "" else None
        if day is None:
            continue
        if isinstance(raw[10], (int, float)):
            totals[day] = totals.get(day, 0.0) + raw[10]
        else:
            totals.setdefault(day, 0.0)
    return totals


def last_data_row(period_column, header_row: int) -> int:
    """Номер последней строки с датой в колонке I."""
    last = header_row
    for row_index, day in period_column:
        if day is not None:
            last = row_index
    return last


def _contiguous(rows):
    """[5,6,7,10,11] → [(5,7),(10,11)] — для deleteDimension пачками."""
    runs = []
    for row in sorted(rows):
        if runs and row == runs[-1][1] + 1:
            runs[-1][1] = row
        else:
            runs.append([row, row])
    return [tuple(r) for r in runs]


def delete_rows(ss, ws, rows) -> int:
    """Удаляет строки целиком (снизу вверх, пачками подряд идущих)."""
    rows = sorted(set(rows))
    if not rows:
        return 0
    if len(rows) > MAX_DELETE_ROWS:
        raise RuntimeError(
            f"Под удаление попало {len(rows)} строк (лимит {MAX_DELETE_ROWS}). "
            "Похоже на ошибку в данных — ничего не удаляю."
        )
    requests = [
        {
            "deleteDimension": {
                "range": {
                    "sheetId": ws.id,
                    "dimension": "ROWS",
                    "startIndex": start - 1,
                    "endIndex": end,
                }
            }
        }
        for start, end in reversed(_contiguous(rows))
    ]
    ss.batch_update({"requests": requests})
    return len(rows)


def _sheet_row_count(ss, ws) -> int:
    """
    Сколько строк на листе СЕЙЧАС — спрашиваем у Google, а не у объекта.

    ws.row_count кэшируется при открытии листа и после наших же вставок и
    удалений врёт. ЖИВЬЁМ 23.09.2026: запись упала с «Range exceeds grid
    limits. Max rows: 17595», хотя место формально проверялось.
    """
    try:
        meta = ss.fetch_sheet_metadata({"fields": "sheets.properties"})
        for sheet in meta.get("sheets", []):
            props = sheet.get("properties", {})
            if props.get("sheetId") == ws.id:
                return int(props.get("gridProperties", {}).get("rowCount", ws.row_count))
    except Exception as e:  # noqa: BLE001 — не смогли спросить, полагаемся на кэш
        print(f"  ! Не удалось уточнить размер листа ({e}) — беру кэш.")
    return ws.row_count


def _ensure_rows(ss, ws, needed_last_row: int) -> None:
    if needed_last_row <= _sheet_row_count(ss, ws):
        return
    current = _sheet_row_count(ss, ws)
    ss.batch_update({
        "requests": [{
            "appendDimension": {
                "sheetId": ws.id,
                "dimension": "ROWS",
                "length": max(needed_last_row - current, 0) + 200,
            }
        }]
    })


def _row_to_blocks(row):
    """Строка из gross_profit → (значения A..I, значения L..V)."""
    first = []
    for idx in range(*FIRST_BLOCK):
        value = row[idx]
        first.append(to_serial(value) if idx == PERIOD_IDX else value)
    second = [row[idx] for idx in range(*SECOND_BLOCK)]
    return first, second



NEW_MARKETPLACE_CLIENT = "MARKETPLACE D"
NEW_MARKETPLACE_CHANNEL = "Маркетплейс D"
RETAIL_CHANNEL = "Розница"
_LAST_MARKETPLACE_TAIL_RE = re.compile(r'"Маркетплейс C"\s*([;,])\s*""')
# Розничные подразделения перечислены в формуле списком
# OR(D<строка>="Магазин 1";…), сразу за которым идёт "Розница". Таких списков
# два: один в ветке BNPL (рассрочка), другой — в обычной продаже.
_RETAIL_OR_RE = re.compile(r'OR\(([^()]*)\)(\s*[;,]\s*)"' + RETAIL_CHANNEL + '"')


def add_marketplace_branch(formula: str, row: int) -> str:
    """
    Дописывает в формулу «Группа» ветку нового маркетплейса (D).

    Его строки приходят как Клиент «MARKETPLACE D», Соглашение
    «Marketplace». Формула знала маркетплейсы A, B и C, а D — нет:
    J оставался пустым, и прогон удалял такие строки как «без канала». В
    основном это возвраты (минусовые суммы), из-за чего наши цифры и
    расходились с ручной таблицей (у них за 21.09 такая строка есть).

    Разделитель аргументов (; или ,) берём из самой формулы — он зависит
    от локали книги.
    """
    if NEW_MARKETPLACE_CLIENT.lower() in formula.lower():
        return formula
    match = _LAST_MARKETPLACE_TAIL_RE.search(formula)
    if not match:
        print("  ! В формуле «Группа» не нашёл ветку «Маркетплейс C» — маркетплейс D не добавляю, "
              "проверь столбец J вручную.")
        return formula
    sep = match.group(1)
    return (formula[:match.start()]
            + f'"Маркетплейс C"{sep}IF(A{row}="{NEW_MARKETPLACE_CLIENT}"{sep}"{NEW_MARKETPLACE_CHANNEL}"{sep}"")'
            + formula[match.end():])


def add_retail_divisions(formula: str, row: int, divisions) -> str:
    """
    Дописывает подразделения в перечень «Розницы» формулы «Группа».

    Формула определяет канал по Подразделению (D) и Соглашению (H).
    Розничные магазины перечислены в ней поимённо, и магазина, которого в
    списке нет, формула не знает: J остаётся пустым, а прогон удаляет
    строку как «без канала».

    ЖИВЬЁМ 05.10.2026: в 1С появилось подразделение «Магазин 7» (продажи
    через BNPL-партнёра, розница и рассрочка) — в регистр эти строки не попадали
    вовсе, и день расходился с ручной таблицей.

    Дописываем в ОБА списка розницы (BNPL и обычный): в новом магазине
    идут и те и другие продажи. Повторный вызов ничего не меняет.
    """
    wanted = [str(d).strip() for d in divisions if str(d).strip()]
    if not wanted:
        return formula

    def patch(match):
        body, gap = match.group(1), match.group(2)
        sep = ";" if ";" in body else ","
        for division in wanted:
            if f'"{division}"' not in body:
                body = f'{body}{sep}D{row}="{division}"'
        return f'OR({body}){gap}"{RETAIL_CHANNEL}"'

    return _RETAIL_OR_RE.sub(patch, formula)


def ensure_group_formula(ws, template_row: int, retail_divisions=()) -> bool:
    """
    Следит, чтобы формула «Группа» (J) знала маркетплейс D и все розничные
    подразделения из RETAIL_DIVISIONS.

    Проверяем строку-образец, с которой формула протягивается на новые
    строки, и дописываем только то, чего в ней нет. Если всё на месте —
    ничего не трогаем и в книгу не пишем.
    """
    current = ws.get(f"J{template_row}", value_render_option="FORMULA")
    formula = current[0][0] if current and current[0] else ""
    if not str(formula).startswith("="):
        return False

    patched = add_marketplace_branch(str(formula), template_row)
    patched = add_retail_divisions(patched, template_row, retail_divisions)

    missing = [str(d).strip() for d in retail_divisions
               if str(d).strip() and f'"{str(d).strip()}"' not in patched]
    if missing:
        print("  ! Не смог добавить в формулу «Группа» подразделения: "
              + ", ".join(missing) + " — проверь столбец J вручную.")

    if patched == formula:
        return False

    ws.update(values=[[patched]], range_name=f"J{template_row}", value_input_option="USER_ENTERED")
    print(f"  Формула «Группа» дополнена (строка-образец {template_row}).")
    return True


def append_rows(ss, ws, rows, last_row: int, header_row: int, retail_divisions=()) -> tuple:
    """
    Дописывает строки после last_row: значения A..I и L..V, затем протягивает
    формат A..V и формулы J/K из строки-образца.

    Образец — последняя строка с данными; если регистр пуст (такого на
    практике не бывает, но пусть будет), формулы не протягиваются и об этом
    сообщается: сами формулы здесь не выдумываем, они живут в книге.
    """
    if not rows:
        return last_row, last_row, 0

    start = last_row + 1
    end = last_row + len(rows)
    _ensure_rows(ss, ws, end)

    first_block, second_block = zip(*(_row_to_blocks(row) for row in rows))

    ws.batch_update(
        [
            {"range": f"A{start}:I{end}", "values": [list(v) for v in first_block]},
            {"range": f"L{start}:V{end}", "values": [list(v) for v in second_block]},
        ],
        value_input_option="RAW",
    )

    _apply_formats(ss, ws, start, end)
    if last_row > header_row:
        # Перед протяжкой убеждаемся, что формула-образец знает маркетплейс D и
        # все розничные подразделения (см. ensure_group_formula) — иначе
        # такие строки снова уйдут в «без канала» и день разойдётся с
        # ручной таблицей.
        ensure_group_formula(ws, last_row, retail_divisions)
        _fill_formulas(ws, source_row=last_row, start=start, end=end)
    else:
        print("  ! В регистре нет строки-образца — формулы J/K не протянуты.")

    return start, end, len(rows)


def _shift_formula(formula: str, source_row: int, target_row: int) -> str:
    """
    Переносит формулу на другую строку: A17536 → A17540 и т.д.

    Меняем только ссылки на СВОЮ строку-образец (без $ перед номером) —
    абсолютные ссылки и любые другие числа в формуле остаются как есть.
    """
    pattern = re.compile(rf"(\$?[A-Za-z]{{1,3}})(?<!\$){source_row}\b")
    return pattern.sub(lambda m: f"{m.group(1)}{target_row}", formula)


def _fill_formulas(ws, source_row: int, start: int, end: int) -> None:
    """
    Протягивает формулы J/K на новые строки ПРЯМОЙ ЗАПИСЬЮ, а не
    копированием диапазона.

    ЖИВЬЁМ 23.09.2026: на листе оказался включён фильтр, и Google отказал
    копированию — «This operation is not supported on a range with a
    filtered out row». Прогон к этому моменту уже удалил старые строки, и
    день остался в таблице без формул: J пустой, канал продаж не
    определился, в дашборде дыра. Прямая запись формул фильтра не боится.
    """
    template = ws.get(f"J{source_row}:K{source_row}", value_render_option="FORMULA")
    row = (template[0] if template else []) + ["", ""]
    j_formula, k_formula = row[0], row[1]
    if not str(j_formula).startswith("="):
        print("  ! В строке-образце нет формулы «Группа» — J/K не заполняю.")
        return
    values = [
        [_shift_formula(j_formula, source_row, target), _shift_formula(k_formula, source_row, target)]
        for target in range(start, end + 1)
    ]
    ws.update(values=values, range_name=f"J{start}:K{end}", value_input_option="USER_ENTERED")


# Форматы колонок регистра: дата в I, числа в P..V (как в исходной книге).
_NUMBER_FORMATS = {
    "I": {"type": "DATE", "pattern": "dd.MM.yyyy"},
    "P": {"type": "NUMBER", "pattern": "0.000"},
    "Q": {"type": "NUMBER", "pattern": "0.00"},
    "R": {"type": "NUMBER", "pattern": "0.00"},
    "S": {"type": "NUMBER", "pattern": "0.00"},
    "T": {"type": "NUMBER", "pattern": "0.00"},
    "U": {"type": "NUMBER", "pattern": "0.00"},
    "V": {"type": "NUMBER", "pattern": "0.00"},
}


def _apply_formats(ss, ws, start: int, end: int) -> None:
    """
    Ставит форматы новым строкам явными запросами.

    Раньше формат копировался с соседней строки (PASTE_FORMAT), но
    копирование ломается о включённый фильтр — см. _fill_formulas. Формат
    здесь предсказуемый и одинаковый для всех строк, так что задаём его
    прямо: без него дата в I показывалась бы числом 46284.
    """
    requests = []
    for letter, number_format in _NUMBER_FORMATS.items():
        column = ord(letter) - ord("A")
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": ws.id,
                    "startRowIndex": start - 1,
                    "endRowIndex": end,
                    "startColumnIndex": column,
                    "endColumnIndex": column + 1,
                },
                "cell": {"userEnteredFormat": {"numberFormat": number_format}},
                "fields": "userEnteredFormat.numberFormat",
            }
        })
    ss.batch_update({"requests": requests})


def drop_unclassified(ss, ws, start: int, end: int, unknown=None) -> int:
    """
    Удаляет из только что добавленного блока строки, где формула «Группа»
    (J) вернула пусто — канал продаж не определился. Ровно то же, что руками
    делается фильтром по пустым в J.

    unknown — необязательный словарь, в который складывается состав
    удалённого: {(подразделение, соглашение): сколько}. Нужен, чтобы
    написать об этом в Telegram: новый магазин иначе молча теряется.
    """
    if end < start:
        return 0
    values = ws.get(f"J{start}:J{end}", value_render_option="UNFORMATTED_VALUE")
    empty = []
    for offset in range(end - start + 1):
        row = values[offset] if offset < len(values) else []
        value = (row[0] if row else "")
        if str(value).strip() == "":
            empty.append(start + offset)
    if not empty:
        return 0

    # Перед удалением показываем, ЧТО именно не распозналось: формула
    # «Группа» смотрит на Подразделение (D) и Соглашение (H), и пустой
    # результат означает, что такой пары в ней просто нет. Без этого списка
    # видно только «удалено N строк», и понять, что день потерял половину
    # выручки из-за нового склада или нового соглашения, невозможно.
    try:
        pairs = {}
        rows_dh = ws.get(f"D{start}:H{end}", value_render_option="UNFORMATTED_VALUE")
        for row_index in empty:
            row = rows_dh[row_index - start] if row_index - start < len(rows_dh) else []
            row = (list(row) + [""] * 5)[:5]
            key = (str(row[0]).strip(), str(row[4]).strip())  # D — подразделение, H — соглашение
            pairs[key] = pairs.get(key, 0) + 1
        if unknown is not None:
            unknown.update(pairs)
        print("  Строки без канала продаж (подразделение / соглашение → сколько):")
        for (dep, agreement), count in sorted(pairs.items(), key=lambda kv: -kv[1]):
            print(f"    {dep or '—'} / {agreement or '—'} → {count}")
    except Exception as e:  # noqa: BLE001 — диагностика не должна мешать удалению
        print(f"  ! Не смог показать состав строк без канала: {e}")

    delete_rows(ss, ws, empty)
    return len(empty)


def sweep_excluded_groups(ss, ws, header_row: int, excluded) -> int:
    """
    Убирает из регистра строки оптовых клиентов, где бы они ни лежали.

    Фильтр по группе доступа (см. gross_profit.drop_excluded_groups) не
    пускает такие строки в новые записи, но в книге могли остаться строки
    от прогонов, сделанных до появления фильтра. ЖИВЬЁМ 23.09.2026: за
    14.09 осталось 6 строк группы optom — ровно на их сумму
    «Optom 2» в нашем дашборде расходился с ручной таблицей,
    хотя дневные цифры сходились идеально.

    Проверка дешёвая (одно чтение колонки L) и делает регистр
    самовосстанавливающимся: пересчитывать старые дни через 1С не нужно.
    """
    wanted = {str(g).strip().lower() for g in excluded if str(g).strip()}
    if not wanted:
        return 0
    first = header_row + 1
    values = ws.get(f"L{first}:L{ws.row_count}", value_render_option="UNFORMATTED_VALUE")
    doomed = [
        first + offset
        for offset, row in enumerate(values)
        if row and str(row[0]).strip().lower() in wanted
    ]
    if not doomed:
        return 0
    print(f"  Нашёл {len(doomed)} строк исключаемых групп, оставшихся от прежних прогонов — удаляю.")
    delete_rows(ss, ws, doomed)
    return len(doomed)


def write_days(cfg, rows, days) -> dict:
    """
    Полный цикл записи: удалить старые строки за days, дописать новые,
    протянуть формулы, вычистить нераспознанные.

    rows — разобранные строки отчёта (gross_profit.parse_report).
    days — множество дат, которые этот прогон перезаписывает.
    """
    client = get_client(cfg.credentials_path)
    ss, ws = open_register(client, cfg.spreadsheet_id, cfg.register_sheet)

    header_row = cfg.register_header_row
    period_column = read_period_column(ws, header_row)
    before_last = last_data_row(period_column, header_row)

    days = set(days)
    stale = [row_index for row_index, day in period_column if day in days]
    # Снимок «до» — чтобы показать, что именно изменилось в прошлых днях.
    before_totals = {day: value for day, value in read_day_totals(ws, header_row).items()
                     if day in days}

    # ВАЖЕН ПОРЯДОК: сначала дописываем новые строки и только потом удаляем
    # прежние. ЖИВЬЁМ 23.09.2026: удаление прошло, а запись упала (Google
    # отказал копированию из-за фильтра на листе) — и в таблице не осталось
    # ни старых строк, ни новых. Старые строки лежат ВЫШЕ нового блока,
    # поэтому их удаление в конце не сдвигает уже записанное.
    last_row = last_data_row(period_column, header_row)
    start, end, appended = append_rows(ss, ws, rows, last_row, header_row,
                                       getattr(cfg, "retail_divisions", ()))
    print(f"  Дописано строк: {appended} (строки {start}–{end})")

    unknown = {}
    dropped = drop_unclassified(ss, ws, start, end, unknown)
    if dropped:
        print(f"  Удалено строк без канала продаж (пусто в J): {dropped}")

    # Теперь, когда новые строки на месте и посчитаны, убираем прежние.
    deleted = delete_rows(ss, ws, stale)
    if deleted:
        print(f"  Удалены прежние строки за эти даты: {deleted}")

    # Подчищаем оптовые строки, оставшиеся от прогонов до появления фильтра
    # (см. sweep_excluded_groups).
    swept = sweep_excluded_groups(ss, ws, header_row, getattr(cfg, "exclude_client_groups", ()))

    after_totals = {day: value for day, value in read_day_totals(ws, header_row).items()
                    if day in days}
    changed = {
        day: (before_totals.get(day, 0.0), after_totals.get(day, 0.0))
        for day in sorted(days)
        if abs(after_totals.get(day, 0.0) - before_totals.get(day, 0.0)) > 0.05
        and before_totals.get(day) is not None
    }
    if changed:
        print("  Прошлые дни изменились (1С правит документы задним числом):")
        for day, (was, now) in changed.items():
            print(f"    {day:%d.%m.%Y}: было {was:,.2f} → стало {now:,.2f} ({now - was:+,.2f})")

    return {
        "deleted": deleted,
        "appended": appended,
        "dropped": dropped,
        "unknown": unknown,
        "swept": swept,
        "changed_days": changed,
        "kept": appended - dropped,
        "first_row": start,
        "last_row": end - dropped,
        "rows_before": before_last - header_row,
    }
