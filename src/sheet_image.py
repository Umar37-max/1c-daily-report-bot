"""
Скриншот диапазона вкладки Google Sheets — настоящий вид листа (заливка,
границы, форматы), а не перерисованные данные.

Механизм тот же, что в соседнем проекте cashflow1c (src/sheets_export.py):
Sheets API картинку диапазона не отдаёт, поэтому используем встроенный
экспорт самой таблицы в PDF (то же, что «Файл → Печать» в браузере) и
растеризуем его в PNG через PyMuPDF. Авторизация — тем же сервис-аккаунтом.
"""
import io
import re

import fitz
import requests
from google.auth.transport.requests import Request
from google.oauth2.service_account import Credentials
from PIL import Image, ImageChops

EXPORT_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
RENDER_ZOOM = 2.0  # растеризация x2 — чтобы в Telegram читались цифры

_creds_cache = {}


def _access_token(credentials_path: str) -> str:
    creds = _creds_cache.get(credentials_path)
    if creds is None:
        creds = Credentials.from_service_account_file(credentials_path, scopes=EXPORT_SCOPES)
        _creds_cache[credentials_path] = creds
    if not creds.valid:
        creds.refresh(Request())
    return creds.token


HEADER_LABEL = "Подразделение"
TOTAL_LABEL = "Итого"
_LETTER_RE = re.compile(r"[A-Za-zА-Яа-яЁё]")


def _column_letter(number: int) -> str:
    """1 → A, 10 → J, 27 → AA."""
    letters = ""
    while number:
        number, rest = divmod(number - 1, 26)
        letters = chr(ord("A") + rest) + letters
    return letters


def find_block_range(values, header_label: str = HEADER_LABEL,
                     total_label: str = TOTAL_LABEL):
    """
    Находит на листе блок «Sales by day» и отдаёт его диапазон A1 (или None).

    Раньше диапазон был зашит в .env (A19:K34). ЖИВЬЁМ 27.10.2026 коллеги
    дописали строки в верхнюю таблицу (новый магазин, маркетплейс и др.), блок уехал
    вниз на 7 строк — и в Telegram пошла картинка с хвостом чужой таблицы
    сверху и обрезанным низом: без части магазинов, маркетплейсов и
    «Итого». Поэтому ищем блок по его содержимому:

      * верх — строка с «Подразделение» в столбце A (последняя такая на
        листе), плюс строка над ней, если там дата дня;
      * низ — строка «Итого» под шапкой и примечания сразу под ней
        (маржа в процентах), пока они идут подряд;
      * ширина — до последнего ПОДПИСАННОГО столбца шапки («Темп»):
        правее коллеги держат служебные расчёты без шапки и границ.
    """
    rows = [list(row) for row in values]
    header_idx = None
    for index, row in enumerate(rows):
        if row and str(row[0]).strip() == header_label:
            header_idx = index
    if header_idx is None:
        return None

    header = rows[header_idx]
    last_col = max((col for col, cell in enumerate(header) if _LETTER_RE.search(str(cell))),
                   default=0)
    if last_col == 0:
        return None

    def filled(row):
        return any(str(cell).strip() for cell in row[:last_col + 1])

    top_idx = header_idx
    if header_idx > 0 and filled(rows[header_idx - 1]) and not str(rows[header_idx - 1][0]).strip():
        top_idx = header_idx - 1  # дата дня над шапкой

    end_idx = None
    for index in range(header_idx + 1, len(rows)):
        label = str(rows[index][0]).strip() if rows[index] else ""
        if label.startswith(total_label):
            end_idx = index
            break
        if not filled(rows[index]):
            end_idx = index - 1  # «Итого» нет — блок кончился пустой строкой
            break
    if end_idx is None:
        end_idx = len(rows) - 1

    # Примечания под «Итого» (маржа 45% / 11%) — часть блока.
    for _ in range(2):
        nxt = end_idx + 1
        if nxt < len(rows) and filled(rows[nxt]) and not str(rows[nxt][0]).strip():
            end_idx = nxt
        else:
            break

    if end_idx <= header_idx:
        return None
    return f"A{top_idx + 1}:{_column_letter(last_col + 1)}{end_idx + 1}"


def resolve_range(spreadsheet, sheet_name: str, configured: str) -> str:
    """Диапазон блока, найденный на листе; если не нашёлся — тот, что в .env."""
    try:
        found = find_block_range(spreadsheet.worksheet(sheet_name).get_all_values())
    except Exception as e:  # noqa: BLE001 — без автопоиска всё равно снимем по .env
        print(f"  ! Не смог найти блок на листе «{sheet_name}»: {e}")
        found = None
    if found:
        if found != configured:
            print(f"  Блок «{sheet_name}» найден на листе: {found} (в .env {configured})")
        return found
    print(f"  ! Блок с шапкой «{HEADER_LABEL}» не найден — снимаю диапазон из .env: {configured}")
    return configured


def sheet_gid(spreadsheet, sheet_name: str) -> int:
    """gid вкладки по её названию (нужен экспорту)."""
    return spreadsheet.worksheet(sheet_name).id


def export_pdf(credentials_path: str, spreadsheet_id: str, gid: int, a1_range: str,
               landscape: bool = True, gridlines: bool = False) -> bytes:
    url = (
        f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export"
        f"?format=pdf&gid={gid}&range={a1_range}"
        f"&portrait={'false' if landscape else 'true'}&fitw=true"
        f"&top_margin=0.00&bottom_margin=0.00&left_margin=0.00&right_margin=0.00"
        f"&gridlines={'true' if gridlines else 'false'}&printtitle=false&sheetnames=false"
        f"&pagenum=UNDEFINED&attachment=false"
        # fzr/fzc=false — не повторять закреплённые строки/столбцы листа.
        # Без этого поверх запрошенного диапазона A19:K34 печатается ещё и
        # «шапка» листа (строки 3–4 с датами 9/1…9/5 и Sales/Gross profit),
        # и картинка приходит с чужой полосой сверху — поймано живьём
        # 20.09.2026 на первом же реальном скриншоте.
        f"&fzr=false&fzc=false"
    )
    token = _access_token(credentials_path)
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    resp.raise_for_status()
    content_type = resp.headers.get("content-type", "").split(";")[0].strip()
    if content_type != "application/pdf":
        raise RuntimeError(f"Экспорт вкладки вернул не PDF (content-type={content_type!r})")
    return resp.content


def _autocrop(img: Image.Image) -> Image.Image:
    """
    Страница PDF всегда фиксированного размера (A4/letter), даже если в
    диапазоне 15 строк — белое поле снизу занимает почти всю страницу.
    Режем по реальному содержимому.
    """
    bg = Image.new(img.mode, img.size, (255, 255, 255))
    bbox = ImageChops.difference(img, bg).getbbox()
    return img.crop(bbox) if bbox else img


def pdf_to_image(pdf_bytes: bytes) -> Image.Image:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        pages = []
        for page in doc:
            pix = page.get_pixmap(matrix=fitz.Matrix(RENDER_ZOOM, RENDER_ZOOM))
            pages.append(_autocrop(Image.open(io.BytesIO(pix.tobytes("png")))))
    finally:
        doc.close()

    if not pages:
        raise RuntimeError("Экспортированный PDF без страниц")
    if len(pages) == 1:
        return pages[0]

    width = max(p.width for p in pages)
    combined = Image.new("RGB", (width, sum(p.height for p in pages)), (255, 255, 255))
    y = 0
    for page_img in pages:
        combined.paste(page_img, (0, y))
        y += page_img.height
    return combined


def _add_margin(img: Image.Image, margin: int = 24) -> Image.Image:
    """Белые поля вокруг — иначе в Telegram таблица упирается в край превью."""
    out = Image.new("RGB", (img.width + margin * 2, img.height + margin * 2), (255, 255, 255))
    out.paste(img, (margin, margin))
    return out


def capture(credentials_path: str, spreadsheet, sheet_name: str, a1_range: str,
            out_path: str, auto: bool = True) -> str:
    """
    Снимок блока вкладки → PNG-файл. Возвращает путь.

    auto=True — диапазон ищется по самой таблице (см. find_block_range), а
    a1_range из .env остаётся запасным вариантом.
    """
    if auto:
        a1_range = resolve_range(spreadsheet, sheet_name, a1_range)
    gid = sheet_gid(spreadsheet, sheet_name)
    pdf = export_pdf(credentials_path, spreadsheet.id, gid, a1_range)
    img = _add_margin(pdf_to_image(pdf))
    img.save(out_path, "PNG", optimize=True)
    return out_path
