"""Настройки проекта из .env — один объект на весь прогон."""
import os
from dataclasses import dataclass, field
from pathlib import Path

from dotenv import load_dotenv

PROJECT_ROOT = Path(__file__).resolve().parent.parent
load_dotenv(PROJECT_ROOT / ".env")


def _abs(path: str) -> str:
    p = Path(path)
    return str(p if p.is_absolute() else PROJECT_ROOT / p)


def _sheet_id(value: str) -> str:
    """Принимает и голый ID, и полную ссылку на таблицу."""
    value = (value or "").strip()
    if "/spreadsheets/d/" in value:
        value = value.split("/spreadsheets/d/", 1)[1].split("/", 1)[0]
    return value


@dataclass
class Config:
    # Google
    credentials_path: str
    spreadsheet_id: str
    source_xlsx_id: str
    register_sheet: str
    register_header_row: int
    summary_sheet: str
    summary_range: str

    # Telegram
    bot_token: str
    chat_ids: list
    log_chat_id: str
    admin_ids: list
    viewer_ids: list

    # 1С
    report_section: str
    report_menu_item: str
    report_title: str
    window_title: str
    reuse_open_1c: bool
    exclude_client_groups: list
    retail_divisions: list

    # Расписание
    run_at: str
    weekly_refresh_weekday: int
    max_chunk_days: int

    project_root: Path = field(default=PROJECT_ROOT)


def load_config() -> Config:
    chat_ids = [c.strip() for c in os.environ.get("TELEGRAM_CHAT_IDS", "").split(",") if c.strip()]
    admin_ids = [int(a) for a in os.environ.get("BOT_ADMIN_IDS", "").replace(" ", "").split(",") if a]
    # Наблюдатели: могут запросить скриншот и статус, но не запускают заливку.
    # По умолчанию это все получатели рассылки — им отчёт и так приходит,
    # логично разрешить попросить его же в любой момент.
    viewers = os.environ.get("BOT_VIEWER_IDS", "").replace(" ", "").split(",")
    viewer_ids = [int(v) for v in viewers if v.strip().lstrip("-").isdigit()]
    if not viewer_ids:
        viewer_ids = [int(c) for c in chat_ids if c.lstrip("-").isdigit()]
    return Config(
        credentials_path=_abs(os.environ.get("GOOGLE_CREDENTIALS_PATH", "credentials/service_account.json")),
        spreadsheet_id=_sheet_id(os.environ.get("REGISTER_SPREADSHEET_ID", "")),
        source_xlsx_id=_sheet_id(os.environ.get("SOURCE_XLSX_FILE_ID", "")),
        register_sheet=os.environ.get("REGISTER_SHEET", "Регистр"),
        register_header_row=int(os.environ.get("REGISTER_HEADER_ROW", "20")),
        summary_sheet=os.environ.get("SUMMARY_SHEET", "Sales by day"),
        summary_range=os.environ.get("SUMMARY_RANGE", "A19:K34"),
        bot_token=os.environ.get("TELEGRAM_BOT_TOKEN", ""),
        chat_ids=chat_ids,
        log_chat_id=os.environ.get("TELEGRAM_LOG_CHAT_ID", "").strip(),
        admin_ids=admin_ids,
        viewer_ids=viewer_ids,
        report_section=os.environ.get("REPORT_SECTION", "Продажи"),
        report_menu_item=os.environ.get("REPORT_MENU_ITEM", "Отчеты по продажам"),
        report_title=os.environ.get("REPORT_TITLE", "Валовая прибыль организаций - копия"),
        window_title=os.environ.get("ONEC_WINDOW_TITLE", "1С:Предприятие"),
        reuse_open_1c=os.environ.get("ONEC_REUSE_OPEN", "0").strip() in ("1", "true", "yes"),
        exclude_client_groups=[g.strip() for g in
                               os.environ.get("EXCLUDE_CLIENT_GROUPS", "B2B,B2B new,optom").split(",")
                               if g.strip()],
        # Подразделения, которые в книге считаются «Розницей». Формула
        # «Группа» (столбец J) перечисляет магазины поимённо, и новый
        # магазин она не знает — строки остаются без канала и удаляются.
        # Открылся магазин — допиши его сюда, остальное код сделает сам
        # (см. sheets_register.ensure_group_formula).
        retail_divisions=[d.strip() for d in
                          os.environ.get("RETAIL_DIVISIONS", "").split(",")
                          if d.strip()],
        run_at=os.environ.get("RUN_AT", "06:30"),
        weekly_refresh_weekday=int(os.environ.get("WEEKLY_REFRESH_WEEKDAY", "0")),
        max_chunk_days=int(os.environ.get("MAX_CHUNK_DAYS", "7")),
    )
