"""
Конвертация книги-регистра из .xlsx (файл Office на Google Диске) в
НАТИВНУЮ Google-таблицу — без этого Sheets API писать в неё не может
(«This operation is not supported for this document. The document must not
be an Office file»).

Делает files.copy с преобразованием: исходный .xlsx остаётся нетронутым,
рядом появляется его Google-версия со всеми листами, формулами, форматами
и диаграммами.

Требуется включённый Drive API в проекте сервис-аккаунта:
  https://console.developers.google.com/apis/api/drive.googleapis.com/overview

    python tools/convert_to_gsheet.py
    python tools/convert_to_gsheet.py --name "Sales by day (Google)"
    python tools/convert_to_gsheet.py --share you@example.com

После конвертации пропиши новый ID в .env → REGISTER_SPREADSHEET_ID.
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(errors="replace")

import requests
from google.auth.transport.requests import Request
from google.oauth2.service_account import Credentials

from src.config import load_config

SCOPES = ["https://www.googleapis.com/auth/drive"]
GSHEET_MIME = "application/vnd.google-apps.spreadsheet"


def _token(credentials_path: str) -> str:
    creds = Credentials.from_service_account_file(credentials_path, scopes=SCOPES)
    creds.refresh(Request())
    return creds.token


def _explain(resp) -> str:
    try:
        error = resp.json()["error"]
    except Exception:  # noqa: BLE001
        return resp.text[:500]
    message = error.get("message", "")
    if "has not been used in project" in message or "disabled" in message:
        return (
            message
            + "\n\nВключи Drive API одной кнопкой по ссылке из сообщения выше и повтори "
              "запуск (изменение применяется за минуту-две)."
        )
    return message


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--name", help="имя новой таблицы")
    parser.add_argument("--share", action="append", default=[],
                        help="кому дать права редактора (можно указать несколько раз)")
    args = parser.parse_args()

    cfg = load_config()
    token = _token(cfg.credentials_path)
    headers = {"Authorization": f"Bearer {token}"}

    meta = requests.get(
        f"https://www.googleapis.com/drive/v3/files/{cfg.source_xlsx_id}",
        params={"fields": "id,name,mimeType", "supportsAllDrives": "true"},
        headers=headers, timeout=60,
    )
    if not meta.ok:
        raise SystemExit(f"Не читается исходный файл:\n{_explain(meta)}")
    info = meta.json()
    print(f"Исходник: {info['name']}  ({info['mimeType']})")

    if info["mimeType"] == GSHEET_MIME:
        print("Это уже нативная Google-таблица — конвертировать нечего.")
        print(f"REGISTER_SPREADSHEET_ID={info['id']}")
        return

    name = args.name or f"{info['name']} (Google)"
    resp = requests.post(
        f"https://www.googleapis.com/drive/v3/files/{cfg.source_xlsx_id}/copy",
        params={"supportsAllDrives": "true", "fields": "id,name"},
        headers=headers,
        json={"name": name, "mimeType": GSHEET_MIME},
        timeout=300,
    )
    if not resp.ok:
        raise SystemExit(f"Конвертация не удалась:\n{_explain(resp)}")
    new = resp.json()
    print(f"Готово: {new['name']}")
    print(f"URL: https://docs.google.com/spreadsheets/d/{new['id']}/edit")

    for email in args.share:
        share = requests.post(
            f"https://www.googleapis.com/drive/v3/files/{new['id']}/permissions",
            params={"sendNotificationEmail": "false", "supportsAllDrives": "true"},
            headers=headers,
            json={"type": "user", "role": "writer", "emailAddress": email},
            timeout=60,
        )
        print(f"  доступ для {email}: {'ок' if share.ok else _explain(share)}")

    print("\nТеперь пропиши в .env:")
    print(f"  REGISTER_SPREADSHEET_ID={new['id']}")


if __name__ == "__main__":
    main()
