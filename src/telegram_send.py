"""Отправка картинки/текста ежедневного отчёта в Telegram."""
import time

import requests

API = "https://api.telegram.org/bot{token}/{method}"


def _call(token: str, method: str, retries: int = 4, **kwargs):
    """Вызов Bot API с повтором: сеть и 429 Too Many Requests — обычное дело."""
    last = None
    for attempt in range(1, retries + 1):
        try:
            resp = requests.post(API.format(token=token, method=method), timeout=90, **kwargs)
            data = resp.json()
            if data.get("ok"):
                return data["result"]
            # 429: Telegram сам говорит, сколько ждать.
            if resp.status_code == 429:
                wait = int(data.get("parameters", {}).get("retry_after", 5))
                print(f"  ! Telegram 429 — жду {wait} c")
                time.sleep(wait + 1)
                continue
            last = RuntimeError(f"{method}: {data.get('description')}")
        except requests.RequestException as e:
            last = e
        time.sleep(min(3 * attempt, 15))
    raise last or RuntimeError(f"{method}: не удалось вызвать")


def send_photo(token: str, chat_id: str, photo_path: str, caption: str = "") -> None:
    with open(photo_path, "rb") as fh:
        _call(
            token,
            "sendPhoto",
            data={"chat_id": chat_id, "caption": caption[:1024], "parse_mode": "HTML"},
            files={"photo": fh},
        )


def send_document(token: str, chat_id: str, path: str, caption: str = "") -> None:
    """Картинка «как документ» — Telegram не сжимает, мелкие цифры остаются читаемыми."""
    with open(path, "rb") as fh:
        _call(
            token,
            "sendDocument",
            data={"chat_id": chat_id, "caption": caption[:1024], "parse_mode": "HTML"},
            files={"document": fh},
        )


def send_text(token: str, chat_id: str, text: str) -> None:
    _call(token, "sendMessage", data={"chat_id": chat_id, "text": text[:4096], "parse_mode": "HTML"})


def broadcast_photo(token: str, chat_ids, photo_path: str, caption: str = "",
                    as_document: bool = False) -> dict:
    """
    Рассылка всем получателям. Падение на одном получателе (заблокировал
    бота, не нажал /start) не должно отменять отправку остальным — поэтому
    ошибки собираются, а не выбрасываются сразу.
    """
    sent, failed = [], {}
    for chat_id in chat_ids:
        try:
            if as_document:
                send_document(token, chat_id, photo_path, caption)
            else:
                send_photo(token, chat_id, photo_path, caption)
            sent.append(chat_id)
        except Exception as e:  # noqa: BLE001 — остальные получатели важнее
            failed[chat_id] = str(e)
            print(f"  ! Не отправилось в {chat_id}: {e}")
    return {"sent": sent, "failed": failed}


def broadcast_text(token: str, chat_ids, text: str) -> dict:
    sent, failed = [], {}
    for chat_id in chat_ids:
        try:
            send_text(token, chat_id, text)
            sent.append(chat_id)
        except Exception as e:  # noqa: BLE001
            failed[chat_id] = str(e)
            print(f"  ! Не отправилось в {chat_id}: {e}")
    return {"sent": sent, "failed": failed}


def chat_access(token: str, chat_ids) -> dict:
    """
    Может ли бот писать каждому получателю: {chat_id: (доступен, описание)}.

    Спрашиваем getChat, а не getUpdates: обновления вычитывает работающий
    бот, и getUpdates у проверки всегда пуст — из-за этого она показывала
    «не виден» даже тем, кто боту писал. getChat отвечает по факту: если
    человек не нажимал Start, Telegram вернёт «chat not found», и отправка
    ему тоже не пройдёт.
    """
    out = {}
    for chat_id in chat_ids:
        try:
            resp = requests.get(API.format(token=token, method="getChat"),
                                params={"chat_id": chat_id}, timeout=30).json()
        except requests.RequestException as e:
            out[str(chat_id)] = (False, f"сеть: {e}")
            continue
        if resp.get("ok"):
            chat = resp["result"]
            name = (chat.get("title")
                    or " ".join(filter(None, [chat.get("first_name"), chat.get("last_name")]))
                    or chat.get("username") or "")
            out[str(chat_id)] = (True, f"{chat.get('type', '')} {name}".strip())
        else:
            out[str(chat_id)] = (False, resp.get("description", "недоступен"))
    return out


def known_chats(token: str) -> dict:
    """
    chat_id всех, кто писал боту (getUpdates) — чтобы узнать id новой группы
    или проверить, нажал ли получатель /start.
    """
    resp = requests.get(API.format(token=token, method="getUpdates"), timeout=30).json()
    out = {}
    for update in resp.get("result", []):
        message = (
            update.get("message")
            or update.get("channel_post")
            or update.get("my_chat_member")
            or {}
        )
        chat = message.get("chat")
        if chat:
            out[str(chat["id"])] = chat.get("title") or chat.get("username") or chat.get("first_name", "")
    return out
