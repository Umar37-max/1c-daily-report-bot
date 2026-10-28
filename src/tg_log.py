"""
Отправка хода работы (логов) синхронизаций в Telegram-группу, чтобы за
прогоном можно было следить, не заходя на сервер.

Куда шлём: TELEGRAM_LOG_CHAT_ID из .env (группа, куда добавлен бот).
Чем шлём: тот же TELEGRAM_BOT_TOKEN, что и у bot.py — отдельный бот не нужен.

Устройство намеренно простое и "неломающееся":
  * любая сетевая ошибка/таймаут телеграма НИКОГДА не роняет прогон —
    логирование это побочная функция, а не сама работа (ошибку печатаем в
    консоль и идём дальше);
  * 429 (Too Many Requests) уважаем: ждём retry_after и пробуем ещё раз —
    длинный прогон легко упирается в лимит групп (~20 сообщений в минуту);
  * длинные тексты режем по лимиту Telegram (4096 символов).

Использование:
    from src.tg_log import TelegramLog
    tg = TelegramLog("Полный прогон")
    tg.step("Контрагенты")          # заголовок шага
    tg.info("38 контрагентов")      # строка хода
    tg.ok("Контрагенты — готово")   # успех
    tg.fail("Касса", exc)           # ошибка с текстом исключения
"""
import os
import sys
import time
import traceback
from datetime import datetime
from zoneinfo import ZoneInfo

import requests
from dotenv import load_dotenv

load_dotenv()

_TASHKENT = ZoneInfo("Asia/Tashkent")
_API = "https://api.telegram.org/bot{token}/sendMessage"
_MAX_LEN = 4000  # с запасом от телеграмных 4096


def _now() -> str:
    return datetime.now(_TASHKENT).strftime("%H:%M:%S")


def _print(text: str) -> None:
    """
    Печать в консоль, которая не падает на Windows-терминале с кодировкой
    cp1251: в сообщениях есть эмодзи (🚀/✅/❌), и обычный print() на такой
    консоли выбрасывает UnicodeEncodeError — то есть логирование уронило бы
    сам прогон, ради которого оно и существует.
    """
    try:
        print(text)
    except UnicodeEncodeError:
        encoding = getattr(sys.stdout, "encoding", "ascii") or "ascii"
        print(text.encode(encoding, errors="replace").decode(encoding, errors="replace"))


class TelegramLog:
    """
    Лог одного прогона. Если в .env нет TELEGRAM_LOG_CHAT_ID или токена —
    молча работает "вхолостую" (только печать в консоль), чтобы скрипты
    можно было запускать и без настроенного телеграма.
    """

    def __init__(self, title: str, chat_id: str | int | None = None, token: str | None = None):
        self.title = title
        self.token = token or os.environ.get("TELEGRAM_BOT_TOKEN", "")
        chat = chat_id if chat_id is not None else os.environ.get("TELEGRAM_LOG_CHAT_ID", "")
        self.chat_id = str(chat).strip()
        self.enabled = bool(self.token and self.chat_id)
        self.started = time.time()
        if not self.enabled:
            _print("[tg] Логирование в Telegram выключено (нет TELEGRAM_BOT_TOKEN/TELEGRAM_LOG_CHAT_ID в .env).")

    # ─────────────────────────── отправка ───────────────────────────

    def send(self, text: str) -> None:
        """Отправить произвольный текст. Никогда не бросает исключений."""
        _print(text)
        if not self.enabled:
            return
        body = text if len(text) <= _MAX_LEN else text[:_MAX_LEN] + "\n…(обрезано)"
        for attempt in range(3):
            try:
                resp = requests.post(
                    _API.format(token=self.token),
                    json={
                        "chat_id": self.chat_id,
                        "text": body,
                        "parse_mode": "HTML",
                        "disable_web_page_preview": True,
                    },
                    timeout=20,
                )
                if resp.status_code == 429:
                    wait = resp.json().get("parameters", {}).get("retry_after", 5)
                    time.sleep(min(int(wait) + 1, 60))
                    continue
                if not resp.ok:
                    _print(f"[tg] Telegram ответил {resp.status_code}: {resp.text[:200]}")
                return
            except Exception as e:  # noqa: BLE001 — лог не должен ронять прогон
                _print(f"[tg] Не удалось отправить лог (попытка {attempt + 1}/3): {e}")
                time.sleep(2)

    # ─────────────────────────── уровни ───────────────────────────

    def start(self, subtitle: str = "") -> None:
        extra = f"\n{subtitle}" if subtitle else ""
        self.send(f"🚀 <b>{self.title}</b> — старт в {_now()}{extra}")

    def step(self, name: str) -> None:
        self.send(f"▶️ <b>{_escape(name)}</b> — начал в {_now()}")

    def info(self, text: str) -> None:
        self.send(f"• {_escape(text)}")

    def ok(self, name: str, detail: str = "") -> None:
        extra = f"\n{detail}" if detail else ""
        self.send(f"✅ <b>{_escape(name)}</b> — готово в {_now()}{extra}")

    def fail(self, name: str, error: BaseException) -> None:
        """
        Ошибка ПОНЯТНЫМ текстом: тип и сообщение крупно, а технический
        traceback — последними строками (чтобы по логу в телеграме можно
        было понять, что случилось, не заходя на сервер).
        """
        # Текст самого исключения тоже экранируем: в нём попадаются угловые
        # скобки (напр. "Could not click '<unreadable>'"), и Telegram
        # считает их HTML-тегом — сообщение отлетает с 400 "Unsupported
        # start tag", то есть об ошибке никто не узнаёт.
        tail = "".join(traceback.format_exception(type(error), error, error.__traceback__))[-1200:]
        self.send(
            f"❌ <b>{_escape(name)}</b> — ошибка в {_now()}\n"
            f"<b>{_escape(type(error).__name__)}:</b> {_escape(str(error))[:800]}\n\n"
            f"<pre>{_escape(tail)}</pre>"
        )

    def finish(self, results: dict) -> None:
        """Итог всего прогона: что прошло, что упало, сколько заняло."""
        minutes = (time.time() - self.started) / 60
        lines = [f"🏁 <b>{self.title}</b> — завершено за {minutes:.1f} мин"]
        for name, status in results.items():
            lines.append(f"{'✅' if status is True else '❌'} {name}" if status is not None else f"⏭ {name} — пропущен")
        self.send("\n".join(lines))


def _escape(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


class StreamToTelegram:
    """
    Пересылает В ГРУППУ всё, что скрипты печатают в консоль, — чтобы в
    телеграме был виден КАЖДЫЙ шаг ("Ставлю контрагента...", "Формирую
    отчёт...", "Копирую данные..."), а не только начало и конец этапа.

    Строки копятся и уходят пачками: у групп лимит ~20 сообщений в минуту,
    а один прогон печатает сотни строк — по сообщению на строку telegram
    просто начал бы отвечать 429 и лог отстал бы на часы. Пачка уходит,
    когда накопилось max_lines строк или прошло flush_after секунд.

    Используется как контекст: with StreamToTelegram(log): ...
    """

    def __init__(self, log: TelegramLog, max_lines: int = 12, flush_after: float = 20.0):
        self.log = log
        self.max_lines = max_lines
        self.flush_after = flush_after
        self._buffer: list[str] = []
        self._last_flush = time.time()
        self._original = None

    def write(self, text: str) -> int:
        if self._original is not None:
            try:
                self._original.write(text)
            except Exception:
                pass
        for line in text.splitlines():
            line = line.strip()
            if line:
                self._buffer.append(line)
        if len(self._buffer) >= self.max_lines or (time.time() - self._last_flush) >= self.flush_after:
            self.flush_to_telegram()
        return len(text)

    def flush(self) -> None:
        if self._original is not None:
            try:
                self._original.flush()
            except Exception:
                pass

    def flush_to_telegram(self) -> None:
        if not self._buffer:
            self._last_flush = time.time()
            return
        chunk, self._buffer = self._buffer, []
        self._last_flush = time.time()
        # send() сам печатает в консоль — здесь это привело бы к рекурсии
        # (мы и есть stdout), поэтому шлём в телеграм напрямую.
        body = "<pre>" + _escape("\n".join(chunk)) + "</pre>"
        original, self.log_print_guard = self._original, True
        try:
            sys.stdout = original or sys.__stdout__
            self.log.send(body)
        finally:
            sys.stdout = self

    def isatty(self) -> bool:
        return False

    def __enter__(self):
        self._original = sys.stdout
        sys.stdout = self
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            self.flush_to_telegram()
        finally:
            sys.stdout = self._original
        return False
