"""
Рабочая дата — всегда по часовому поясу бизнеса, а не по локальному
времени машины.

ЖИВЬЁМ (21.09.2026): сервер работает в поясе на 12 часов позади Ташкента
(в Планировщике «9/20/2026 5:22 AM», когда в Ташкенте было 20.09 17:22).
Кнопка «Обновить за вчера», нажатая 21.09 в 07:43 по Ташкенту, попадала на
сервере в 20.09 19:43 — date.today() отдавала 20-е, «вчера» становилось
19-м, и прогон заново перезаливал позавчерашний день вместо вчерашнего.

Пояс берётся из .env (TIMEZONE), по умолчанию Asia/Tashkent. Расписание
бота уже привязано к тому же поясу, так что кнопка и автозапуск считают
дату одинаково.
"""
import os
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

from dotenv import load_dotenv

load_dotenv()

TZ = ZoneInfo(os.environ.get("TIMEZONE", "Asia/Tashkent"))


def now() -> datetime:
    return datetime.now(TZ)


def today() -> date:
    return now().date()


def yesterday() -> date:
    return today() - timedelta(days=1)
