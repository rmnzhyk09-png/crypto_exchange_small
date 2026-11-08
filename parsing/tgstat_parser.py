#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Механика парсинга TGStat, по которой собраны выгрузки в ../out.

Оригинальные скрипты в папке проекта не сохранились. Этот файл —
реконструкция пайплайна: он повторяет шаги, которые видны по структуре
выгрузок и по рабочему парсеру tg_parser_bot (Playwright + CDP к Edge).

Пайплайн:
  1. Открыть https://tgstat.ru/channels/search (или /chats/search).
  2. Выставить фильтры: категория, язык, диапазон подписчиков, ключевое слово.
  3. Собрать карточки каналов/чатов из выдачи.
  4. Для каждой карточки открыть страницу: название, описание, гео, язык,
     контакты в описании.
  5. Открыть /stat: ERR%, ERR24%, охват 1 поста.
  6. Разметить LIVE / DEAD и выгрузить в Excel + TXT.

Запуск (нужен Edge с --remote-debugging-port=9222, см. README):
  python parsing/tgstat_parser.py --query "crypto" --language 12 \
      --subs-from 100000 --subs-to 2000000 --out out/parsed.xlsx
"""

import argparse
import asyncio
import random
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

from playwright.async_api import async_playwright

CDP_URL = "http://127.0.0.1:9222"
BASE = "https://tgstat.ru"

# Служебные/рекламные аккаунты, которые не считаем контактом админа.
SERVICE_HANDLES = {
    "tgstat", "tgstat_chat", "telepulse", "tgstatapi", "tgstat_bot",
    "searcheebot", "tgalertsbot", "tg_analytics_bot", "tgstatchatbot",
    "joinchat",
}
ADMIN_WORDS = re.compile(
    r"admin|админ|contact|контакт|support|поддержк|help|менеджер|вопрос|"
    r"пишите|write|связь|телеграм|telegram|dm|лс|владелец|создатель|"
    r"брать|заказать|куп|продаж|обмен|администратор",
    re.I,
)

# Собираем карточки из выдачи: ссылка, тип, заголовок и текст карточки.
JS_CARDS = """() => {
  const out = [];
  document.querySelectorAll("a[href*='/channel/'], a[href*='/chat/']").forEach(a => {
    const href = a.getAttribute("href") || "";
    if (!/\\/(channel|chat)\\/(@[^/]+)/.test(href)) return;
    const card = a.closest(".col") || a.parentElement;
    out.push({
      href,
      kind: href.includes("/chat/") ? "chat" : "channel",
      title: ((card && card.innerText) || "").split("\\n")[0].trim(),
      txt: ((card && card.innerText) || "").replace(/\\s+/g, " ").slice(0, 300),
    });
  });
  return out;
}"""


class Tgstat:
    def __init__(self, page):
        self.page = page

    async def _text(self):
        return await self.page.evaluate("document.body.innerText")

    async def _check_captcha(self):
        txt = await self._text()
        if "Подозрение на робота" in txt or "проверки безопасности" in txt:
            raise RuntimeError("TGStat показал капчу — увеличьте паузу или смените IP")
        return txt

    async def search(self, query, category, language, subs_from, subs_to, kind="channel"):
        """Один поисковый запрос. Возвращает список карточек."""
        path = "channels/search" if kind == "channel" else "chats/search"
        await self.page.goto(f"{BASE}/{path}", timeout=90_000)
        await self.page.wait_for_timeout(3_500)
        await self.page.wait_for_selector("select[name='categories[]']", timeout=15_000)
        await self.page.select_option("select[name='categories[]']", category)
        try:
            await self.page.select_option("select[name='countries[]']", [])
        except Exception:
            pass
        await self.page.select_option("select[name='languages[]']", str(language))
        await self.page.fill("input[name='participantsCountFrom']", str(subs_from))
        await self.page.fill("input[name='participantsCountTo']", str(subs_to))
        try:
            await self.page.check("input[name='inAbout']", timeout=2_000)
        except Exception:
            pass
        await self.page.fill("input[name='q']", query)
        await self.page.locator("input[name='q']").press("Enter")
        await self.page.wait_for_timeout(7_000)
        await self._check_captcha()
        return await self.page.evaluate(JS_CARDS)

    async def channel(self, kind, username):
        """Страница канала/чата: название, описание, гео, язык, контакты."""
        await self.page.goto(f"{BASE}/{kind}/@{username}", timeout=60_000)
        await self.page.wait_for_timeout(2_000)
        txt = await self._check_captcha()
        lines = [line.strip() for line in txt.split("\n") if line.strip()]

        rec = {"title": lines[1] if len(lines) > 1 else ""}
        try:
            start = next(i for i, line in enumerate(lines) if line.startswith("@"))
            end = next(
                i for i, line in enumerate(lines[start:], start)
                if line == "Связанные каналы"
            )
            rec["description"] = " ".join(lines[start + 1:end])[:1_000]
        except Exception:
            rec["description"] = ""

        rec["geo"] = rec["language"] = ""
        try:
            i = next(i for i, line in enumerate(lines)
                     if re.fullmatch(r"Гео и язык (канала|чата)", line))
            geo_line = lines[i + 1] if i + 1 < len(lines) else ""
            parts = [p.strip() for p in geo_line.split(",")]
            rec["geo"] = parts[0] if parts else ""
            rec["language"] = parts[1] if len(parts) > 1 else ""
        except Exception:
            pass

        contacts, seen = [], set()
        for match in re.finditer(r"@([A-Za-z0-9_]{4,32})", rec["description"]):
            handle = match.group(1)
            if handle.lower() in SERVICE_HANDLES or handle.lower() in seen:
                continue
            seen.add(handle.lower())
            ctx = rec["description"][max(0, match.start() - 100):match.end() + 40]
            contacts.append({
                "handle": handle,
                "in_admin_context": bool(ADMIN_WORDS.search(ctx)),
            })
        rec["contacts"] = contacts
        rec["admin_contact"] = next(
            (c["handle"] for c in contacts if c["in_admin_context"]),
            contacts[0]["handle"] if contacts else "",
        )
        return rec

    async def stats(self, kind, username):
        """Страница /stat: ERR%, ERR24%, охват одного поста."""
        if kind == "chat":
            return {}
        await self.page.goto(f"{BASE}/{kind}/@{username}/stat", timeout=60_000)
        await self.page.wait_for_timeout(2_000)
        txt = await self._check_captcha()

        rec = {}
        for name, pattern in (
            ("err", r"([\d.,]+)%\s*ERR\b"),
            ("err24", r"([\d.,]+)%\s*ERR24\b"),
        ):
            match = re.search(pattern, txt)
            if match:
                try:
                    rec[name] = float(match.group(1).replace(",", "."))
                except ValueError:
                    pass

        lines = [line.strip() for line in txt.split("\n") if line.strip()]
        for i, line in enumerate(lines):
            if "охват 1 поста" in line.lower() or "просмотров на пост" in line.lower():
                nums = re.findall(r"\d[\d\s\u00a0]*", line)
                if nums:
                    rec["views_per_post"] = int(re.sub(r"\D", "", nums[0]))
                elif i > 0:
                    nums = re.findall(r"\d[\d\s\u00a0]*", lines[i - 1])
                    if nums:
                        rec["views_per_post"] = int(re.sub(r"\D", "", nums[0]))
                break
        return rec


def parse_subs(text):
    match = re.search(r"([\d\s\u00a0]{1,15})\s*(?:подписчик|участник)", text)
    return int(re.sub(r"\D", "", match.group(1))) if match else None


def classify(rec, err_min, views_min):
    subs = rec.get("subscribers")
    err = rec.get("err")
    views = rec.get("views_per_post")
    rec["is_live"] = bool(
        subs and err is not None and err_min <= err <= 100
        and views is not None and views >= views_min
    )
    rec["is_dead"] = bool(views is not None and views < views_min)
    return rec


async def enrich(tg, cards, err_min, views_min, delay_min, delay_max, limit):
    rows, seen = [], set()
    for i, card in enumerate(cards[:limit], 1):
        match = re.search(r"/(channel|chat)/(@[^/]+)/", card["href"] + "/")
        if not match:
            continue
        kind, username = match.group(1), match.group(2)
        if username.lower() in seen:
            continue
        seen.add(username.lower())

        rec = {
            "url": f"https://t.me/{username.lstrip('@')}",
            "tgstat_url": f"{BASE}/{kind}/{username}",
            "type": kind,
            "username": username,
            "subscribers": parse_subs(card["txt"]),
            "query": card.get("query", ""),
        }
        try:
            rec.update(await tg.channel(kind, username))
        except Exception as exc:
            rec["parse_error"] = str(exc)[:200]
        try:
            rec.update(await tg.stats(kind, username))
        except Exception as exc:
            rec["stat_error"] = str(exc)[:200]

        classify(rec, err_min, views_min)
        rec["collected_at"] = datetime.now(timezone.utc).isoformat()
        rows.append(rec)
        print(
            f"[{i}/{min(len(cards), limit)}] {username} "
            f"subs={rec.get('subscribers')} err={rec.get('err')} "
            f"views={rec.get('views_per_post')} live={rec.get('is_live')}",
            flush=True,
        )
        await asyncio.sleep(random.uniform(delay_min, delay_max))
    return rows


def export_xlsx(rows, path):
    from openpyxl import Workbook
    from openpyxl.styles import Font

    columns = [
        ("Ссылка", "url"), ("TGStat", "tgstat_url"), ("Тип", "type"),
        ("Название", "title"), ("Подписчики", "subscribers"),
        ("ERR%", "err"), ("ERR24%", "err24"), ("Просм/пост", "views_per_post"),
        ("Язык", "language"), ("Гео", "geo"),
        ("Контакт админа", "admin_contact"), ("Все контакты", "contacts"),
        ("Описание", "description"), ("LIVE", "is_live"),
        ("Запрос", "query"), ("Собрано", "collected_at"),
    ]
    wb = Workbook()
    ws = wb.active
    ws.title = "parsed"
    ws.append([title for title, _ in columns])
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for rec in rows:
        values = []
        for _, key in columns:
            value = rec.get(key)
            if isinstance(value, list):
                value = ", ".join(
                    f"@{c['handle']}" if isinstance(c, dict) else str(c)
                    for c in value
                )
            values.append(value)
        ws.append(values)
    ws.freeze_panes = "A2"
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def export_links(rows, path):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    Path(path).write_text(
        "\n".join(rec["url"] for rec in rows) + "\n", encoding="utf-8"
    )
    return path


async def main():
    parser = argparse.ArgumentParser(description="Парсер TGStat (реконструкция)")
    parser.add_argument("--query", action="append", required=True,
                        help="ключевое слово, можно передать несколько раз")
    parser.add_argument("--category", default="crypto")
    parser.add_argument("--language", default="12",
                        help="код языка TGStat: 3 — русский, 12 — английский")
    parser.add_argument("--kind", default="channel", choices=["channel", "chat"])
    parser.add_argument("--subs-from", type=int, default=300)
    parser.add_argument("--subs-to", type=int, default=9_000)
    parser.add_argument("--limit", type=int, default=30)
    parser.add_argument("--err-min", type=float, default=16.0)
    parser.add_argument("--views-min", type=int, default=150)
    parser.add_argument("--delay-min", type=float, default=45)
    parser.add_argument("--delay-max", type=float, default=75)
    parser.add_argument("--out", default="out/parsed.xlsx")
    args = parser.parse_args()

    playwright = await async_playwright().start()
    browser = await playwright.chromium.connect_over_cdp(CDP_URL)
    context = browser.contexts[0]
    page = context.pages[0] if context.pages else await context.new_page()
    tg = Tgstat(page)

    try:
        cards = []
        for query in args.query:
            found = await tg.search(
                query, args.category, args.language,
                args.subs_from, args.subs_to, args.kind,
            )
            print(f"Запрос '{query}': найдено {len(found)} карточек", flush=True)
            for card in found:
                card["query"] = query
            cards.extend(found)

        rows = await enrich(
            tg, cards, args.err_min, args.views_min,
            args.delay_min, args.delay_max, args.limit,
        )
        xlsx = export_xlsx(rows, args.out)
        links = export_links(rows, str(Path(args.out).with_suffix(".txt")))
        print(f"Готово: {xlsx} ({len(rows)} строк), {links}", flush=True)
    finally:
        await playwright.stop()


if __name__ == "__main__":
    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding="utf-8")
    asyncio.run(main())
