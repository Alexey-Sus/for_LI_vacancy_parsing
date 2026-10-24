#!/usr/bin/env python3
"""
Мониторинг вакансий LinkedIn через публичный (гостевой) поиск.
Находит новые вакансии по поисковым запросам из config_old.json,
отсеивает уже виденные и нерелевантные, присылает новые в Telegram.

Запуск:  python linkedin_jobs.py
Зависимости:  pip install requests beautifulsoup4
"""

import json
import random
import re
import sys
import time
from pathlib import Path
from urllib.parse import urlencode

import requests
from bs4 import BeautifulSoup

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config_old.json"
SEEN_PATH = BASE_DIR / "seen_jobs.json"

SEARCH_URL = "https://www.linkedin.com/jobs-guest/jobs/api/seeMoreJobPostings/search"
DETAIL_URL = "https://www.linkedin.com/jobs-guest/jobs/api/jobPosting/{}"
PAGE_SIZE = 10          # гостевой поиск отдаёт около 10 вакансий за запрос
SEEN_LIMIT = 5000       # сколько ID вакансий помнить

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}


class RateLimited(Exception):
    """LinkedIn ответил 429 — пора остановиться."""


# ---------- вспомогательное ----------

def load_json(path, default):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return default


def save_json(path, data):
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def pause(cfg):
    lo, hi = cfg.get("delay_between_requests_sec", [4, 9])
    time.sleep(random.uniform(lo, hi))


def find_word(words, text):
    """
    Ищет слова из списка в тексте. Слово должно стоять целиком
    (допускается окончание -s): 'crypto' найдёт 'crypto' и 'cryptos',
    но не 'cryptography'; 'defi' не сработает на 'definitely'.
    Возвращает первое найденное слово или None.
    """
    text = text.lower()
    for w in words:
        pattern = r"(?<![a-z0-9])" + re.escape(w.lower()) + r"s?(?![a-z0-9])"
        if re.search(pattern, text):
            return w
    return None


# ---------- поиск и разбор ----------

def build_url(keywords, cfg, start):
    params = {
        "keywords": keywords,
        "location": cfg.get("location", "Worldwide"),
        "sortBy": "DD",                                         # сначала новые
        "f_TPR": f"r{int(cfg.get('posted_within_hours', 24)) * 3600}",
        "start": start,
    }
    if cfg.get("remote_only"):
        params["f_WT"] = "2"                                    # 2 = удалёнка
    return f"{SEARCH_URL}?{urlencode(params)}"


def parse_jobs(html):
    soup = BeautifulSoup(html, "html.parser")
    jobs = []
    for card in soup.select("div.base-card, div.job-search-card"):
        urn = card.get("data-entity-urn", "")
        job_id = urn.rsplit(":", 1)[-1] if urn else ""

        def text(sel):
            el = card.select_one(sel)
            return el.get_text(" ", strip=True) if el else ""

        link_el = card.select_one("a.base-card__full-link, a[href*='/jobs/view/']")
        link = link_el["href"].split("?")[0] if link_el and link_el.get("href") else ""
        if not job_id and link:
            job_id = link.rstrip("/").rsplit("-", 1)[-1]
        if not job_id:
            continue

        time_el = card.select_one("time")
        jobs.append({
            "id": job_id,
            "title": text("h3.base-search-card__title, h3"),
            "company": text("h4.base-search-card__subtitle, h4"),
            "location": text("span.job-search-card__location"),
            "posted": time_el.get_text(strip=True) if time_el else "",
            "link": link or f"https://www.linkedin.com/jobs/view/{job_id}/",
        })
    return jobs


def fetch_search(session, keywords, cfg):
    found = []
    for page in range(int(cfg.get("pages_per_search", 5))):
        url = build_url(keywords, cfg, page * PAGE_SIZE)
        try:
            resp = session.get(url, headers=HEADERS, timeout=20)
        except requests.RequestException as e:
            print(f"  ! сеть: {e}")
            break
        if resp.status_code == 429:
            print("  ! LinkedIn просит притормозить (429), останавливаю этот запрос")
            break
        if resp.status_code != 200:
            print(f"  ! ответ {resp.status_code}, пропускаю")
            break
        jobs = parse_jobs(resp.text)
        found.extend(jobs)
        if len(jobs) < PAGE_SIZE:
            break                                               # дальше страниц нет
        pause(cfg)
    return found


def parse_description(html):
    soup = BeautifulSoup(html, "html.parser")
    el = soup.select_one(
        "div.show-more-less-html__markup, div.description__text, section.description"
    )
    return el.get_text(" ", strip=True) if el else None


def fetch_description(session, job_id):
    """Текст описания вакансии или None, если получить не удалось."""
    try:
        resp = session.get(DETAIL_URL.format(job_id), headers=HEADERS, timeout=20)
    except requests.RequestException as e:
        print(f"  ! сеть (описание {job_id}): {e}")
        return None
    if resp.status_code == 429:
        raise RateLimited()
    if resp.status_code != 200:
        print(f"  ! описание {job_id}: ответ {resp.status_code}")
        return None
    return parse_description(resp.text)


# ---------- Telegram ----------

def send_telegram(cfg, text):
    token, chat = cfg.get("telegram_bot_token"), cfg.get("telegram_chat_id")
    if not token or not chat:
        print(text + "\n")
        return
    try:
        r = requests.post(
            f"https://api.telegram.org/bot{token}/sendMessage",
            data={"chat_id": chat, "text": text, "parse_mode": "HTML",
                  "disable_web_page_preview": "true"},
            timeout=20,
        )
        if r.status_code != 200:
            print(f"  ! Telegram: {r.status_code} {r.text[:200]}")
    except requests.RequestException as e:
        print(f"  ! Telegram: {e}")
    time.sleep(1)


def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def format_job(job, query):
    lines = [f"<b>{esc(job['title'])}</b>", esc(job["company"])]
    meta = " · ".join(x for x in (job["location"], job["posted"]) if x)
    if meta:
        lines.append(meta)
    lines.append(f"🔎 {esc(query)}")
    if job.get("matched_in_description"):
        lines.append(f"📄 в описании: {esc(job['matched_in_description'])}")
    lines.append(job["link"])
    return "\n".join(lines)


# ---------- main ----------

def main():
    cfg = load_json(CONFIG_PATH, None)
    if cfg is None:
        sys.exit(f"Не найден или повреждён {CONFIG_PATH}")

    exclude = cfg.get("exclude_words", [])
    require = cfg.get("require_any_words", [])
    check_desc = cfg.get("check_description", True)
    max_checks = int(cfg.get("max_description_checks", 60))

    seen = load_json(SEEN_PATH, [])
    seen_set = set(seen)
    first_run = not seen
    session = requests.Session()

    def mark_seen(job_id):
        seen_set.add(job_id)
        seen.append(job_id)

    results = []          # (job, query) — пойдут в Telegram
    to_check = []         # (job, query) — нужно проверить описание
    this_run = set()
    dropped = 0

    # 1. Поиск по всем запросам
    for query in cfg.get("searches", []):
        print(f"Ищу: {query}")
        for job in fetch_search(session, query, cfg):
            if job["id"] in seen_set or job["id"] in this_run:
                continue
            this_run.add(job["id"])
            head = f"{job['title']} {job['company']}"

            if find_word(exclude, head):                  # стоп-слово в названии/компании
                mark_seen(job["id"])
                dropped += 1
            elif not require or find_word(require, head): # крипто-слово уже в названии/компании
                mark_seen(job["id"])
                results.append((job, query))
            elif check_desc:                              # проверим текст вакансии
                to_check.append((job, query))
            else:
                mark_seen(job["id"])
                dropped += 1
        pause(cfg)

    # 2. Проверка описаний
    postponed = 0
    if to_check:
        print(f"Проверяю описания: {min(len(to_check), max_checks)} из {len(to_check)}")
    for i, (job, query) in enumerate(to_check):
        if i >= max_checks:
            postponed = len(to_check) - i                 # проверим в следующий запуск
            break
        try:
            desc = fetch_description(session, job["id"])
        except RateLimited:
            postponed = len(to_check) - i
            print("  ! LinkedIn просит притормозить (429), остальное проверю в следующий раз")
            break
        if desc is None:
            postponed += 1                                # не удалось — попробуем в следующий раз
        else:
            mark_seen(job["id"])
            word = find_word(require, desc)
            if word:
                job["matched_in_description"] = word
                results.append((job, query))
            else:
                dropped += 1
        pause(cfg)

    save_json(SEEN_PATH, seen[-SEEN_LIMIT:])

    print(f"Подходящих новых вакансий: {len(results)} | отсеяно: {dropped}"
          + (f" | отложено до следующего запуска: {postponed}" if postponed else ""))
    if not results:
        return
    if first_run:
        send_telegram(cfg, f"Мониторинг LinkedIn запущен. Сейчас найдено {len(results)} вакансий:")
    for job, query in results:
        send_telegram(cfg, format_job(job, query))


if __name__ == "__main__":
    main()
