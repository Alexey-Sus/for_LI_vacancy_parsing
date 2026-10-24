#!/usr/bin/env python3
"""
Мониторинг вакансий LinkedIn для русского языка (локализация, копирайтинг, MTPE, LQA и т.п.).
Отдельная версия скрипта: свои настройки, история и таблица.
Находит новые вакансии по поисковым запросам из config_russian_old.json,
отсеивает уже виденные и нерелевантные, присылает новые в Telegram.
Все найденные вакансии записываются в vacancies_russian.csv рядом со скриптом.

Запуск:  python linkedin_jobs_russian.py
Зависимости:  pip install requests beautifulsoup4
"""

import csv
import json
import random
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urlencode

import requests
from bs4 import BeautifulSoup

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config_russian_old.json"
SEEN_PATH = BASE_DIR / "seen_jobs_russian.json"
PENDING_PATH = BASE_DIR / "pending_jobs_russian.json"   # очередь непроверенных описаний
LOG_PATH = BASE_DIR / "vacancies_russian.csv"           # таблица всех найденных вакансий

SEARCH_URL = "https://www.linkedin.com/jobs-guest/jobs/api/seeMoreJobPostings/search"
DETAIL_URL = "https://www.linkedin.com/jobs-guest/jobs/api/jobPosting/{}"
PAGE_SIZE = 10          # гостевой поиск отдаёт около 10 вакансий за запрос
SEEN_LIMIT = 5000       # сколько ID вакансий помнить
MAX_ATTEMPTS = 3        # сколько раз пытаться открыть описание, потом сдаться

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


def pause_desc(cfg):
    lo, hi = cfg.get("description_delay_sec", [2, 4])
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


def format_job(job, query, tag=""):
    lines = [f"{esc(tag)} <b>{esc(job['title'])}</b>" if tag else f"<b>{esc(job['title'])}</b>", esc(job["company"])]
    meta = " · ".join(x for x in (job["location"], job["posted"]) if x)
    if meta:
        lines.append(meta)
    lines.append(f"🔎 {esc(query)}")
    if job.get("matched_in_description"):
        lines.append(f"📄 в описании: {esc(job['matched_in_description'])}")
    lines.append(job["link"])
    return "\n".join(lines)


# ---------- таблица вакансий ----------

LOG_FIELDS = ["Дата проверки", "Статус", "Причина", "Название", "Компания",
              "Локация", "Опубликовано", "Запрос", "Ссылка"]


def log_job(job, query, status, reason):
    """Дописывает строку в vacancies_russian.csv (открывается в Excel)."""
    new_file = not LOG_PATH.exists()
    try:
        with LOG_PATH.open("a", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter=";")
            if new_file:
                w.writerow(LOG_FIELDS)
            w.writerow([datetime.now().strftime("%Y-%m-%d %H:%M"), status, reason,
                        job.get("title", ""), job.get("company", ""), job.get("location", ""),
                        job.get("posted", ""), query, job.get("link", "")])
    except PermissionError:
        print("  ! vacancies_russian.csv открыт в другой программе (Excel?) — строка не записана", flush=True)


# ---------- main ----------

def main():
    cfg = load_json(CONFIG_PATH, None)
    if cfg is None:
        sys.exit(f"Не найден или повреждён {CONFIG_PATH}")

    exclude = cfg.get("exclude_words", [])
    require = cfg.get("require_any_words", [])
    roles = cfg.get("require_role_words", [])     # профессия — проверяется только в названии
    check_desc = cfg.get("check_description", True)
    max_checks = int(cfg.get("max_description_checks", 60))

    seen = load_json(SEEN_PATH, [])
    seen_set = set(seen)
    first_run = not seen
    session = requests.Session()

    def mark_seen(job_id):
        if job_id not in seen_set:
            seen_set.add(job_id)
            seen.append(job_id)

    # Очередь на проверку описаний: сначала отложенные с прошлых запусков
    queue = []            # элементы: {"job": {...}, "query": "...", "attempts": N}
    queued_ids = set()
    for item in load_json(PENDING_PATH, []):
        jid = item.get("job", {}).get("id")
        if jid and jid not in seen_set and jid not in queued_ids:
            queue.append(item)
            queued_ids.add(jid)
    from_pending = len(queue)

    results = []          # (job, query) — пойдут в Telegram
    this_run = set()
    dropped = 0

    # 1. Поиск по всем запросам
    for query in cfg.get("searches", []):
        print(f"Ищу: {query}", flush=True)
        for job in fetch_search(session, query, cfg):
            jid = job["id"]
            if jid in seen_set or jid in this_run or jid in queued_ids:
                continue
            this_run.add(jid)
            head = f"{job['title']} {job['company']}"

            stop = find_word(exclude, head)
            if stop:                                      # стоп-слово в названии/компании
                mark_seen(jid)
                dropped += 1
                log_job(job, query, "отсеяна", f"стоп-слово: {stop}")
                continue
            if roles and not find_word(roles, job["title"]):  # не контентная профессия
                mark_seen(jid)
                dropped += 1
                log_job(job, query, "отсеяна", "не та профессия (нет слов из require_role_words в названии)")
                continue
            word = find_word(require, head) if require else "фильтр не задан"
            if word:                                      # нужное слово уже в названии/компании
                mark_seen(jid)
                results.append((job, query))
                log_job(job, query, "отправлена", f"в названии/компании: {word}")
            elif check_desc:                              # проверим текст вакансии
                queue.append({"job": job, "query": query, "attempts": 0})
                queued_ids.add(jid)
            else:
                mark_seen(jid)
                dropped += 1
                log_job(job, query, "отсеяна", "нет нужных слов в названии")
        pause(cfg)

    sent = 0
    retry = []            # не открылись в этот раз — попробуем снова
    pos = 0               # сколько элементов очереди уже обработано

    def save():
        save_json(SEEN_PATH, seen[-SEEN_LIMIT:])
        save_json(PENDING_PATH, retry + queue[pos:])

    def notify(job, query):
        nonlocal sent, first_run
        if first_run:
            send_telegram(cfg, "Мониторинг LinkedIn (русский язык) запущен. Присылаю найденные вакансии:")
            first_run = False
        send_telegram(cfg, format_job(job, query, cfg.get("message_tag", "")))
        sent += 1

    # 2. Сразу отправляем то, что прошло по названию/компании, и сохраняем прогресс
    print(f"Прошли по названию: {len(results)} | на проверку описаний: {len(queue)}"
          + (f" (из них отложенных с прошлого раза: {from_pending})" if from_pending else ""),
          flush=True)
    for job, query in results:
        notify(job, query)
    save()

    # 3. Проверка описаний — с прогрессом; каждая находка отправляется сразу
    total = min(len(queue), max_checks)
    if queue:
        print(f"Проверяю описания ({total} из {len(queue)}), это займёт несколько минут...", flush=True)
    try:
        while pos < len(queue) and pos < max_checks:
            item = queue[pos]
            job, query = item["job"], item["query"]
            label = f"  [{pos + 1}/{total}] {job['title'][:50]} — {job['company'][:30]}"
            try:
                desc = fetch_description(session, job["id"])
            except RateLimited:
                print("  ! LinkedIn просит притормозить (429), остальное проверю в следующий раз", flush=True)
                break
            pos += 1
            if desc is None:
                item["attempts"] = item.get("attempts", 0) + 1
                if item["attempts"] >= MAX_ATTEMPTS:
                    mark_seen(job["id"])
                    dropped += 1
                    log_job(job, query, "отсеяна", f"описание не открылось {MAX_ATTEMPTS} раза")
                    print(f"{label} → не открылось {MAX_ATTEMPTS} раза, пропускаю", flush=True)
                else:
                    retry.append(item)
                    print(f"{label} → не открылось, попробую в следующий раз", flush=True)
            else:
                mark_seen(job["id"])
                word = find_word(require, desc)
                if word:
                    job["matched_in_description"] = word
                    print(f"{label} → ✅ {word}", flush=True)
                    log_job(job, query, "отправлена", f"в описании: {word}")
                    notify(job, query)
                else:
                    dropped += 1
                    log_job(job, query, "отсеяна", "нет нужных слов в описании")
                    print(f"{label} → мимо", flush=True)
            save()                                        # прогресс не теряется при прерывании
            if pos < total:
                pause_desc(cfg)
    except KeyboardInterrupt:
        print("\nОстановлено вручную. Прогресс сохранён.", flush=True)
    finally:
        save()

    postponed = len(retry) + len(queue) - pos
    print(f"Готово. Отправлено в Telegram: {sent} | отсеяно: {dropped}"
          + (f" | в очереди на следующий запуск: {postponed}" if postponed else ""), flush=True)


if __name__ == "__main__":
    main()
