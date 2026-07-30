# -*- coding: utf-8 -*-
"""
Парсер результатов IPBL с 2score.pro (адаптирован под Basket_3.xlsx).
Использует рабочий код знакомого, но сохраняет данные в Excel.

Запуск: python update_results_v2.py
(или вызывается из бота по /push)
"""

import re
import sys
import datetime
import urllib.request
import urllib.error
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
import os
import pandas as pd
import subprocess

# ---------- НАСТРОЙКИ ----------
BASE = "https://2score.pro/ru/basketball"
EXCEL_FILE = "Basket_3.xlsx"
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")

MAX_WORKERS = 4          # уменьшим для экономии памяти (бесплатный Railway)
RETRIES = 3
TIMEOUT = 30
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"}

# Нас интересует только мужской PRO-дивизион
SOURCES = {
    "ipbl-pro-division-2496666": "by_team",   # PRO муж
}

# ---------- ТАБЛИЦЫ КОМАНД (как у знакомого) ----------
PRO_MEN = {
    'A': ['Novosibirsk', 'Barnaul', 'St. Petersburg', 'Sochi'],
    'B': ['Tyumen', 'Krasnodar', 'Samara', 'Kazan'],
    'C': ['Plavsk', 'Moscow', 'Voronezh', 'Kaliningrad'],
    'D': ['Volgograd', 'Rostov-on-Don', 'Nizhny Novgorod', 'Krasnoyarsk'],
    'F': ['Kemerovo', 'Irkutsk', 'Essentuki', 'Asbest'],
    'G': ['Tver', 'Gelendzhik', 'Astrakhan', 'Kachkanar'],
    'H': ['Belgorod', 'Maykop', 'Nalchik', 'Saratov'],
    'J': ['Khabarovsk', 'Chelyabinsk', 'Ulan-Ude', 'Chita'],
    'K': ['Kirov', 'Orel', 'Murmansk', 'Perm'],
    'L': ['Surgut', 'Adler', 'Kurgan', 'Yakutsk'],
    'M': ['Stavropol', 'Tikhoretsk', 'Engels', 'Sevastopol'],
    'N': ['Arkhangelsk', 'Vladikavkaz', 'Lipetsk', 'Tula'],
    'O': ['Makhachkala', 'Abakan', 'Petrozavodsk', 'Taganrog'],
    'U': ['Smolensk', 'Salavat', 'Serov', 'Ryazan'],
    'X': ['Vorkuta', 'Syktyvkar', 'Ukhta', 'Omsk'],
    'Z': ['Ufa', 'Anapa', 'Magadan', 'Revda'],
}
# Другие дивизионы нам не нужны, но оставим структуру
def _norm(name: str) -> str:
    s = name.lower().strip()
    s = s.replace('(жен)', '').replace('(ж)', '')
    s = re.sub(r'^ipbl\s+', '', s)
    s = re.sub(r'[^0-9a-zа-яё]', '', s)
    return s

def _is_women(name: str) -> bool:
    return ('(жен)' in name) or ('(ж)' in name)

TEAM_CLASS = {}
for _mp, _div, _w in [(PRO_MEN, 'PRO', False)]:
    for _cls, _teams in _mp.items():
        for _t in _teams:
            TEAM_CLASS[(_norm(_t), _w)] = (_div, _cls)

def lookup(name: str):
    return TEAM_CLASS.get((_norm(name), _is_women(name)))

# ---------- ПАРСИНГ (скопирован у знакомого) ----------
_GAME_SPLIT = re.compile(r'<div class="game-line ')
_DATA_ID = re.compile(r'data-id="(\d+)"')
_STATUS = re.compile(r'game-line__status[^>]*>(.*?)</div>', re.S)
_DT = re.compile(r'(\d{2})\.(\d{2})\s+(\d{2}:\d{2})')
_TEAM = re.compile(r'game-line__team[ "][^>]*>([^<]+)<')
_SUM = re.compile(r'game-line__score-sum[^>]*>\s*(\d+)')

def fetch(url: str):
    for attempt in range(RETRIES):
        try:
            req = urllib.request.Request(url, headers=HEADERS)
            with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
                return r.read().decode('utf-8', 'ignore')
        except urllib.error.HTTPError as e:
            if e.code == 404:
                return ""
            if attempt == RETRIES - 1:
                print(f"  ! HTTP {e.code}: {url}")
        except Exception as e:
            if attempt == RETRIES - 1:
                print(f"  ! {type(e).__name__}: {url}")
    return None

def parse_page(html: str, page_date: datetime.date):
    games = []
    for seg in _GAME_SPLIT.split(html)[1:]:
        mid = _DATA_ID.search(seg)
        if not mid:
            continue
        teams = [t.strip() for t in _TEAM.findall(seg) if t.strip()]
        sums = _SUM.findall(seg)
        st = _STATUS.search(seg)
        dt = _DT.search(st.group(1)) if st else None
        if len(teams) < 2 or len(sums) < 2 or not dt:
            continue
        dd, mm, hhmm = dt.groups()
        try:
            mdate = datetime.date(page_date.year, int(mm), int(dd))
            if abs((mdate - page_date).days) > 3:
                mdate = page_date
        except ValueError:
            mdate = page_date
        games.append({
            "match_id": mid.group(1),
            "date": mdate,
            "time": hhmm,
            "team1": teams[0],
            "team2": teams[1],
            "score1": int(sums[0]),
            "score2": int(sums[1]),
        })
    return games

def daterange(start: datetime.date, end: datetime.date):
    d = start
    while d <= end:
        yield d
        d += datetime.timedelta(days=1)

def collect(start: datetime.date, end: datetime.date):
    tasks = [(slug, d) for slug in SOURCES for d in daterange(start, end)]
    print(f"Запрашиваю {len(tasks)} страниц...")
    rows = {}
    done = 0
    def work(task):
        slug, d = task
        html = fetch(f"{BASE}/{slug}/{d.strftime('%Y-%m-%d')}/")
        if not html:
            return slug, []
        return slug, parse_page(html, d)
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        for slug, games in ex.map(work, tasks):
            done += 1
            if done % 20 == 0:
                print(f"  …{done}/{len(tasks)} страниц, матчей: {len(rows)}")
            mode = SOURCES[slug]
            for g in games:
                if mode == "by_team":
                    r1 = lookup(g["team1"])
                    r2 = lookup(g["team2"])
                    if r1 is None or r2 is None:
                        continue
                    if r1 != r2:
                        continue
                    division, cls = r1
                else:
                    continue  # нам нужен only by_team
                if g["date"] < start or g["date"] > end:
                    continue
                g2 = dict(g, division=division, cls=cls)
                rows[g["match_id"]] = g2
    return list(rows.values())

# ---------- ОБНОВЛЕНИЕ EXCEL ----------
def update_excel():
    # Парсим только за сегодня и вчера (на случай, если матчи закончились поздно)
    today = datetime.date.today()
    start = today - datetime.timedelta(days=2)
    end = today

    print(f"📡 Парсим 2score.pro за {start} … {end} ...")
    new_matches = collect(start, end)
    if not new_matches:
        print("⚠️ Новых матчей не найдено.")
        return

    # Преобразуем в DataFrame
    df_new = pd.DataFrame(new_matches)
    # Добавим недостающие колонки для совместимости с вашим форматом
    df_new['сумма очков'] = df_new['score1'] + df_new['score2']
    df_new['Статус'] = 'Завершена'
    # Переименуем колонки
    df_new.rename(columns={
        'date': 'дата',
        'time': 'время',
        'team1': 'Команда 1',
        'team2': 'Команда 2',
        'score1': 'Счет 1',
        'score2': 'Счет 2',
    }, inplace=True)
    # Добавим пустые колонки для четвертей (в 2score нет данных по четвертям в этом парсере)
    # Мы можем оставить их пустыми, но лучше заполнить из других источников.
    # Поскольку в вашем прогнозе используются четверти, это проблема.
    # В парсере знакомого нет четвертей, только общий счёт.
    # Это ограничение нужно учесть.

    # Загружаем существующий файл
    if os.path.exists(EXCEL_FILE):
        df_existing = pd.read_excel(EXCEL_FILE, sheet_name='Лист1', engine='openpyxl')
    else:
        df_existing = pd.DataFrame()

    # Проверяем дубликаты
    if not df_existing.empty:
        df_existing['key'] = df_existing['дата'].astype(str) + '|' + df_existing['время'].astype(str) + '|' + df_existing['Команда 1'] + '|' + df_existing['Команда 2']
        df_new['key'] = df_new['дата'].astype(str) + '|' + df_new['время'].astype(str) + '|' + df_new['Команда 1'] + '|' + df_new['Команда 2']
        df_new = df_new[~df_new['key'].isin(df_existing['key'])]
        df_new = df_new.drop(columns=['key'])

    if df_new.empty:
        print("✅ Все матчи уже есть в файле.")
        return

    # Приводим дату к строковому формату
    df_new['дата'] = df_new['дата'].dt.strftime('%Y-%m-%d')
    # Объединяем
    df_combined = pd.concat([df_existing, df_new], ignore_index=True)
    # Сортируем по дате и времени
    df_combined['datetime'] = pd.to_datetime(df_combined['дата'].astype(str) + ' ' + df_combined['время'].astype(str), errors='coerce')
    df_combined = df_combined.sort_values('datetime').drop(columns=['datetime'])

    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
        df_combined.to_excel(writer, sheet_name='Лист1', index=False)

    print(f"✅ Файл обновлён! Добавлено {len(df_new)} новых матчей. Всего записей: {len(df_combined)}")

    # Пуш на GitHub
    push_to_github()

def push_to_github():
    if not GITHUB_TOKEN:
        print("⚠️ GITHUB_TOKEN не задан, пуш не выполнен.")
        return
    subprocess.run(["git", "config", "user.name", "basket-bot"], check=False)
    subprocess.run(["git", "config", "user.email", "bot@example.com"], check=False)
    result_add = subprocess.run(["git", "add", EXCEL_FILE], capture_output=True, text=True)
    if result_add.returncode != 0:
        print(f"❌ git add: {result_add.stderr}")
        return
    commit_msg = f"Автоматическое обновление Basket_3.xlsx ({datetime.datetime.now().strftime('%Y-%m-%d %H:%M')})"
    result_commit = subprocess.run(["git", "commit", "-m", commit_msg], capture_output=True, text=True)
    if result_commit.returncode != 0:
        if "nothing to commit" in result_commit.stderr:
            print("ℹ️ Нет изменений для коммита.")
            return
        print(f"❌ git commit: {result_commit.stderr}")
        return
    remote_url = f"https://x-access-token:{GITHUB_TOKEN}@github.com/Gleb708/basket-bot.git"
    result_push = subprocess.run(["git", "push", remote_url, "main"], capture_output=True, text=True)
    if result_push.returncode != 0:
        print(f"❌ git push: {result_push.stderr}")
    else:
        print("✅ Пуш выполнен.")

if __name__ == "__main__":
    # Добавим datetime для коммита
    import datetime as dt
    # Переопределим datetime для коммита
    def get_now():
        return dt.datetime.now()
    # Исправление: в коде выше используется datetime.datetime.now, поэтому подставим
    update_excel()
