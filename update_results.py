import requests
import re
import pandas as pd
from datetime import datetime
import os
import subprocess

# ---------- НАСТРОЙКИ ----------
URL = "https://2score.pro/ru/basketball/ipbl-pro-division-2496666/"
EXCEL_FILE = "Basket_3.xlsx"
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")

# ---------- ОТЛАДКА: показываем строки с матчами ----------
def debug_html(html):
    lines = html.split('\n')
    match_lines = [line for line in lines if 'Завершена' in line and 'href=' in line]
    print(f"🔍 Найдено строк с 'Завершена': {len(match_lines)}")
    if match_lines:
        print("📄 Первые 5 таких строк:")
        for i, line in enumerate(match_lines[:5]):
            print(f"   {i+1}: {line[:200]}...")  # первые 200 символов

# ---------- ФУНКЦИЯ ПУША ----------
def push_to_github():
    if not GITHUB_TOKEN:
        print("❌ GITHUB_TOKEN не задан.")
        return

    subprocess.run(["git", "config", "user.name", "basket-bot"], check=False)
    subprocess.run(["git", "config", "user.email", "bot@example.com"], check=False)

    result_add = subprocess.run(["git", "add", EXCEL_FILE], capture_output=True, text=True)
    if result_add.returncode != 0:
        print(f"❌ git add: {result_add.stderr}")
        return

    commit_msg = f"Автоматическое обновление Basket_3.xlsx ({datetime.now().strftime('%Y-%m-%d %H:%M')})"
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

# ---------- ПАРСИНГ ----------
def parse_match(line):
    if "Завершена" not in line:
        return None

    date_match = re.search(r'(\d{2}\.\d{2})\s+(\d{2}:\d{2})', line)
    if not date_match:
        return None
    date_str, time_str = date_match.groups()

    link_match = re.search(r'href="([^"]+)"', line)
    if not link_match:
        return None
    slug = link_match.group(1).split('/')[-2] if '/' in link_match.group(1) else ''
    if '-' in slug:
        parts = slug.split('-')
        if parts[-1].isdigit():
            parts = parts[:-1]
        mid = len(parts) // 2
        team1 = ' '.join(parts[:mid]).replace('-', ' ').title().strip()
        team2 = ' '.join(parts[mid:]).replace('-', ' ').title().strip()
    else:
        rest = re.sub(r'\d{2}\.\d{2}\s+\d{2}:\d{2}\s+', '', line)
        teams_part = rest.split("Завершена")[0].strip()
        words = teams_part.split()
        if len(words) >= 2:
            team1 = ' '.join(words[:-1])
            team2 = words[-1]
        else:
            team1 = teams_part
            team2 = ''

    numbers = re.findall(r'\b\d{1,3}\b', line)
    if len(numbers) < 10:
        return None
    scores = numbers[-10:]
    scores_int = [int(x) for x in scores]

    return {
        'дата': date_str,
        'время': time_str,
        'Команда 1': team1,
        'Команда 2': team2,
        'Счет 1': scores_int[0],
        'Q1_1': scores_int[1],
        'Q2_1': scores_int[2],
        'Q3_1': scores_int[3],
        'Q4_1': scores_int[4],
        'Счет 2': scores_int[5],
        'Q1_2': scores_int[6],
        'Q2_2': scores_int[7],
        'Q3_2': scores_int[8],
        'Q4_2': scores_int[9],
        'сумма очков': scores_int[0] + scores_int[5],
        'Статус': 'Завершена'
    }

def update_excel():
    print("📡 Загрузка страницы...")
    try:
        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(URL, headers=headers, timeout=15)
        response.encoding = 'utf-8'
        html = response.text
    except Exception as e:
        print(f"❌ Ошибка загрузки: {e}")
        return

    # Отладочный вывод
    debug_html(html)

    lines = html.split('\n')
    new_matches = []
    for line in lines:
        if 'Завершена' in line and 'href=' in line:
            match = parse_match(line)
            if match:
                new_matches.append(match)

    if not new_matches:
        print("⚠️ Новых матчей не найдено.")
        return

    print(f"✅ Найдено завершённых матчей: {len(new_matches)}")

    if os.path.exists(EXCEL_FILE):
        df_existing = pd.read_excel(EXCEL_FILE, sheet_name='Лист1', engine='openpyxl')
    else:
        df_existing = pd.DataFrame()

    df_new = pd.DataFrame(new_matches)

    if not df_existing.empty:
        df_existing['key'] = df_existing['дата'].astype(str) + '|' + df_existing['время'].astype(str) + '|' + df_existing['Команда 1'] + '|' + df_existing['Команда 2']
        df_new['key'] = df_new['дата'].astype(str) + '|' + df_new['время'].astype(str) + '|' + df_new['Команда 1'] + '|' + df_new['Команда 2']
        df_new = df_new[~df_new['key'].isin(df_existing['key'])]
        df_new = df_new.drop(columns=['key'])

    if df_new.empty:
        print("✅ Все матчи уже есть в файле.")
        return

    df_combined = pd.concat([df_existing, df_new], ignore_index=True)
    df_combined['datetime'] = pd.to_datetime(df_combined['дата'].astype(str) + ' ' + df_combined['время'].astype(str), errors='coerce')
    df_combined = df_combined.sort_values('datetime').drop(columns=['datetime'])

    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
        df_combined.to_excel(writer, sheet_name='Лист1', index=False)

    print(f"✅ Файл обновлён! Добавлено {len(df_new)} новых матчей. Всего записей: {len(df_combined)}")

    # Пуш
    push_to_github()

if __name__ == "__main__":
    update_excel()
