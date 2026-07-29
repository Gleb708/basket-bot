import subprocess
import sys
import os
import warnings
warnings.filterwarnings('ignore')

# ---------- АВТОУСТАНОВКА БИБЛИОТЕК ----------
required_packages = ['pandas', 'openpyxl', 'numpy', 'pytz', 'apscheduler', 'python-telegram-bot', 'scipy']
for pkg in required_packages:
    try:
        __import__(pkg.replace('-', '_'))
    except ImportError:
        print(f"⚠️ Устанавливаю {pkg}...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', pkg])

# Теперь импорты
import pandas as pd
import numpy as np
from scipy.stats import t as t_dist
from scipy.stats import norm
from datetime import datetime, timedelta
import logging
import pytz
from telegram import Update
from telegram.ext import Application, CommandHandler, ContextTypes
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger

# ---------- КОНФИГУРАЦИЯ ----------
TOKEN = "8814276089:AAFT6vuV1wsU2W-Wpv7Z4QYYYVNExaVM1I8"
CHAT_IDS = [771729237, 749809260]
BASKEt_FILE = "Basket_3.xlsx"
FORECAST_FILE = "forecast_bayesian_clean.xlsx"
TIMEZONE = pytz.timezone("Europe/Moscow")
S0 = 13.5
k = 3
Z_75 = norm.ppf(0.75)
MIN_GAMES_FOR_STABLE = 3
CV_THRESHOLD = 10.0

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ---------- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ РАСЧЁТА ----------
def round_up_to_half(x):
    return np.ceil(x * 2) / 2

def round_down_to_half(x):
    return np.floor(x * 2) / 2

def winsorize_series(y, n_std=2.0):
    n = len(y)
    if n < 4:
        return y, False
    y = np.array(y)
    med = np.median(y)
    mad = np.median(np.abs(y - med))
    if mad == 0:
        return y, False
    robust_std = 1.4826 * mad
    lower = med - n_std * robust_std
    upper = med + n_std * robust_std
    y_w = np.where(y > upper, upper, np.where(y < lower, lower, y))
    changed = not np.array_equal(y, y_w)
    return y_w.tolist(), changed

def bayesian_forecast(y):
    n = len(y)
    if n == 0:
        return None
    mu = np.mean(y)
    if n == 1:
        sigma2 = S0 ** 2
        df = k
    else:
        s2 = np.var(y, ddof=1)
        sigma2 = (k * S0**2 + (n - 1) * s2) / (k + n - 1)
        df = k + n - 1
    se = np.sqrt(sigma2 * (1 + 1 / n))
    t_val = t_dist.ppf(0.85, df)
    L70 = mu - t_val * se
    U70 = mu + t_val * se
    tm_min = round_up_to_half(mu + Z_75 * se)
    tb_max = round_down_to_half(mu - Z_75 * se)
    return mu, se, L70, U70, tm_min, tb_max

# ---------- ГЕНЕРАЦИЯ ПРОГНОЗОВ ИЗ Basket_3.xlsx (с фильтром по типу дня) ----------
def generate_forecasts():
    if not os.path.exists(BASKEt_FILE):
        logger.error(f"Файл {BASKEt_FILE} не найден.")
        return False
    try:
        df = pd.read_excel(BASKEt_FILE, sheet_name='Лист1', engine='openpyxl')
        df = df.dropna(subset=['сумма очков'])
        df['datetime'] = pd.to_datetime(df['дата'].astype(str) + ' ' + df['время'].astype(str), errors='coerce')
        df = df.dropna(subset=['datetime'])
        if 'Статус' in df.columns:
            df = df[df['Статус'] == 'Завершена'].copy()

        # --- ОПРЕДЕЛЯЕМ ТИП ДНЯ (БУДНИЙ / ВЫХОДНОЙ) И ФИЛЬТРУЕМ ИСТОРИЮ ---
        today = datetime.now(TIMEZONE).date()
        is_weekend = today.weekday() >= 5  # 5=суббота, 6=воскресенье
        day_type = 'weekend' if is_weekend else 'weekday'
        logger.info(f"Сегодня {'выходной' if is_weekend else 'будний'} день. Использую историю для {day_type}.")

        # Добавляем колонку с типом дня для каждого матча
        df['day_type'] = df['datetime'].dt.weekday.apply(lambda x: 'weekend' if x >= 5 else 'weekday')
        # Оставляем только матчи, совпадающие с сегодняшним типом дня
        df = df[df['day_type'] == day_type].copy()
        if df.empty:
            logger.warning("Нет исторических матчей для данного типа дня. Прогнозы не будут сгенерированы.")
            return False

        # Формируем ключи пар
        teams1 = df['Команда 1'].astype(str).str.strip()
        teams2 = df['Команда 2'].astype(str).str.strip()
        pairs = []
        for a, b in zip(teams1, teams2):
            if a < b:
                pairs.append(f"{a} – {b}")
            else:
                pairs.append(f"{b} – {a}")
        df['pair_key'] = pairs
        df = df.sort_values(['pair_key', 'datetime']).reset_index(drop=True)

        stats_rows = []
        forecast_rows = []

        for key, group in df.groupby('pair_key'):
            y_all = group['сумма очков'].tolist()
            n_all = len(y_all)
            mean_all = np.mean(y_all)
            std_all = np.std(y_all, ddof=1) if n_all > 1 else np.nan
            min_all = np.min(y_all)
            max_all = np.max(y_all)
            range_all = max_all - min_all
            cv_all = (std_all / mean_all) * 100 if (n_all > 1 and mean_all != 0) else np.nan
            last_date = group['datetime'].max().strftime('%Y-%m-%d')
            if n_all < MIN_GAMES_FOR_STABLE:
                stability = "Нестабильная"
            else:
                stability = "Нестабильная" if (cv_all is not None and cv_all > CV_THRESHOLD) else "Стабильная"

            y_clean, was_changed = winsorize_series(y_all, n_std=2.0)
            n_clean = len(y_clean)
            forecast = bayesian_forecast(y_clean)
            if forecast is not None:
                mu, se, L70, U70, tm_min, tb_max = forecast
                forecast_rows.append({
                    'Матчап': key,
                    'Игр (всего)': n_all,
                    'Игр (после очистки)': n_clean,
                    'Были выбросы': 'Да' if was_changed else 'Нет',
                    'Последняя игра': last_date,
                    'Прогноз μ': round(mu, 1),
                    'SE прогноза': round(se, 2),
                    '70% интервал': f"{round(L70,1)} – {round(U70,1)}",
                    'ТМ (мин. линия ≥75%)': tm_min,
                    'ТБ (макс. линия ≥75%)': tb_max,
                    'Стабильность': stability
                })
            stats_rows.append({
                'Матчап': key,
                'Игр': n_all,
                'Последняя игра': last_date,
                'Средний тотал': round(mean_all, 1),
                'Стд откл': round(std_all, 1) if not np.isnan(std_all) else '-',
                'Мин': min_all,
                'Макс': max_all,
                'Размах': range_all,
                'CV (%)': round(cv_all, 1) if not np.isnan(cv_all) else '-',
                'Стабильность': stability
            })

        stats_df = pd.DataFrame(stats_rows).sort_values('Последняя игра', ascending=False)
        forecast_df = pd.DataFrame(forecast_rows).sort_values('Последняя игра', ascending=False)

        with pd.ExcelWriter(FORECAST_FILE, engine='openpyxl') as writer:
            forecast_df.to_excel(writer, sheet_name='Прогноз', index=False)
            stats_df.to_excel(writer, sheet_name='Статистика', index=False)

        logger.info(f"Прогнозы успешно сгенерированы и сохранены в {FORECAST_FILE}")
        return True
    except Exception as e:
        logger.error(f"Ошибка генерации прогнозов: {e}")
        return False

# ---------- ГЕНЕРАЦИЯ РАСПИСАНИЯ НА СЕГОДНЯ (с учётом типа дня) ----------
def get_typical_time(group):
    return group['datetime'].max().time()

def generate_schedule_today():
    if not os.path.exists(BASKEt_FILE):
        return pd.DataFrame()
    df = pd.read_excel(BASKEt_FILE, sheet_name='Лист1')
    df['datetime'] = pd.to_datetime(df['дата'].astype(str) + ' ' + df['время'].astype(str), errors='coerce')
    df = df.dropna(subset=['datetime'])
    # Фильтруем только те дни, которые совпадают с сегодняшним типом
    today = datetime.now(TIMEZONE).date()
    is_weekend = today.weekday() >= 5
    day_type = 'weekend' if is_weekend else 'weekday'
    df['day_type'] = df['datetime'].dt.weekday.apply(lambda x: 'weekend' if x >= 5 else 'weekday')
    df = df[df['day_type'] == day_type].copy()
    if df.empty:
        return pd.DataFrame()
    teams1 = df['Команда 1'].astype(str).str.strip()
    teams2 = df['Команда 2'].astype(str).str.strip()
    pairs = []
    for a, b in zip(teams1, teams2):
        if a < b:
            pairs.append(f"{a} – {b}")
        else:
            pairs.append(f"{b} – {a}")
    df['матчап'] = pairs

    schedule = []
    today = datetime.now(TIMEZONE).date()
    for matchup, group in df.groupby('матчап'):
        last_time = get_typical_time(group)
        dt = datetime.combine(today, last_time)
        dt = TIMEZONE.localize(dt)
        schedule.append({'матчап': matchup, 'datetime': dt})
    return pd.DataFrame(schedule)

# ---------- ЗАГРУЗКА РАСПИСАНИЯ ----------
def load_schedule():
    if not os.path.exists(BASKEt_FILE):
        return generate_schedule_today()
    try:
        df = pd.read_excel(BASKEt_FILE, sheet_name='Лист1')
        df['datetime'] = pd.to_datetime(df['дата'].astype(str) + ' ' + df['время'].astype(str), errors='coerce')
        df = df.dropna(subset=['datetime'])
        today = datetime.now(TIMEZONE).date()
        df_today = df[df['datetime'].dt.date == today].copy()
        if not df_today.empty:
            teams1 = df_today['Команда 1'].astype(str).str.strip()
            teams2 = df_today['Команда 2'].astype(str).str.strip()
            pairs = []
            for a, b in zip(teams1, teams2):
                if a < b:
                    pairs.append(f"{a} – {b}")
                else:
                    pairs.append(f"{b} – {a}")
            df_today['матчап'] = pairs
            return df_today[['матчап', 'datetime']]
        else:
            logger.info("В файле нет матчей на сегодня, генерирую расписание из истории (с учётом типа дня).")
            return generate_schedule_today()
    except Exception as e:
        logger.error(f"Ошибка загрузки расписания: {e}")
        return generate_schedule_today()

# ---------- ЗАГРУЗКА ПРОГНОЗОВ ----------
def load_forecasts():
    if not os.path.exists(FORECAST_FILE):
        return pd.DataFrame()
    return pd.read_excel(FORECAST_FILE, sheet_name='Прогноз')

# ---------- НОВОЕ ФОРМИРОВАНИЕ СООБЩЕНИЯ (по шаблону) ----------
def build_message(forecast_df, schedule_df):
    if forecast_df.empty or schedule_df.empty:
        return "Нет данных для прогнозов."

    now = datetime.now(TIMEZONE)
    hour_start = now.replace(minute=0, second=0, microsecond=0)
    hour_end = hour_start + timedelta(hours=1)
    mask = (schedule_df['datetime'] >= hour_start) & (schedule_df['datetime'] < hour_end)
    upcoming = schedule_df[mask]
    if upcoming.empty:
        return f"В этом часе ({hour_start.strftime('%H:%M')}) матчей нет."

    merged = pd.merge(upcoming, forecast_df, left_on='матчап', right_on='Матчап', how='inner')
    if merged.empty:
        return "Нет прогнозов для предстоящих матчей."

    # Определяем день недели для заголовка
    day_name = now.strftime('%A')  # Monday, Tuesday, ...
    # Переводим на русский (можно расширить)
    ru_days = {
        'Monday': 'Понедельник', 'Tuesday': 'Вторник', 'Wednesday': 'Среда',
        'Thursday': 'Четверг', 'Friday': 'Пятница', 'Saturday': 'Суббота', 'Sunday': 'Воскресенье'
    }
    day_ru = ru_days.get(day_name, day_name)

    lines = [f"📅 {day_ru}, {hour_start.strftime('%H:%M')}\n"]
    for _, row in merged.iterrows():
        match = row['матчап']
        mu = row['Прогноз μ']
        se = row['SE прогноза']
        interval = row['70% интервал']
        tm = row['ТМ (мин. линия ≥75%)']
        tb = row['ТБ (макс. линия ≥75%)']
        stable = row['Стабильность']

        # Иконка стабильности
        stable_icon = "✅" if stable == "Стабильная" else "⚠️"

        lines.append(f"📊 {match}")
        lines.append(f"🔹 70% интервал: {interval}")
        lines.append(f"🔹 ТМ (мин. линия ≥75%): {tm}")
        lines.append(f"🔹 ТБ (макс. линия ≥75%): {tb}")
        lines.append(f"🔹 Стабильность: {stable_icon} {stable}")
        lines.append("")  # пустая строка между матчами

    return "\n".join(lines)

# ---------- ЗАДАЧА ДЛЯ ОТПРАВКИ ----------
async def send_forecast(context: ContextTypes.DEFAULT_TYPE):
    logger.info("Запуск отправки прогнозов...")
    if not generate_forecasts():
        for chat_id in CHAT_IDS:
            await context.bot.send_message(chat_id=chat_id, text="Не удалось сгенерировать прогнозы.")
        return
    forecast_df = load_forecasts()
    schedule_df = load_schedule()
    if forecast_df.empty or schedule_df.empty:
        for chat_id in CHAT_IDS:
            await context.bot.send_message(chat_id=chat_id, text="Нет данных для прогнозов (возможно, нет матчей на сегодня).")
        return
    msg = build_message(forecast_df, schedule_df)
    for chat_id in CHAT_IDS:
        await context.bot.send_message(chat_id=chat_id, text=msg)
    logger.info("Прогнозы отправлены всем получателям.")

# ---------- КОМАНДЫ ----------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Бот запущен! Прогнозы приходят за 5 минут до каждого часа.")

async def now(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not generate_forecasts():
        await update.message.reply_text("Не удалось сгенерировать прогнозы.")
        return
    forecast_df = load_forecasts()
    schedule_df = load_schedule()
    msg = build_message(forecast_df, schedule_df)
    await update.message.reply_text(msg)

# ---------- ЗАПУСК ----------
def main():
    logger.info("Генерация прогнозов при старте...")
    generate_forecasts()

    application = Application.builder().token(TOKEN).build()
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("now", now))

    scheduler = BackgroundScheduler(timezone=TIMEZONE)
    scheduler.add_job(
        lambda: application.create_task(send_forecast(application.bot, None)),
        CronTrigger(minute=55, hour="*", timezone=TIMEZONE)
    )
    scheduler.start()
    logger.info("Планировщик запущен. Бот будет отправлять прогнозы в 55 минут каждого часа.")
    application.run_polling()

if __name__ == "__main__":
    main()
