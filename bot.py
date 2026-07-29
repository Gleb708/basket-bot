import subprocess
import sys
import os

# ---------- АВТОУСТАНОВКА БИБЛИОТЕК (если их нет) ----------
required_packages = ['pandas', 'openpyxl', 'numpy', 'pytz', 'apscheduler', 'python-telegram-bot']

def install_package(package):
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])

for pkg in required_packages:
    try:
        __import__(pkg.replace('-', '_'))  # для telegram-bot нужно заменить дефис
    except ImportError:
        print(f"⚠️ Устанавливаю {pkg}...")
        install_package(pkg)

# Теперь импортируем всё как обычно
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import logging
import pytz
from telegram import Update
from telegram.ext import Application, CommandHandler, ContextTypes
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger

# ---------- НАСТРОЙКИ ----------
TOKEN = "8814276089:AAFT6vuV1wsU2W-Wpv7Z4QYYYVNExaVM1I8"
CHAT_IDS = [771729237, 749809260]
EXCEL_FILE = "forecast_bayesian_clean.xlsx"
TIMEZONE = pytz.timezone("Europe/Moscow")

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ---------- ЗАГРУЗКА ДАННЫХ ----------
def load_forecasts():
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame()
    return pd.read_excel(EXCEL_FILE, sheet_name='Прогноз')

def load_schedule():
    basket_file = "Basket_3.xlsx"
    if not os.path.exists(basket_file):
        return pd.DataFrame()
    df = pd.read_excel(basket_file, sheet_name='Лист1')
    df['datetime'] = pd.to_datetime(df['дата'].astype(str) + ' ' + df['время'].astype(str), errors='coerce')
    df = df[df['datetime'].dt.date == datetime.now(TIMEZONE).date()]
    teams1 = df['Команда 1'].astype(str).str.strip()
    teams2 = df['Команда 2'].astype(str).str.strip()
    pairs = []
    for a, b in zip(teams1, teams2):
        pairs.append(f"{a} – {b}" if a < b else f"{b} – {a}")
    df['матчап'] = pairs
    return df[['матчап', 'datetime']]

# ---------- ФОРМИРОВАНИЕ СООБЩЕНИЯ ----------
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

    lines = [f"⏰ Прогнозы на матчи в {hour_start.strftime('%H:%M')}:\n"]
    for _, row in merged.iterrows():
        match = row['матчап']
        mu = row['Прогноз μ']
        se = row['SE прогноза']
        interval = row['70% интервал']
        tm = row['ТМ (мин. линия ≥75%)']
        tb = row['ТБ (макс. линия ≥75%)']
        stable = row['Стабильность']

        rec = ""
        if stable == "Стабильная" and pd.notna(tm) and pd.notna(tb):
            if tm - tb > 5:
                rec = "✅ ВАЛЮЙ! (Широкий коридор)"
            elif mu > (tm + tb)/2:
                rec = "📈 Склонность к ТБ"
            else:
                rec = "📉 Склонность к ТМ"
        else:
            rec = "⚠️ Нестабильная пара – осторожно."

        lines.append(f"🔹 {match}")
        lines.append(f"   μ = {mu:.1f}  SE = {se:.2f}")
        lines.append(f"   70% интервал: {interval}")
        lines.append(f"   ТМ≥75%: {tm}  |  ТБ≥75%: {tb}")
        lines.append(f"   {rec}\n")

    return "\n".join(lines)

# ---------- ЗАДАЧА ДЛЯ ОТПРАВКИ ВСЕМ ПОЛУЧАТЕЛЯМ ----------
async def send_forecast(context: ContextTypes.DEFAULT_TYPE):
    logger.info("Запуск отправки прогнозов...")
    forecast_df = load_forecasts()
    schedule_df = load_schedule()
    if forecast_df.empty or schedule_df.empty:
        for chat_id in CHAT_IDS:
            await context.bot.send_message(chat_id=chat_id, text="Не удалось загрузить данные.")
        return
    msg = build_message(forecast_df, schedule_df)
    for chat_id in CHAT_IDS:
        await context.bot.send_message(chat_id=chat_id, text=msg)
    logger.info("Прогнозы отправлены всем получателям.")

# ---------- КОМАНДЫ ----------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Бот запущен! Прогнозы приходят за 5 минут до каждого часа.")

async def now(update: Update, context: ContextTypes.DEFAULT_TYPE):
    forecast_df = load_forecasts()
    schedule_df = load_schedule()
    msg = build_message(forecast_df, schedule_df)
    await update.message.reply_text(msg)

# ---------- ЗАПУСК ----------
def main():
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
