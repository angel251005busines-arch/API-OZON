#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Бот MAX для генерации финансовых отчётов Ozon
"""
import json
import asyncio
import logging
import os
import requests
from datetime import datetime, timedelta
from pathlib import Path
import re
import time

import aiofiles
import aiohttp
from dotenv import load_dotenv
from maxapi import Bot, Dispatcher
from maxapi.types import MessageCreated, Command

# Импорт функций из вашего скрипта
from sales_cost1 import (
    OzonFinanceAPI,
    load_cost_data,
    process_financial_data,
    save_to_excel,
)

load_dotenv()

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Инициализация
bot = Bot(token=os.getenv('MAX_BOT_TOKEN'))
dp = Dispatcher()

# Хранилище состояний пользователей
user_states: dict = {}

MAX_TOKEN = os.getenv('MAX_BOT_TOKEN')
BASE_URL = "https://platform-api.max.ru"


def send_message(chat_id: int, text: str, format_type: str = "html"):
    """Отправляет сообщение через API MAX"""
    url = f"{BASE_URL}/messages?chat_id={chat_id}"
    payload = {"text": text, "format": format_type}
    
    try:
        response = requests.post(
            url, 
            json=payload, 
            headers={'Authorization': MAX_TOKEN},
            timeout=30
        )
        if response.status_code == 200:
            logger.info(f"Сообщение отправлено в чат {chat_id}")
            return response.json()
        else:
            logger.error(f"Ошибка отправки: {response.status_code}")
            return None
    except Exception as e:
        logger.error(f"Ошибка при отправке: {e}")
        return None


def send_file_message(chat_id: int, text: str, file_token: str):
    """Отправляет сообщение с файлом"""
    url = f"{BASE_URL}/messages?chat_id={chat_id}"
    payload = {
        "text": text,
        "format": "html",
        "attachments": [{'type': 'file', 'payload': {'token': file_token}}]
    }
    
    try:
        response = requests.post(url, json=payload, headers={'Authorization': MAX_TOKEN}, timeout=30)
        if response.status_code == 200:
            logger.info(f"Файл отправлен в чат {chat_id}")
            return True
        else:
            logger.error(f"Ошибка отправки файла: {response.status_code} - {response.text}")
            return False
    except Exception as e:
        logger.error(f"Ошибка при отправке файла: {e}")
        return False


async def download_file_from_max(file_url: str, local_path: Path) -> bool:
    """Скачивает файл с серверов MAX по URL из payload"""
    try:
        async with aiohttp.ClientSession() as session:
            async with session.get(
                file_url,
                headers={'Authorization': MAX_TOKEN}
            ) as resp:
                if resp.status == 200:
                    content = await resp.read()
                    async with aiofiles.open(local_path, 'wb') as f:
                        await f.write(content)
                    logger.info(f"Файл скачан: {local_path}, размер: {len(content)} байт")
                    return True
                else:
                    logger.error(f"Ошибка загрузки файла: {resp.status}")
                    return False
    except Exception as e:
        logger.error(f"Ошибка при скачивании файла: {e}")
        return False


async def upload_file_to_max(file_path: Path) -> str | None:
    """Загружает файл на сервера MAX и возвращает токен"""
    try:
        # Шаг 1: Получаем URL для загрузки
        async with aiohttp.ClientSession() as session:
            async with session.post(
                f'{BASE_URL}/uploads?type=file',
                headers={'Authorization': MAX_TOKEN}
            ) as resp:
                if resp.status != 200:
                    logger.error(f"Ошибка получения URL: {resp.status}")
                    return None
                
                upload_data = await resp.json()
                upload_url = upload_data.get('url')
                
                if not upload_url:
                    logger.error("Не удалось получить URL для загрузки файла")
                    return None
            
            # Шаг 2: Загружаем файл по полученному URL
            async with aiofiles.open(file_path, 'rb') as f:
                file_content = await f.read()
            
            # Используем multipart/form-data как в документации
            form = aiohttp.FormData()
            form.add_field('data', file_content, filename=file_path.name)
            
            async with session.post(upload_url, data=form) as upload_resp:
                if upload_resp.status != 200:
                    logger.error(f"Ошибка загрузки файла: {upload_resp.status}")
                    return None
                
                result = await upload_resp.json()
                file_token = result.get('token')
                
                if file_token:
                    logger.info(f"Файл загружен, токен: {file_token}")
                    # Ждём обработки файла сервером (согласно документации)
                    await asyncio.sleep(3)
                    return file_token
                else:
                    logger.error("Не удалось получить токен файла")
                    return None
    except Exception as e:
        logger.error(f"Ошибка при загрузке файла: {e}")
        return None


def parse_date(text: str) -> datetime | None:
    """Парсит дату из текста"""
    patterns = [
        r'(\d{4})-(\d{1,2})-(\d{1,2})',
        r'(\d{1,2})\.(\d{1,2})\.(\d{4})',
        r'(\d{1,2})/(\d{1,2})/(\d{4})',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            groups = match.groups()
            if len(groups[0]) == 4:
                year, month, day = int(groups[0]), int(groups[1]), int(groups[2])
            else:
                day, month, year = int(groups[0]), int(groups[1]), int(groups[2])
            
            try:
                return datetime(year, month, day)
            except:
                continue
    return None


async def generate_report(chat_id: int, date_from: datetime, date_to: datetime, cost_file_path: str = None):
    """Генерирует отчёт за указанный период"""
    send_message(chat_id, f"⏳ Генерация отчёта...\n\nПериод: {date_from.date()} - {date_to.date()}\nЗагружаю данные с Ozon...")
    
    try:
        api = OzonFinanceAPI()
        
        logger.info(f"Получение типов начислений...")
        accrual_types = api.get_accrual_types()
        
        logger.info(f"Получение отправлений за период {date_from.date()} - {date_to.date()}")
        posting_numbers = api.get_posting_numbers_by_period(date_from, date_to)
        
        if not posting_numbers:
            send_message(chat_id, f"⚠️ Нет отправлений за период {date_from.date()} - {date_to.date()}")
            return
        
        logger.info(f"Получение финансовых данных по {len(posting_numbers)} отправлениям...")
        posting_accruals = api.get_accrual_postings(posting_numbers)
        
        cost_df, sku_col, cost_col, category_col, name_col = None, None, None, None, None
        if cost_file_path and os.path.exists(cost_file_path):
            cost_df, sku_col, cost_col, category_col, name_col = load_cost_data(cost_file_path)
            logger.info(f"Загружен файл себестоимости: {cost_file_path}")
        
        logger.info("Обработка финансовых данных...")
        df = process_financial_data(
            posting_accruals, accrual_types,
            cost_df, sku_col, cost_col, category_col, name_col
        )
        
        if df.empty:
            send_message(chat_id, "⚠️ Нет данных для формирования отчёта")
            return
        
        temp_dir = Path(os.getenv('TEMP_FOLDER', './temp'))
        temp_dir.mkdir(exist_ok=True)
        output_path = temp_dir / f"ozon_report_{chat_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        logger.info(f"Сохранение отчёта в {output_path}")
        save_to_excel(df, str(output_path))
        
        # Загружаем файл на сервер MAX
        file_token = await upload_file_to_max(output_path)
        
        if not file_token:
            raise Exception("Не удалось загрузить файл на сервер MAX")
        
        # ========== ИЗВЛЕКАЕМ ЧИСТУЮ ПРИБЫЛЬ ИЗ ФАЙЛА ==========
        total_profit = 0
        
        try:
            import pandas as pd
            
            # Читаем лист "Общая статистика"
            df_stats = pd.read_excel(output_path, sheet_name="Общая статистика", header=None)
            
            # Ищем строку с чистой прибылью
            for i in range(len(df_stats)):
                cell_value = str(df_stats.iloc[i, 0]).upper()
                if 'ЧИСТАЯ ПРИБЫЛЬ (после всех расходов)' in cell_value:
                    profit_cell = df_stats.iloc[i, 1]
                    logger.info(f"Найдена чистая прибыль в строке {i}: {profit_cell}")
                    
                    if isinstance(profit_cell, (int, float)):
                        total_profit = float(profit_cell)
                    else:
                        profit_str = str(profit_cell)
                        cleaned = profit_str.replace(' ', '').replace('₽', '').replace('руб', '').strip()
                        cleaned = cleaned.replace(',', '.')
                        total_profit = float(cleaned)
                    break
            
            if total_profit == 0:
                # Если не нашли, берём 8-ю строку (индекс 7)
                profit_cell = df_stats.iloc[7, 1]
                logger.info(f"Берём значение по индексу 7: {profit_cell}")
                
                if isinstance(profit_cell, (int, float)):
                    total_profit = float(profit_cell)
                else:
                    profit_str = str(profit_cell)
                    cleaned = profit_str.replace(' ', '').replace('₽', '').replace('руб', '').strip()
                    cleaned = cleaned.replace(',', '.')
                    total_profit = float(cleaned)
            
            logger.info(f"Распарсена чистая прибыль: {total_profit}")
            
        except Exception as e:
            logger.error(f"Ошибка при чтении чистой прибыли: {e}")
            total_profit = "внутри файла"  # Значение по умолчанию из вашего файла
        
        # Форматируем с пробелами как разделителями тысяч
        if total_profit > 0:
            profit_text = f"{total_profit:,.2f} ₽".replace(",", " ")
        else:
            profit_text = "не удалось рассчитать"
        
        logger.info(f"Итоговая чистая прибыль для отправки: {profit_text}")
        # ========================================================
        
        # Отправляем файл с сообщением
        success = send_file_message(
            chat_id,
            f"✅ <b>Отчёт готов!</b>\n\n"
            f"📊 Период: {date_from.date()} — {date_to.date()}\n"
            f"📄 Записей в отчёте: {len(df)}\n"
            f"💰 <b>Чистая прибыль: {profit_text}</b>\n\n"
            f"💾 Файл Excel с детальной аналитикой прикреплён ниже",
            file_token
        )
        
        if not success:
            await asyncio.sleep(3)
            send_file_message(chat_id, f"✅ Отчёт готов!\n\nПериод: {date_from.date()} — {date_to.date()}\nЗаписей: {len(df)}\nЧистая прибыль: {profit_text}", file_token)
        
        for f in [output_path, cost_file_path]:
            if f and os.path.exists(f):
                try:
                    os.remove(f)
                    logger.info(f"Удалён временный файл: {f}")
                except:
                    pass
        
    except Exception as e:
        logger.error(f"Ошибка при генерации отчёта: {e}", exc_info=True)
        send_message(chat_id, f"❌ Ошибка: {str(e)[:200]}")

# ==================== ОБРАБОТЧИКИ СОБЫТИЙ ====================

@dp.message_created(Command('start'))
async def on_start(event: MessageCreated):
    """Обработка команды /start"""
    chat_id = event.message.recipient.chat_id
    
    help_text = """🤖 <b>Ozon Finance Bot</b>

Я помогу получить финансовый отчёт по вашим продажам на Ozon.

<b>📊 Как получить отчёт:</b>

1. <b>Выберите период</b> командой:
   • /7 - последние 7 дней
   • /30 - последние 30 дней
   • /month - текущий месяц
   • /prev - прошлый месяц
   • /custom - свой период

2. <b>Для своего периода:</b>
   Напишите: с 2024-01-01 по 2024-01-31

3. <b>Загрузите файл</b> с себестоимостью (xlsx/csv)

<b>Формат файла себестоимости:</b>
• Колонка 'Артикул Озон' — SKU товара
• Колонка 'Трансферная цена комплект' — себестоимость

<b>Команды:</b>
• /help - Показать справку
• /cancel - Отменить операцию"""
    
    send_message(chat_id, help_text)


@dp.message_created(Command('help'))
async def on_help(event: MessageCreated):
    """Обработка команды /help"""
    chat_id = event.message.recipient.chat_id
    
    help_text = """📚 <b>Справка по боту</b>

<b>Быстрые команды:</b>
• /7 - Последние 7 дней
• /30 - Последние 30 дней
• /month - Текущий месяц
• /prev - Прошлый месяц

<b>Свой период:</b>
• /custom - затем введите: с 2024-01-01 по 2024-01-31

<b>Форматы дат:</b>
• 2024-01-15
• 15.01.2024
• 15/01/2024

<b>После выбора периода:</b>
Загрузите файл с себестоимостью

<b>Команды:</b>
• /start - Главное меню
• /cancel - Отменить операцию"""
    
    send_message(chat_id, help_text)


@dp.message_created(Command('cancel'))
async def on_cancel(event: MessageCreated):
    """Отмена текущей операции"""
    chat_id = event.message.recipient.chat_id
    user_states.pop(chat_id, None)
    send_message(chat_id, "✅ Операция отменена. Используйте /start для начала.")


@dp.message_created(Command('7'))
async def report_7(event: MessageCreated):
    """Отчёт за 7 дней"""
    chat_id = event.message.recipient.chat_id
    date_to = datetime.now()
    date_from = date_to - timedelta(days=7)
    user_states[chat_id] = {'date_from': date_from, 'date_to': date_to, 'awaiting_file': True}
    send_message(chat_id, f"✅ Выбран период: последние 7 дней ({date_from.date()} - {date_to.date()})\n\n📎 Теперь загрузите файл с себестоимостью (xlsx/csv)")


@dp.message_created(Command('30'))
async def report_30(event: MessageCreated):
    """Отчёт за 30 дней"""
    chat_id = event.message.recipient.chat_id
    date_to = datetime.now()
    date_from = date_to - timedelta(days=30)
    user_states[chat_id] = {'date_from': date_from, 'date_to': date_to, 'awaiting_file': True}
    send_message(chat_id, f"✅ Выбран период: последние 30 дней ({date_from.date()} - {date_to.date()})\n\n📎 Теперь загрузите файл с себестоимостью (xlsx/csv)")


@dp.message_created(Command('month'))
async def report_month(event: MessageCreated):
    """Отчёт за текущий месяц"""
    chat_id = event.message.recipient.chat_id
    now = datetime.now()
    date_from = datetime(now.year, now.month, 1)
    date_to = now
    user_states[chat_id] = {'date_from': date_from, 'date_to': date_to, 'awaiting_file': True}
    send_message(chat_id, f"✅ Выбран период: текущий месяц ({date_from.date()} - {date_to.date()})\n\n📎 Теперь загрузите файл с себестоимостью (xlsx/csv)")


@dp.message_created(Command('prev'))
async def report_prev(event: MessageCreated):
    """Отчёт за прошлый месяц"""
    chat_id = event.message.recipient.chat_id
    now = datetime.now()
    if now.month == 1:
        date_from = datetime(now.year - 1, 12, 1)
    else:
        date_from = datetime(now.year, now.month - 1, 1)
    date_to = datetime(now.year, now.month, 1) - timedelta(days=1)
    user_states[chat_id] = {'date_from': date_from, 'date_to': date_to, 'awaiting_file': True}
    send_message(chat_id, f"✅ Выбран период: прошлый месяц ({date_from.date()} - {date_to.date()})\n\n📎 Теперь загрузите файл с себестоимостью (xlsx/csv)")


@dp.message_created(Command('custom'))
async def report_custom(event: MessageCreated):
    """Ожидание ввода дат"""
    chat_id = event.message.recipient.chat_id
    user_states[chat_id] = {'awaiting_dates': True}
    send_message(chat_id, "📅 Введите период в формате:\n\nс 2024-01-01 по 2024-01-31\n\nПоддерживаемые форматы дат:\n• 2024-01-15\n• 15.01.2024\n• 15/01/2024")


@dp.message_created()
async def handle_messages(event: MessageCreated):
    """Обработка обычных сообщений"""
    chat_id = event.message.recipient.chat_id
    state = user_states.get(chat_id, {})
    
    # Получаем текст сообщения
    text = ""
    if event.message.body and event.message.body.text:
        text = event.message.body.text.strip()
    
    # Обработка ожидания ввода дат
    if state.get('awaiting_dates'):
        date_pattern = r'с\s+(\S+)\s+по\s+(\S+)'
        match = re.search(date_pattern, text, re.IGNORECASE)
        
        if match:
            date_from_str = match.group(1)
            date_to_str = match.group(2)
            
            date_from = parse_date(date_from_str)
            date_to = parse_date(date_to_str)
            
            if date_from and date_to:
                if date_from <= date_to:
                    user_states[chat_id] = {'date_from': date_from, 'date_to': date_to, 'awaiting_file': True}
                    send_message(chat_id, f"✅ Выбран период: {date_from.date()} - {date_to.date()}\n\n📎 Теперь загрузите файл с себестоимостью (xlsx/csv)")
                else:
                    send_message(chat_id, "❌ Дата начала не может быть позже даты окончания. Попробуйте ещё раз.")
            else:
                send_message(chat_id, "❌ Не удалось распознать даты. Используйте формат: с 2024-01-01 по 2024-01-31")
        else:
            send_message(chat_id, "❌ Неверный формат. Используйте: с 2024-01-01 по 2024-01-31")
        return
    
    # Обработка ожидания файла
    if state.get('awaiting_file'):
        # Проверяем attachments в body
        attachments = None
        if hasattr(event.message, 'body') and hasattr(event.message.body, 'attachments'):
            attachments = event.message.body.attachments
            logger.info(f"Найдены attachments: {attachments}")
        
        if attachments:
            for attachment in attachments:
                # attachment это объект File
                if hasattr(attachment, 'type') and attachment.type == 'file':
                    if hasattr(attachment, 'payload'):
                        payload = attachment.payload
                        file_token = payload.token if hasattr(payload, 'token') else None
                        file_url = payload.url if hasattr(payload, 'url') else None
                        file_name = attachment.filename if hasattr(attachment, 'filename') else 'file.xlsx'
                        
                        logger.info(f"Найден файл: {file_name}, токен: {file_token}, url: {file_url}")
                        
                        if file_token and file_url:
                            send_message(chat_id, f"📎 Скачиваю файл {file_name}...")
                            
                            temp_dir = Path(os.getenv('TEMP_FOLDER', './temp'))
                            temp_dir.mkdir(exist_ok=True)
                            local_path = temp_dir / f"cost_{chat_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                            
                            # Скачиваем файл по URL
                            success = await download_file_from_max(file_url, local_path)
                            
                            if success:
                                send_message(chat_id, f"📎 Файл получен! Начинаю генерацию отчёта...")
                                
                                await generate_report(
                                    chat_id, 
                                    state['date_from'], 
                                    state['date_to'], 
                                    str(local_path)
                                )
                                user_states.pop(chat_id, None)
                            else:
                                send_message(chat_id, "❌ Ошибка при скачивании файла")
                        else:
                            send_message(chat_id, "❌ Не удалось получить URL или токен файла")
                    else:
                        logger.error("Attachment не имеет payload")
                    return
                else:
                    logger.info(f"Attachment не является файлом: {attachment.type if hasattr(attachment, 'type') else 'unknown'}")
        else:
            logger.info("Attachments не найдены в сообщении")
            # Если нет файла, но пользователь отправил текст
            if text and not text.startswith('/'):
                send_message(chat_id, "❌ Пожалуйста, загрузите файл в формате xlsx или csv\n\nИспользуйте /cancel для отмены")
        return
    
    # Если нет активного состояния
    if text and not text.startswith('/'):
        send_message(chat_id, "🤔 Не понял команду.\n\nИспользуйте /start для просмотра доступных команд")


# ==================== ЗАПУСК ====================

async def main():
    """Точка входа"""
    logger.info("🚀 Запуск Ozon Finance Bot для MAX")
    
    print("✅ Бот инициализирован, запускаю...")
    await bot.delete_webhook()
    print("🔄 Webhook удалён")
    print("👂 Бот слушает сообщения...")
    print("=" * 50)
    
    await dp.start_polling(bot)


if __name__ == '__main__':
    asyncio.run(main())
