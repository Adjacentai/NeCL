import asyncio
from aiogram import Bot, Dispatcher
from aiogram.filters import Command
from aiogram.types import Message, FSInputFile
from serpapi import GoogleSearch
import pandas as pd
import os
import logging
from typing import List, Dict
from dotenv import load_dotenv
import openpyxl

load_dotenv()

# Укажите ваш API-токен от Telegram
TELEGRAM_API_TOKEN = os.getenv("TELEGRAM_API_TOKEN")
SERP_API_KEY = os.getenv("SERP_API_KEY")

# Проверяем наличие токенов
if not TELEGRAM_API_TOKEN:
    raise ValueError("TELEGRAM_API_TOKEN не указан в .env файле")
if not SERP_API_KEY:
    raise ValueError("SERP_API_KEY не указан в .env файле")

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('bot.log'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

# Инициализация бота и диспетчера
bot = Bot(token=TELEGRAM_API_TOKEN)
dp = Dispatcher()

# Обновляем обработчик команды /start
@dp.message(Command("start"))
async def start_command(message: Message):
    await message.answer("Введите ключевое слово для поиска:")

# Обновляем обработчик текстовых сообщений
@dp.message()
async def process_keyword(message: Message):
    try:
        keyword = message.text
        await message.answer(f"Выполняю поиск по ключевому слову: {keyword}...")

        # Выполняем поиск
        search_results = perform_google_search(keyword)
        
        # Обрабатываем результаты
        file_path = save_results_to_excel(search_results, keyword)
        
        # Создаем InputFile объект и отправляем файл
        document = FSInputFile(file_path)
        await message.answer_document(document)
        
        # Удаляем временный файл
        os.remove(file_path)
    except Exception as e:
        logger.error(f"Ошибка при обработке запроса: {e}")
        await message.answer("Произошла ошибка при выполнении поиска. Попробуйте позже.")

# Функция выполнения поиска через SerpAPI
def perform_google_search(query: str) -> List[Dict]:
    try:
        params = {
            "engine": "google",
            "q": query,
            "location": "Moscow,Russia",
            "hl": "ru",
            "gl": "ru",
            "num": 10,
            "api_key": SERP_API_KEY
        }
        
        search = GoogleSearch(params)
        results = search.get_dict()
        extracted_data = []
        
        for result in results.get("organic_results", []):
            snippet = result.get("snippet", "")
            
            # Ищем email
            email = "Не указано"
            if "@" in snippet:
                words = snippet.split()
                for word in words:
                    if "@" in word:
                        email = word.strip(".,;()")
                        break
            
            # Ищем Telegram
            telegram = "Не указано"
            telegram_patterns = ["t.me/", "telegram.me/", "@"]
            for pattern in telegram_patterns:
                if pattern in snippet:
                    words = snippet.split()
                    for word in words:
                        if pattern in word:
                            # Очищаем имя пользователя от лишних символов
                            telegram = word.strip(".,;()[]{}").replace('t.me/', '@').replace('telegram.me/', '@')
                            if not telegram.startswith('@'):
                                telegram = '@' + telegram
                            break
                    if telegram != "Не указано":
                        break
            
            # Создаем формулу Excel для гиперссылки
            link = result.get("link", "Не указано")
            url_formula = f'=HYPERLINK("{link}","Перейти")'
            
            extracted_data.append({
                "Название": result.get("title", "Не указано"),
                "Ссылка": url_formula,
                "Описание": snippet,
                "Email": email,
                "Telegram": telegram
            })
        return extracted_data
    except Exception as e:
        logger.error(f"Ошибка при выполнении поиска: {e}")
        raise

# Сохранение результатов в Excel
def save_results_to_excel(results, keyword):
    try:
        safe_keyword = ''.join(c for c in keyword if c.isalnum() or c in (' ', '-', '_'))
        file_path = f"search_results_{safe_keyword}.xlsx"
        
        # Добавляем колонку Telegram
        df = pd.DataFrame(results, columns=["Название", "Ссылка", "Описание", "Email", "Telegram"])
        
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Результаты поиска', index=False)
            worksheet = writer.sheets['Результаты поиска']
            
            # Настраиваем ширину колонок
            worksheet.column_dimensions['A'].width = 40
            worksheet.column_dimensions['B'].width = 20
            worksheet.column_dimensions['C'].width = 60
            worksheet.column_dimensions['D'].width = 30
            worksheet.column_dimensions['E'].width = 30  # Для Telegram
            
            # Применяем форматирование
            for row in worksheet.iter_rows():
                for cell in row:
                    cell.alignment = openpyxl.styles.Alignment(wrap_text=True)
        
        return file_path
    except Exception as e:
        logger.error(f"Ошибка при сохранении результатов в Excel: {e}")
        raise

# Основной асинхронный цикл
async def main():
    try:
        # Запускаем бота
        await dp.start_polling(bot)
    except Exception as e:
        logger.error(f"Ошибка при запуске бота: {e}")
    finally:
        await bot.session.close()

# Запуск бота
if __name__ == "__main__":
    asyncio.run(main())
