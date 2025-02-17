import asyncio
from aiogram import Bot, Dispatcher
from aiogram.filters import Command
from aiogram.types import Message, FSInputFile
from serpapi import GoogleSearch
import pandas as pd
import os
from dotenv import load_dotenv
from url_util import find_telegram_links
import time

load_dotenv()

TELEGRAM_API_TOKEN = os.getenv("TELEGRAM_API_TOKEN")
SERP_API_KEY = os.getenv("SERP_API_KEY")

if not all([TELEGRAM_API_TOKEN, SERP_API_KEY]):
    raise ValueError("Отсутствуют необходимые токены в .env файле")

bot = Bot(token=TELEGRAM_API_TOKEN)
dp = Dispatcher()

@dp.message(Command("start"))
async def start_command(message: Message):
    await message.answer("Введите ключевое слово для поиска:")

@dp.message()
async def process_keyword(message: Message):
    try:
        keyword = message.text
        await message.answer(f"Выполняю поиск по ключевому слову: {keyword}...")
        
        # Поиск и сохранение результатов
        results = perform_google_search(keyword)
        file_path = save_results_to_excel(results, keyword)
        
        # Отправка файла
        await message.answer_document(FSInputFile(file_path))
        os.remove(file_path)
        
    except Exception as e:
        await message.answer("Произошла ошибка при выполнении поиска. Попробуйте позже.")

def perform_google_search(query: str):
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
        link = result.get("link", "")
        email = "Не указано"
        telegram = "Не указано"
        
        # Поиск email в тексте
        if "@" in snippet:
            email = next((word.strip(".,;()") for word in snippet.split() if "@" in word), "Не указано")
        
        # Сначала ищем в сниппете
        telegram_links = find_telegram_links(snippet + " " + link)
        if telegram_links:
            telegram = telegram_links[0]
        else:
            # Если в сниппете не нашли, пытаемся найти на странице
            page_telegram = find_telegram_on_page(link)
            if page_telegram:
                telegram = page_telegram
        
        # Добавляем задержку между запросами
        time.sleep(1)
        
        extracted_data.append({
            "Название": f'=HYPERLINK("{link}","{result.get("title", "")}")',
            "Описание": snippet,
            "Email": email,
            "Telegram": telegram
        })
    
    return extracted_data

def save_results_to_excel(results, keyword):
    file_path = f"search_results_{''.join(c for c in keyword if c.isalnum() or c in (' ', '-', '_'))}.xlsx"
    
    df = pd.DataFrame(results)
    df.to_excel(file_path, sheet_name='Результаты поиска', index=False)
    
    return file_path

if __name__ == "__main__":
    asyncio.run(dp.start_polling(bot))
