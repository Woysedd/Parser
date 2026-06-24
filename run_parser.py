import asyncio
import os
import requests
import openpyxl
from playwright.async_api import async_playwright

# Твой список устройств
TARGET_DEVICES = ["iPhone 13 128Gb", "iPhone 15 128Gb", "iPhone 16 128Gb", "iPhone 16 256Gb", "iPhone 17 Pro Max 256Gb"]

SITES = {
    "re:luxon": {"url": "https://re-luxe42.ru/iphone/iphone-new", "item": ".product-thumb", "title": ".caption a", "price": ".price"},
    "re:premium": {"url": "https://repremium.ru/kemerovo/catalog/apple/iphone/", "item": ".catalog-item", "title": ".item-title", "price": ".price_val"},
    "Like Store": {"url": "https://kemerovo.lstore.ru/catalog/iphone_1/", "item": ".product-card", "title": ".product-card__title", "price": ".product-card__price-current"}
}

async def run():
    async with async_playwright() as p:
        # Браузер без прокси, чтобы не было ошибок сети
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
        
        results = {}
        for name, cfg in SITES.items():
            page = await context.new_page()
            try:
                await page.goto(cfg["url"], timeout=40000)
                await asyncio.sleep(5)
                items = await page.query_selector_all(cfg["item"])
                for item in items:
                    t_el = await item.query_selector(cfg["title"])
                    p_el = await item.query_selector(cfg["price"])
                    if t_el and p_el:
                        t = await t_el.text_content()
                        p = await p_el.text_content()
                        results[t.strip()] = p.strip()
            except Exception as e:
                print(f"Ошибка на {name}: {e}")
            await page.close()
        
        await browser.close()
        
        # Запись в Excel
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Название", "Цена"])
        for k, v in results.items():
            ws.append([k, v])
        wb.save("Price.xlsx")
        
        # Отправка
        token = os.environ.get('TELEGRAM_BOT_TOKEN')
        chat_id = os.environ.get('TELEGRAM_CHAT_ID')
        if token and chat_id:
            requests.post(f"https://api.telegram.org/bot{token}/sendDocument", 
                          data={"chat_id": chat_id}, files={"document": open("Price.xlsx", "rb")})

if __name__ == "__main__":
    asyncio.run(run())
