import asyncio
import requests
import openpyxl
from playwright.async_api import async_playwright
from fp.fp import FreeProxy

TARGET_DEVICES = ["iPhone 13 128Gb", "iPhone 15 128Gb", "iPhone 16 128Gb", "iPhone 16 256Gb", "iPhone 17 Pro Max 256Gb"]
SITES = {
    "re:luxon": {"url": "https://re-luxe42.ru/iphone/iphone-new", "item": ".product-thumb", "title": ".caption a", "price": ".price"},
    "re:premium": {"url": "https://repremium.ru/kemerovo/catalog/apple/iphone/", "item": ".catalog-item", "title": ".item-title", "price": ".price_val"},
    "Like Store": {"url": "https://kemerovo.lstore.ru/catalog/iphone_1/", "item": ".product-card", "title": ".product-card__title", "price": ".product-card__price-current"}
}

async def run():
    proxy = FreeProxy(country_id=['RU'], timeout=1).get()
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True, proxy={"server": proxy})
        context = await browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Устройство", "Цена"])

        for name, cfg in SITES.items():
            page = await context.new_page()
            try:
                await page.goto(cfg["url"], timeout=30000)
                await asyncio.sleep(5)
                items = await page.query_selector_all(cfg["item"])
                for item in items:
                    t = await item.query_selector(cfg["title"])
                    p_el = await item.query_selector(cfg["price"])
                    if t and p_el:
                        ws.append([await t.text_content(), await p_el.text_content()])
            except: pass
            await page.close()
        
        await browser.close()
        wb.save("Price.xlsx")
        
        token, chat_id = os.environ.get('TELEGRAM_BOT_TOKEN'), os.environ.get('TELEGRAM_CHAT_ID')
        if token:
            requests.post(f"https://api.telegram.org/bot{token}/sendDocument", 
                          data={"chat_id": chat_id}, files={"document": open("Price.xlsx", "rb")})

if __name__ == "__main__":
    asyncio.run(run())
