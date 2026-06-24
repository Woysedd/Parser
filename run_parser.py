import asyncio
import os
import requests
import openpyxl
from playwright.async_api import async_playwright

# Твоя матрица устройств
BASE_DEVICES = [
    "iPhone 13 128Gb", "iPhone 15 128Gb", "iPhone 16 128Gb", "iPhone 16 256Gb", "iPhone Air 256Gb",
    "iPhone 17 128Gb eSIM", "iPhone 17 256Gb eSIM", "iPhone 17 256Gb SIM+eSIM", "iPhone 17 512Gb SIM+eSIM"
]
TARGET_DEVICES = list(BASE_DEVICES)
for model in ["iPhone 17 Pro", "iPhone 17 Pro Max"]:
    for mem in ["256Gb", "512Gb", "1Tb"]:
        for color in ["Black", "White", "Natural"]:
            for sim in ["eSIM", "SIM+eSIM"]:
                TARGET_DEVICES.append(f"{model} {mem} {color} {sim}")

SITES = {
    "re:luxon": {"url": "https://re-luxe42.ru/iphone/iphone-new", "item": ".product-thumb", "title": ".caption a", "price": ".price"},
    "re:premium": {"url": "https://repremium.ru/kemerovo/catalog/apple/iphone/", "item": ".catalog-item", "title": ".item-title", "price": ".price_val"},
    "Like Store": {"url": "https://kemerovo.lstore.ru/catalog/iphone_1/", "item": ".product-card", "title": ".product-card__title", "price": ".product-card__price-current"},
    "Prostore": {"url": "https://prostore-shop.ru/catalog_iphone", "item": ".t-store__card", "title": ".t-store__card__title", "price": ".t-store__card__price-value"},
    "KingStore Уфа": {"url": "https://kingstore.link/catalog/iphone/", "item": ".catalog-item", "title": ".item-title", "price": ".price_val"}
}

async def run_parser():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Linux; Android 14; SM-S928B) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Mobile Safari/537.36",
            viewport={"width": 393, "height": 851},
            is_mobile=True,
            has_touch=True
        )
        await context.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        results = {}
        for name, cfg in SITES.items():
            page = await context.new_page()
            try:
                # 10 секунд на сайт — если больше, пропускаем
                await page.goto(cfg["url"], timeout=10000)
                await asyncio.sleep(2)
                items = await page.query_selector_all(cfg["item"])
                site_data = {}
                for item in items:
                    t_el = await item.query_selector(cfg["title"])
                    p_el = await item.query_selector(cfg["price"])
                    if t_el and p_el:
                        t = await t_el.text_content()
                        p = await p_el.text_content()
                        site_data[t.lower().strip()] = "".join(filter(str.isdigit, p))
                results[name] = site_data
            except:
                results[name] = {}
            await page.close()
        await browser.close()

        # Excel
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Устройство"] + list(SITES.keys()))
        for dev in TARGET_DEVICES:
            row = [dev]
            for site in SITES.keys():
                found = next((p for t, p in results.get(site, {}).items() if all(w in t for w in dev.lower().split() if w not in ["iphone"])), "По запросу")
                row.append(f"{found[:-3]}.{found[-3:]}" if len(found) >= 4 else "По запросу")
            ws.append(row)
        
        wb.save("Price_Result.xlsx")
        
        # Отправка
        token, chat_id = os.environ.get('TELEGRAM_BOT_TOKEN'), os.environ.get('TELEGRAM_CHAT_ID')
        if token and chat_id:
            requests.post(f"https://api.telegram.org/bot{token}/sendDocument", 
                          data={"chat_id": chat_id}, files={"document": open("Price_Result.xlsx", "rb")})

if __name__ == "__main__":
    asyncio.run(run_parser()) 
