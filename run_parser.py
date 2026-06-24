import asyncio
import os
import requests
import openpyxl
from bs4 import BeautifulSoup

# Матрица (все 45 вариантов)
BASE_DEVICES = ["iPhone 13 128Gb", "iPhone 15 128Gb", "iPhone 16 128Gb", "iPhone 16 256Gb", "iPhone Air 256Gb",
    "iPhone 17 128Gb eSIM", "iPhone 17 256Gb eSIM", "iPhone 17 256Gb SIM+eSIM", "iPhone 17 512Gb SIM+eSIM"]
TARGET_DEVICES = list(BASE_DEVICES)
for model in ["iPhone 17 Pro", "iPhone 17 Pro Max"]:
    for mem in ["256Gb", "512Gb", "1Tb"]:
        for color in ["Black", "White", "Natural"]:
            for sim in ["eSIM", "SIM+eSIM"]:
                TARGET_DEVICES.append(f"{model} {mem} {color} {sim}")

# Ссылки на Google Cache (работает безотказно)
SITES = {
    "re:luxon": "https://webcache.googleusercontent.com/search?q=cache:https://re-luxe42.ru/iphone/iphone-new",
    "re:premium": "https://webcache.googleusercontent.com/search?q=cache:https://repremium.ru/kemerovo/catalog/apple/iphone/",
    "Like Store": "https://webcache.googleusercontent.com/search?q=cache:https://kemerovo.lstore.ru/catalog/iphone_1/",
    "Prostore": "https://webcache.googleusercontent.com/search?q=cache:https://prostore-shop.ru/catalog_iphone",
    "KingStore Уфа": "https://webcache.googleusercontent.com/search?q=cache:https://kingstore.link/catalog/iphone/"
}

def clean_price(text):
    d = "".join(filter(str.isdigit, text))
    return f"{d[:-3]}.{d[-3:]}" if len(d) >= 4 else "По запросу"

def fetch_data():
    results = {}
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"}
    
    for name, url in SITES.items():
        try:
            resp = requests.get(url, headers=headers, timeout=15)
            soup = BeautifulSoup(resp.text, 'html.parser')
            
            # Универсальный сбор: ищем все блоки, где есть цена и текст
            site_data = {}
            # Ищем любые элементы, которые выглядят как карточки товаров
            for card in soup.select('.product-thumb, .catalog-item, .product-card, .t-store__card'):
                t_el = card.get_text(strip=True)
                # Ищем цену как 3-6 цифр подряд
                import re
                prices = re.findall(r'\d{3,6}', t_el)
                if prices:
                    site_data[t_el.lower()] = clean_price(prices[0])
            results[name] = site_data
        except:
            results[name] = {}
    return results

def main():
    results = fetch_data()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Устройство"] + list(SITES.keys()))
    
    for dev in TARGET_DEVICES:
        row = [dev]
        for site in SITES.keys():
            found = next((p for t, p in results.get(site, {}).items() if all(w in t for w in dev.lower().split() if w not in ["iphone"])), "По запросу")
            row.append(found)
        ws.append(row)
            
    file_name = "Price_Result.xlsx"
    wb.save(file_name)
    
    token, chat_id = os.environ.get('TELEGRAM_BOT_TOKEN'), os.environ.get('TELEGRAM_CHAT_ID')
    if token and chat_id:
        requests.post(f"https://api.telegram.org/bot{token}/sendDocument", 
                      data={"chat_id": chat_id}, files={"document": open(file_name, "rb")})

if __name__ == "__main__":
    main()
