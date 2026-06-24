import asyncio
from playwright.async_api import async_playwright
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import datetime
import requests
import os
import re

HEADERS = ["Название устройства", "KingStore Кемерово (Наш магазин)", "re:luxon", "re:premium", "Like Store", "Prostore", "KingStore Уфа"]

# Цветовая разметка таблиц (пастельные тона)
FILL_YELLOW = PatternFill(start_color="FFF2CC", fill_type="solid")
FILL_GREEN = PatternFill(start_color="E2EFDA", fill_type="solid")
FILL_PINK = PatternFill(start_color="FCE4D6", fill_type="solid")
FILL_GREY = PatternFill(start_color="F2F2F2", fill_type="solid")
FILL_HEADER = PatternFill(start_color="A9D08E", fill_type="solid")

FONT_BOLD = Font(name="Arial", size=11, bold=True)
FONT_REGULAR = Font(name="Arial", size=10)
THIN_BORDER = Border(
    left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
    top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF')
)

# Список запрашиваемых моделей
DEVICES = [
    ("iPhone 13 128Gb", FILL_YELLOW),
    ("iPhone 15 128Gb", FILL_YELLOW),
    ("iPhone 16 128Gb", FILL_GREEN),
    ("iPhone 16 256Gb", FILL_GREEN),
    ("iPhone 17 256Gb eSIM", FILL_GREY),
    ("iPhone 17 512Gb eSIM", FILL_GREY),
    ("iPhone 17 256Gb SIM+eSIM", FILL_GREY),
    ("iPhone 17 512Gb SIM+eSIM", FILL_GREY),
]

async def parse_shop_price(page, url, device_name):
    """
    Заходит на сайт, ищет карточку товара с названием устройства и вытаскивает цену.
    """
    try:
        await page.goto(url, timeout=60000, wait_until="domcontentloaded")
        await asyncio.sleep(3) # Даем подгрузиться скриптам цен
        
        # Переводим поисковый запрос в регулярку (ищем, например, "13" и "128")
        keywords = device_name.replace("Gb", "").replace("eSIM", "").replace("SIM+eSIM", "").split()
        
        # Получаем все текстовые элементы с ценниками и названиями на странице
        elements = await page.query_selector_all("//*[contains(text(), 'iPhone')]")
        
        for el in elements:
            text = await el.text_content()
            if text and all(k.lower() in text.lower() for k in keywords):
                # Ищем родительский контейнер карточки товара, чтобы забрать цену
                parent = await el.evaluate_handle("el => el.closest('div')")
                parent_text = await parent.text_content()
                
                # Ищем цифры формата 45 000, 45.000, 129990
                prices = re.findall(r'(\d+[\s\.,]?\d{3})', parent_text)
                if prices:
                    # Очищаем цену от лишних пробелов и точек, приводим к стандарту "45.990"
                    raw_price = prices[0].replace(" ", "").replace(",", "").replace(".", "")
                    if len(raw_price) >= 5:
                        formatted_price = f"{raw_price[:-3]}.{raw_price[-3:]}"
                        return formatted_price
        return "По запросу"
    except Exception as e:
        return "Ошибка"

async def main():
    async with async_playwright() as p:
        # Запускаем браузер в режиме маскировки под обычного пользователя
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            viewport={"width": 1920, "height": 1080}
        )
        page = await context.new_page()
        
        # Актуальные каталоги айфонов на сайтах конкурентов
        urls = {
            "re:luxon": "https://re-luxe42.ru/iphone/iphone-new",
            "re:premium": "https://repremium.ru/kemerovo/catalog/apple/iphone/",
            "Like Store": "https://kemerovo.lstore.ru/catalog/iphone_1/",
            "Prostore": "https://prostore-shop.ru/catalog_iphone",
            "KingStore Уфа": "https://kingstore.link/catalog/iphone/"
        }
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Срез цен"
        ws.views.sheetView[0].showGridLines = True
        
        # Сборка шапки
        ws.append(HEADERS)
        for c_idx in range(1, len(HEADERS) + 1):
            cell = ws.cell(row=1, column=c_idx)
            cell.fill = FILL_HEADER
            cell.font = FONT_BOLD
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="left" if c_idx == 1 else "center", vertical="center")
        
        # Обход моделей и сайтов
        for dev_name, fill in DEVICES:
            row_data = [dev_name, ""] # Наш магазин — пустая колонка для ручного ввода
            
            for shop in HEADERS[2:]:
                url = urls.get(shop)
                price = await parse_shop_price(page, url, dev_name)
                row_data.append(price)
                
            ws.append(row_data)
            
            # Стилизация строк табличной части
            r_idx = ws.max_row
            for c_idx in range(1, len(HEADERS) + 1):
                c = ws.cell(row=r_idx, column=c_idx)
                c.fill = fill
                c.font = FONT_REGULAR
                c.border = THIN_BORDER
                c.alignment = Alignment(horizontal="left" if c_idx == 1 else "center", vertical="center")
        
        # Красивое выравнивание ширины столбцов под текст
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 16)
        
        date_str = datetime.datetime.now().strftime("%Y-%m-%d")
        file_name = f"Прайс_Конкурентов_{date_str}.xlsx"
        wb.save(file_name)
        await browser.close()
        
        # Пересылка файла в Telegram
        token = os.environ['TELEGRAM_BOT_TOKEN']
        chat_ids = os.environ['TELEGRAM_CHAT_ID'].split(',')
        url_tg = f"https://api.telegram.org/bot{token}/sendDocument"
        
        for chat_id in chat_ids:
            chat_id = chat_id.strip()
            if chat_id:
                with open(file_name, "rb") as f:
                    requests.post(url_tg, data={"chat_id": chat_id, "caption": f"📊 Полный срез цен в Кемерово на {date_str} готов!"}, files={"document": f})

if __name__ == "__main__":
    asyncio.run(main())
