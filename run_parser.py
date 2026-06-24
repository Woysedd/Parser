import asyncio
from playwright.async_api import async_playwright
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import datetime
import requests
import os
import re

HEADERS = ["Название устройства", "re:luxon", "re:premium", "Like Store", "Prostore", "KingStore Уфа"]

FILL_YELLOW = PatternFill(start_color="FFF2CC", fill_type="solid") 
FILL_GREEN = PatternFill(start_color="E2EFDA", fill_type="solid")  
FILL_GREY = PatternFill(start_color="F2F2F2", fill_type="solid")   
FILL_PINK = PatternFill(start_color="FCE4D6", fill_type="solid")   
FILL_HEADER = PatternFill(start_color="A9D08E", fill_type="solid")

FONT_BOLD = Font(name="Arial", size=11, bold=True)
FONT_REGULAR = Font(name="Arial", size=10)
THIN_BORDER = Border(
    left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
    top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF')
)

PRO_COLORS = ["Black", "White", "Natural", "Gold"]
PRO_SIZES = ["256Gb", "512Gb", "1Tb"]
SIM_TYPES = ["eSIM", "SIM+eSIM"]

# Строго твой список моделей
BASE_DEVICES = [
    ("iPhone 13 128Gb", FILL_YELLOW),
    ("iPhone 15 128Gb", FILL_YELLOW),
    ("iPhone 16 128Gb", FILL_GREEN),
    ("iPhone 16 256Gb", FILL_GREEN),
    ("iPhone Air 256Gb", FILL_GREEN),
    ("iPhone 17 128Gb eSIM", FILL_GREY),
    ("iPhone 17 256Gb eSIM", FILL_GREY),
    ("iPhone 17 256Gb SIM+eSIM", FILL_GREY),
    ("iPhone 17 512Gb SIM+eSIM", FILL_GREY),
]

DEVICES = list(BASE_DEVICES)
for model in ["iPhone 17 Pro", "iPhone 17 Pro Max"]:
    for size in PRO_SIZES:
        for color in PRO_COLORS:
            for sim in SIM_TYPES:
                DEVICES.append((f"{model} {size} {color} {sim}", FILL_PINK))

def clean_price(price_str):
    if not price_str:
        return "По запросу"
    digits = "".join([c for c in price_str if c.isdigit()])
    if not digits or len(digits) < 4:
        return "По запросу"
    if len(digits) > 6:
        digits = digits[:6]
    return f"{digits[:-3]}.{digits[-3:]}"

async def parse_universal(page, url, device_name):
    """Глубокий универсальный поиск по тексту всей страницы без привязки к селекторам"""
    try:
        await page.goto(url, timeout=60000, wait_until="load")
        await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        await asyncio.sleep(3) # Ждем подгрузки всех динамических прайсов
        
        # Забираем вообще весь видимый текст с сайта
        body_text = await page.inner_text("body")
        lines = body_text.split('\n')
        
        name_lower = device_name.lower()
        # Определяем маркеры для поиска
        tokens = name_lower.split()
        
        # Синонимы цветов для русскоязычной верстки конкурентов
        color_synonyms = {
            "black": ["черн", "black"],
            "white": ["бел", "white", "silver", "сереб"],
            "natural": ["титан", "natural", "нат"],
            "gold": ["золот", "gold", "desert"]
        }
        
        best_price = None
        
        # Перебираем строки текста в поисках совпадений
        for i, line in enumerate(lines):
            line_l = line.lower()
            
            # Проверяем основные токены (модель, память)
            if not all(t in line_l for t in tokens if t not in ["esim", "sim+esim"] and t not in color_synonyms):
                continue
                
            # Проверяем тип сим-карты отдельно
            if "sim+esim" in name_lower and "sim" not in line_l:
                continue
            if "esim" in name_lower and "sim+esim" not in name_lower and "esim" not in line_l:
                continue
                
            # Проверяем цвет, если он указан в модели
            color_match = True
            for eng_color, rus_list in color_synonyms.items():
                if eng_color in name_lower:
                    if not any(r in line_l for r in rus_list):
                        color_match = False
            if not color_match:
                continue
                
            # Если нашли строку с товаром, ищем цену в радиусе 3 строк вокруг неё
            for offset in range(-1, 4):
                if 0 <= i + offset < len(lines):
                    check_line = lines[i + offset]
                    # Ищем подстроку, похожую на ценник (от 30 000 до 350 000)
                    nums = "".join([c for c in check_line if c.isdigit()])
                    if nums and 4 <= len(nums) <= 6:
                        val = int(nums)
                        if 30000 <= val <= 350000:
                            best_price = clean_price(check_line)
                            break
            if best_price:
                break
                
        return best_price if best_price else "По запросу"
    except:
        return "Ошибка"

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
        )
        page = await context.new_page()
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Срез цен"
        ws.views.sheetView[0].showGridLines = True
        
        ws.append(HEADERS)
        for c_idx in range(1, len(HEADERS) + 1):
            cell = ws.cell(row=1, column=c_idx)
            cell.fill = FILL_HEADER
            cell.font = FONT_BOLD
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="left" if c_idx == 1 else "center", vertical="center")
        
        urls = {
            "re:luxon": "https://re-luxe42.ru/iphone/iphone-new",
            "re:premium": "https://repremium.ru/kemerovo/catalog/apple/iphone/",
            "Like Store": "https://kemerovo.lstore.ru/catalog/iphone_1/",
            "Prostore": "https://prostore-shop.ru/catalog_iphone",
            "KingStore Уфа": "https://kingstore.link/catalog/iphone/"
        }
        
        for dev_name, fill in DEVICES:
            print(f"Ищем актуальные цены для: {dev_name}...")
            
            p_luxon = await parse_universal(page, urls["re:luxon"], dev_name)
            p_premium = await parse_universal(page, urls["re:premium"], dev_name)
            p_like = await parse_universal(page, urls["Like Store"], dev_name)
            p_pro = await parse_universal(page, urls["Prostore"], dev_name)
            p_king_ufa = await parse_universal(page, urls["KingStore Уфа"], dev_name)
            
            ws.append([dev_name, p_luxon, p_premium, p_like, p_pro, p_king_ufa])
            
            r_idx = ws.max_row
            for c_idx in range(1, len(HEADERS) + 1):
                c = ws.cell(row=r_idx, column=c_idx)
                c.fill = fill
                c.font = FONT_REGULAR
                c.border = THIN_BORDER
                c.alignment = Alignment(horizontal="left" if c_idx == 1 else "center", vertical="center")
                
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 18)
        
        date_str = datetime.datetime.now().strftime("%Y-%m-%d")
        file_name = f"Прайс_Конкурентов_{date_str}.xlsx"
        wb.save(file_name)
        await browser.close()
        
        token = os.environ['TELEGRAM_BOT_TOKEN']
        chat_ids = os.environ['TELEGRAM_CHAT_ID'].split(',')
        url_tg = f"https://api.telegram.org/bot{token}/sendDocument"
        
        for chat_id in chat_ids:
            chat_id = chat_id.strip()
            if chat_id:
                with open(file_name, "rb") as f:
                    requests.post(url_tg, data={"chat_id": chat_id, "caption": f"📊 Свежий срез цен реального времени на {date_str} готов!"}, files={"document": f})

if __name__ == "__main__":
    asyncio.run(main())
