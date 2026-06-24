import asyncio
import os
import requests
import openpyxl
from playwright.async_api import async_playwright

# [Твой список TARGET_DEVICES оставляем без изменений]

async def run_parser():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        # Упрощенный контекст: теперь нам не нужны костыли, 
        # так как весь трафик идет через WireGuard
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
        )
        
        results = {}
        for name, cfg in SITES.items():
            page = await context.new_page()
            try:
                await page.goto(cfg["url"], timeout=30000)
                await asyncio.sleep(3)
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
            except Exception as e:
                print(f"Ошибка {name}: {e}")
            await page.close()
        await browser.close()

        # [Код записи в Excel и отправки в ТГ оставляем прежним]
