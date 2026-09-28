# -*- coding: utf-8 -*-
"""补截 在线练习(quiz) 与 学情分析(dash) 两页。"""
import sys, os, io, asyncio
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
HTML_URL = "file:///" + os.path.abspath(r"C:\Users\zmt\WorkBuddy\2026-09-08-20-01-33\小C老师-C语言选择结构智能助教.html").replace("\\", "/")
OUTDIR = r"C:\Users\zmt\WorkBuddy\2026-09-08-20-01-33\作品说明文档_截图"

from playwright.async_api import async_playwright

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            executable_path=CHROME, headless=True,
            args=['--no-sandbox', '--disable-gpu', '--disable-crashpad',
                  '--disable-features=SpareRendererForSitePerProcess']
        )
        ctx = await browser.new_context(viewport={"width": 1440, "height": 900}, device_scale_factor=1.5, locale='zh-CN')
        page = await ctx.new_page()
        await page.goto(HTML_URL, wait_until='networkidle', timeout=20000)
        await page.wait_for_timeout(1500)

        # 在线练习 quiz
        await page.click('.tab[data-page="quiz"]')
        await page.wait_for_timeout(1500)
        await page.screenshot(path=os.path.join(OUTDIR, "03_在线练习.png"))
        print("✓ 03_在线练习.png")

        # 学情分析 dash
        await page.click('.tab[data-page="dash"]')
        await page.wait_for_timeout(1500)
        await page.screenshot(path=os.path.join(OUTDIR, "05_学情分析.png"))
        print("✓ 05_学情分析.png")

        await browser.close()

asyncio.run(main())