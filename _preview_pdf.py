"""Render PDF preview via Playwright (more reliable in sandbox)."""
import asyncio
from playwright.async_api import async_playwright
import os

pdf_path = r'C:\Users\zmt\WorkBuddy\2026-09-08-20-01-33\作品说明文档.pdf'
out_dir = r'C:\Users\zmt\WorkBuddy\2026-09-08-20-01-33\作品说明文档_pdf预览'
os.makedirs(out_dir, exist_ok=True)

CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
url = 'file:///' + pdf_path.replace('\\', '/')


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            executable_path=CHROME,
            args=['--no-sandbox', '--disable-gpu', '--disable-dev-shm-usage'],
        )
        context = await browser.new_context(viewport={'width': 900, 'height': 1280})
        page = await context.new_page()
        await page.goto(url, wait_until='load', timeout=30000)
        # PDF viewer takes a moment to render
        await page.wait_for_timeout(3000)
        # Get page 1, 2, 4, 7 (sample pages)
        for n in [1, 2, 4, 7]:
            shot_path = os.path.join(out_dir, f'pdf_p{n}.png')
            await page.evaluate(f'window.location.hash = "#page={n}"')
            await page.wait_for_timeout(1000)
            await page.screenshot(path=shot_path, full_page=False)
            print(f'Saved {shot_path}')
        await browser.close()


asyncio.run(main())