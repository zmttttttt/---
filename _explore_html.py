# -*- coding: utf-8 -*-
"""探索 HTML 页面结构：找出 tab / 按钮 / 输入框。"""
import sys, os, io, asyncio
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
HTML_URL = "file:///" + os.path.abspath(r"C:\Users\zmt\WorkBuddy\2026-09-08-20-01-33\小C老师-C语言选择结构智能助教.html").replace("\\", "/")

from playwright.async_api import async_playwright

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(executable_path=CHROME, headless=True)
        ctx = await browser.new_context(viewport={"width": 1440, "height": 900})
        page = await ctx.new_page()
        await page.goto(HTML_URL, wait_until='networkidle')
        await page.wait_for_timeout(1000)

        # 所有 button / tab / nav 链接
        info = await page.evaluate("""
        () => {
            const out = [];
            document.querySelectorAll('button, [role=tab], nav a, .tab, [onclick]').forEach((el, i) => {
                const r = el.getBoundingClientRect();
                out.push({
                    idx: i,
                    tag: el.tagName,
                    cls: (el.className||'').toString().slice(0,50),
                    id: el.id,
                    text: (el.textContent||'').trim().slice(0,40),
                    visible: r.width>0 && r.height>0,
                    rect: [Math.round(r.x), Math.round(r.y), Math.round(r.width), Math.round(r.height)]
                });
            });
            return out;
        }
        """)
        print(f"=== clickables: {len(info)} ===")
        for it in info:
            if it['visible']:
                print(f"  #{it['idx']} <{it['tag']} class='{it['cls']}' id='{it['id']}'> text='{it['text']}' rect={it['rect']}")

        # 输入框
        inputs = await page.evaluate("""
        () => [...document.querySelectorAll('textarea, input')].map(el => ({
            tag: el.tagName, type: el.type, placeholder: el.placeholder, id: el.id, cls: el.className.toString().slice(0,40)
        }))
        """)
        print(f"\n=== inputs: {len(inputs)} ===")
        for it in inputs:
            print(f"  <{it['tag']} type={it['type']}> placeholder='{it['placeholder']}' id='{it['id']}' cls='{it['cls']}'")

        await browser.close()

asyncio.run(main())