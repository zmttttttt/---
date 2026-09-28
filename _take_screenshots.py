# -*- coding: utf-8 -*-
"""使用 playwright 精确截取 HTML 应用的各功能页面（基于 data-page tab）。"""
import sys, os, io, asyncio
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
HTML_URL = "file:///" + os.path.abspath(r"C:\Users\zmt\WorkBuddy\2026-09-08-20-01-33\小C老师-C语言选择结构智能助教.html").replace("\\", "/")
OUTDIR = r"C:\Users\zmt\WorkBuddy\2026-09-08-20-01-33\作品说明文档_截图"
os.makedirs(OUTDIR, exist_ok=True)

from playwright.async_api import async_playwright

async def shot(page, name):
    path = os.path.join(OUTDIR, name)
    await page.screenshot(path=path, full_page=False)
    print(f"✓ {name}")

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            executable_path=CHROME, headless=True,
            args=['--no-sandbox', '--disable-gpu']
        )
        ctx = await browser.new_context(
            viewport={"width": 1440, "height": 900},
            device_scale_factor=1.5, locale='zh-CN'
        )
        page = await ctx.new_page()
        await page.goto(HTML_URL, wait_until='networkidle', timeout=20000)
        await page.wait_for_timeout(1500)

        # 1. 学习路径主界面
        await shot(page, "01_学习路径主界面.png")

        # 8. 任务详情（学一学）— 先截，之后回主界面
        try:
            await page.click('button:has-text("学一学") >> nth=0', timeout=3000)
            await page.wait_for_timeout(1200)
            await shot(page, "08_任务详情学习页.png")
            # 关闭可能的弹窗
            await page.keyboard.press("Escape")
            await page.evaluate("() => { document.querySelectorAll('.modal-mask').forEach(m => m.style.display = 'none'); }")
            await page.click('.tab[data-page="path"]')
            await page.wait_for_timeout(800)
        except Exception as e:
            print(f"✗ 08: {e}")

        # 2. 知识讲堂
        try:
            await page.click('.tab[data-page="know"]', timeout=3000)
            await page.wait_for_timeout(1200)
            await shot(page, "02_知识讲堂.png")
        except Exception as e:
            print(f"✗ 02: {e}")

        # 3. 在线练习
        try:
            await page.click('.tab[data-page="practice"]', timeout=3000)
            await page.wait_for_timeout(1200)
            await shot(page, "03_在线练习.png")
        except Exception as e:
            print(f"✗ 03: {e}")

        # 4. AI 答疑：输入问题 + 发送
        try:
            await page.click('.tab[data-page="chat"]', timeout=3000)
            await page.wait_for_timeout(800)
            await page.fill('#chat-in', 'if 和 else 怎么配对？请举个例子')
            await page.wait_for_timeout(400)
            await page.keyboard.press('Enter')
            await page.wait_for_timeout(4000)  # 等本地 AI 回复
            await shot(page, "04_AI答疑对话.png")
        except Exception as e:
            print(f"✗ 04: {e}")

        # 5. 学情分析
        try:
            await page.click('.tab[data-page="analysis"]', timeout=3000)
            await page.wait_for_timeout(1500)
            await shot(page, "05_学情分析.png")
        except Exception as e:
            print(f"✗ 05: {e}")

        # 6. AI 设置弹窗
        try:
            await page.click('.tab[data-page="path"]')
            await page.wait_for_timeout(500)
            await page.click('#hdr-ai', timeout=3000)
            await page.wait_for_timeout(1200)
            await shot(page, "06_AI模型设置.png")
        except Exception as e:
            print(f"✗ 06: {e}")

        await ctx.close()

        # 7. 手机视图
        try:
            ctx2 = await browser.new_context(
                viewport={"width": 390, "height": 844},
                device_scale_factor=2.0, is_mobile=True, has_touch=True, locale='zh-CN'
            )
            page2 = await ctx2.new_page()
            await page2.goto(HTML_URL, wait_until='networkidle', timeout=20000)
            await page2.wait_for_timeout(1500)
            await page2.screenshot(path=os.path.join(OUTDIR, "07_手机端视图.png"), full_page=False)
            print("✓ 07_手机端视图.png")
            await ctx2.close()
        except Exception as e:
            print(f"✗ 07: {e}")

        await browser.close()

        print(f"\n=== 截图列表 ===")
        for f in sorted(os.listdir(OUTDIR)):
            sz = os.path.getsize(os.path.join(OUTDIR, f))
            print(f"  {f} ({sz//1024} KB)")

asyncio.run(main())