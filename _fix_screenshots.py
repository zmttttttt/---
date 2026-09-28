# -*- coding: utf-8 -*-
"""补截两张有误的截图：
   06_AI模型设置.png      —— 点开右上角徽章，弹出大模型 API 设置面板
   08_答题解析与错因诊断.png —— 在线练习答错后展开的解析 + 追问入口
"""
import asyncio
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
HTML_PATH = r"C:\Users\zmt\WorkBuddy\2026-09-08-20-01-33\小C老师-C语言选择结构智能助教.html"
HTML_URL = "file:///" + os.path.abspath(HTML_PATH).replace("\\", "/")
OUTDIR = r"C:\Users\zmt\WorkBuddy\2026-09-08-20-01-33\作品说明文档_截图"
os.makedirs(OUTDIR, exist_ok=True)

from playwright.async_api import async_playwright


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            executable_path=CHROME,
            headless=True,
            args=["--no-sandbox", "--disable-gpu", "--disable-crashpad",
                  "--disable-features=SpareRendererForSitePerProcess"],
        )
        ctx = await browser.new_context(
            viewport={"width": 1440, "height": 900},
            device_scale_factor=1.5, locale="zh-CN",
        )
        page = await ctx.new_page()
        await page.goto(HTML_URL, wait_until="networkidle", timeout=30000)
        await page.wait_for_timeout(1800)

        # ---------------- 06 大模型 API 设置面板 ----------------
        try:
            await page.click("#hdr-ai", timeout=5000)
            await page.wait_for_selector("#ai-modal.show", timeout=5000)
            await page.wait_for_timeout(1200)
            await page.screenshot(path=os.path.join(OUTDIR, "06_AI模型设置.png"))
            print("OK  06_AI模型设置.png")
        except Exception as exc:
            print("FAIL 06:", exc)

        # 关闭弹窗
        try:
            await page.evaluate(
                "() => { document.querySelectorAll('.modal-mask')"
                ".forEach(m => m.classList.remove('show')); }"
            )
            await page.wait_for_timeout(500)
        except Exception:
            pass

        # ---------------- 08 答题解析与错因诊断 ----------------
        try:
            await page.click('.tab[data-page="quiz"]', timeout=5000)
            await page.wait_for_timeout(1200)
            # 选第一个选项（Q1 正确答案是 B，选 A 即答错）
            await page.click("#opts .opt >> nth=0", timeout=5000)
            await page.wait_for_timeout(1200)
            await page.evaluate("() => window.scrollTo(0, 0)")
            await page.wait_for_timeout(400)
            await page.screenshot(path=os.path.join(OUTDIR, "08_答题解析与错因诊断.png"))
            print("OK  08_答题解析与错因诊断.png")
        except Exception as exc:
            print("FAIL 08:", exc)

        await ctx.close()
        await browser.close()

    print("\n=== 截图目录 ===")
    for name in sorted(os.listdir(OUTDIR)):
        size = os.path.getsize(os.path.join(OUTDIR, name))
        print("  %-34s %6.0f KB" % (name, size / 1024.0))


asyncio.run(main())
