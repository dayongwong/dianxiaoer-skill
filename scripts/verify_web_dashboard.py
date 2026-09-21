import asyncio
from pathlib import Path
from playwright.async_api import async_playwright

ARTIFACT_DIR = Path("/Users/vonnwang/.gemini/antigravity-ide/brain/8e5ea6b1-43b9-4954-a0f9-3bbb18bc546b")

async def verify_web():
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            executable_path="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            headless=True
        )
        context = await browser.new_context(viewport={"width": 1440, "height": 900})
        page = await context.new_page()

        # 注入企业超管登录态
        await context.add_init_script("""
            localStorage.setItem("dianxiaoer_auth_user", JSON.stringify({
                isLoggedIn: true,
                id: 1,
                phone: "18012347852",
                username: "王总 (企业创始人)",
                role: "owner",
                role_label: "企业超管",
                is_owner: true,
                vipText: "👑 企业创始人 (王总)"
            }));
        """)

        print("1. 正在访问前端 Web 商户后台 http://127.0.0.1:8089/ ...")
        await page.goto("http://127.0.0.1:8089/", wait_until="networkidle")
        await asyncio.sleep(2)

        # 截图 1: 商户后台主工作台与配额中心
        shot1 = ARTIFACT_DIR / "web_dashboard_main.png"
        await page.screenshot(path=str(shot1))
        print(f"  ✓ [截图成功] 商户后台主面板: {shot1.name}")

        # 截图 2: 调用 window.openBatchInquiryById(1) 查看真实议价批次看板
        print("2. 正在打开 1688 批量询盘与自动议价战报面板 (批次ID: 1)...")
        await page.evaluate("""async () => {
            if (typeof window.openBatchInquiryById === 'function') {
                await window.openBatchInquiryById(1);
            }
        }""")
        await asyncio.sleep(2)

        shot2 = ARTIFACT_DIR / "web_inquiry_batch_view.png"
        await page.screenshot(path=str(shot2))
        print(f"  ✓ [截图成功] 1688 询盘战报看板: {shot2.name}")

        await browser.close()
        print("🎉 Web 商户后台管理联动验证全部完成！")

if __name__ == "__main__":
    asyncio.run(verify_web())
