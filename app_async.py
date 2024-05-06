import asyncio
from dotenv import load_dotenv
from funcs_async import download_report, login_d365, process_report
# Cambiar el import a la versión asíncrona de Playwright
from playwright.async_api import async_playwright
import os

async def main():
    load_dotenv()
    url_erp, user_erp, passw_erp, url_blueline, user_blueline, pass_bluline = os.getenv("URL"), os.getenv("USER"), os.getenv("PASS"), os.getenv("url_blueline_prod"), os.getenv("user_blueline"), os.getenv("pass_blueline")
    
    async with async_playwright() as playwright:
        browser = await playwright.chromium.launch(headless=True)
        try:
            # Ejecutar download_report y login_d365 en paralelo
            await asyncio.gather(
                download_report(playwright, url_blueline, user_blueline, pass_bluline, browser),
                #login_d365(playwright, url_erp, user_erp, passw_erp, browser)
            )
        finally:
            process_report()
            await browser.close()

if __name__ == "__main__":
    asyncio.run(main())