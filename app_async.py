import asyncio
from dotenv import load_dotenv
from funcs_async import download_report, login_d365, process_report
from playwright.async_api import async_playwright
import os
import argparse

async def run_task(task_func, url, user, password, playwright):
    browser = await playwright.chromium.launch(headless=False)
    try:
        await task_func(playwright, url, user, password, browser)
    finally:
        await browser.close()

async def main(task_name):
    load_dotenv()
    url_erp, user_erp, passw_erp, url_blueline, user_blueline, pass_bluline = os.getenv("url_erp_test"), os.getenv("USER_DIEGO"), os.getenv("PASS_DIEGO"), os.getenv("url_blueline_prod"), os.getenv("user_blueline"), os.getenv("pass_blueline")

    task_map = {
        'download_report': (download_report, url_blueline, user_blueline, pass_bluline),
        'login_d365': (login_d365, url_erp, user_erp, passw_erp),
        'process_report': (process_report,)  # Asume que process_report no necesita navegador ni argumentos adicionales
    }

    async with async_playwright() as playwright:
        if task_name in ['download_report', 'login_d365']:
            func, url, user, passw = task_map[task_name]
            await run_task(func, url, user, passw, playwright)
        elif task_name == 'process_report':
            func = task_map[task_name][0]
            func()  # Ejecutar funciones que no necesitan navegador

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Run specific tasks with Playwright.')
    parser.add_argument('task', choices=['download_report', 'login_d365', 'process_report'], help='The task to run')
    args = parser.parse_args()
    asyncio.run(main(args.task))



