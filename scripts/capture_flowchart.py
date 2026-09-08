import asyncio
from playwright.async_api import async_playwright
import os

async def main():
    scripts_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(scripts_dir)
    html_path = os.path.join(scripts_dir, 'render_flowchart.html')
    output_path = os.path.join(project_dir, 'docs', 'screenshots', 'user_flow_diagram.png')
    brain_path = '/Users/dima/.gemini/antigravity/brain/34855149-e773-4669-80ec-f5eb04043611/user_flow_diagram.png'

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--disable-gpu']
        )
        page = await browser.new_page(device_scale_factor=2)
        await page.goto(f'file://{html_path}')
        await page.wait_for_selector('.mermaid svg', timeout=10000)
        await page.wait_for_timeout(500)

        element = await page.query_selector('.mermaid')
        await element.screenshot(path=output_path)
        await element.screenshot(path=brain_path)
        await browser.close()
        print("Flowchart diagram captured successfully!")

if __name__ == '__main__':
    asyncio.run(main())
