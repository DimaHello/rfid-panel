import asyncio
from playwright.async_api import async_playwright
import os

async def main():
    scripts_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(scripts_dir)
    output_path = os.path.join(project_dir, 'docs', 'screenshots', 'main_screen_en.png')
    brain_path = '/Users/dima/.gemini/antigravity/brain/34855149-e773-4669-80ec-f5eb04043611/rfid_panel_screenshot.png'

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--disable-gpu']
        )
        page = await browser.new_page(viewport={'width': 1366, 'height': 800})
        await page.goto('http://127.0.0.1:3333/index.html')
        await page.wait_for_timeout(600)

        await page.evaluate("""() => {
            document.getElementById('lang-en-btn').click();
            const badge = document.getElementById('pos-token-label');
            if (badge) badge.textContent = 'ACTIVATIONS:';

            // Localize parent and children names in DOM
            const replacements = [
                ['Смирнова Анна Юрьевна', 'Anna Smirnova'],
                ['Кузнецов Дмитрий Алексеевич', 'Dmitry Kuznetsov'],
                ['Иванов Михаил Сергеевич', 'Mikhail Ivanov'],
                ['Егор', 'Egor'],
                ['София', 'Sophia'],
                ['Артем', 'Artem'],
                ['Виктория', 'Victoria'],
                ['Илья', 'Ilya'],
                ['Алиса', 'Alice'],
                ['Максим', 'Max']
            ];

            const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
            let node;
            while ((node = walker.nextNode())) {
                replacements.forEach(([ru, en]) => {
                    if (node.nodeValue.includes(ru)) {
                        node.nodeValue = node.nodeValue.replaceAll(ru, en);
                    }
                });
            }

            // Update avatar circle initials
            const allElements = document.querySelectorAll('*');
            allElements.forEach(el => {
                if (el.children.length === 0) {
                    if (el.textContent.trim() === 'С') el.textContent = 'A';
                    else if (el.textContent.trim() === 'К') el.textContent = 'D';
                    else if (el.textContent.trim() === 'И') el.textContent = 'M';
                }
            });
        }""")

        await page.wait_for_timeout(400)
        await page.screenshot(path=output_path)
        await page.screenshot(path=brain_path)
        await browser.close()
        print("Screenshots saved successfully with English initials!")

if __name__ == '__main__':
    asyncio.run(main())
