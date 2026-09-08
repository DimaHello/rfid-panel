import { chromium } from 'playwright';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.join(__dirname, '..');

(async () => {
    const browser = await chromium.launch({ headless: true });
    const page = await browser.newPage();
    await page.setViewportSize({ width: 1366, height: 860 });
    
    const indexPath = path.join(projectRoot, 'index.html');
    await page.goto(`file://${indexPath}`);
    
    // Switch to English and populate clean English names
    await page.evaluate(() => {
        window.setLanguage('en');
        window.accounts = [
            {
                id: 1,
                fio: "Anna Smirnova",
                phone: "+7 (926) 765-43-21",
                createdAt: "29.05.2026",
                expanded: true,
                children: [
                    { id: 11, name: "Egor", dob: "2016-08-10", status: "Avatar", rfid: "RFID_EGOR", uid: "EGOR20160810LCRXLVFEG2123" },
                    { id: 12, name: "Sophia", dob: "2019-01-25", status: "None", rfid: null, uid: "SOFIA20190125LCRXLVFEG2123" }
                ]
            },
            {
                id: 2,
                fio: "Dmitry Kuznetsov",
                phone: "+7 (903) 555-11-22",
                createdAt: "22.05.2026",
                expanded: false,
                children: [
                    { id: 21, name: "Artem", dob: "2014-05-12", status: "Bracelet", rfid: "RFID_ARTEM", uid: "ARTEM20140512LCRXLVFEG2123" },
                    { id: 22, name: "Victoria", dob: "2017-09-30", status: "None", rfid: null, uid: "VICTORIA20170930LCRXLV" },
                    { id: 23, name: "Ilya", dob: "2020-12-05", status: "None", rfid: null, uid: "ILYA20201205LCRXLVFEG2123" }
                ]
            },
            {
                id: 3,
                fio: "Mikhail Ivanov",
                phone: "+7 (916) 123-45-67",
                createdAt: "18.05.2026",
                expanded: false,
                children: [
                    { id: 31, name: "Alice", dob: "2015-11-15", status: "Avatar", rfid: "RFID_ALISA", uid: "ALISA20151115LCRXLVFEG212" },
                    { id: 32, name: "Max", dob: "2018-04-20", status: "Bracelet", rfid: "RFID_MAXIM", uid: "MAXIM20180420LCRXLVFEG212" }
                ]
            }
        ];
        window.renderAccounts();
    });
    
    await page.waitForTimeout(500);

    // 1. Screenshot of the Main English Table
    const mainScreenPath = path.join(projectRoot, 'docs', 'screenshots', 'main_screen_en.png');
    await page.screenshot({ path: mainScreenPath });
    console.log("Main screen EN saved:", mainScreenPath);

    // 2. Open Wristband Binding Modal for Sophia (id 12)
    await page.evaluate(() => {
        if (typeof window.openAttachModal === 'function') {
            window.openAttachModal(1, 12, 'Sophia');
        } else {
            // fallback: trigger button click
            const attachBtns = Array.from(document.querySelectorAll('button'));
            const btn = attachBtns.find(b => b.textContent.includes('Attach') || b.textContent.includes('Wristband') || b.textContent.includes('Привязать'));
            if (btn) btn.click();
        }
    });

    await page.waitForTimeout(600);

    // 2. Screenshot of the Binding Modal
    const bindModalPath = path.join(projectRoot, 'docs', 'screenshots', 'bind_bracelet_en.png');
    const modalEl = page.locator('#attach-modal > div, #chain-binding-modal > div, .rfid-modal-content, #attach-modal').first();
    
    if (await modalEl.count() > 0 && await modalEl.isVisible()) {
        await page.screenshot({ path: bindModalPath });
    } else {
        await page.screenshot({ path: bindModalPath });
    }
    console.log("Bind modal EN saved:", bindModalPath);

    await browser.close();
})();
