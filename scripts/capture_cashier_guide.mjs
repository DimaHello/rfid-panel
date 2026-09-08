import { chromium } from 'playwright';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.join(__dirname, '..');
const outputDir = path.join(projectRoot, 'docs', 'screenshots', 'cashier_guide');

if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
}

// Helper to inject highlight annotation into page using Playwright Locator bounding box
async function addAnnotation(page, { locator, number, text, position = 'top', color = '#ff5a28', customStyle = '' }) {
    const box = await locator.boundingBox();
    if (!box) {
        console.warn('Bounding box not found for annotation');
        return;
    }

    await page.evaluate(({ b, num, txt, pos, col, style }) => {
        const overlay = document.createElement('div');
        overlay.className = 'cashier-guide-annotation';

        overlay.style.cssText = `
            position: absolute;
            left: ${window.scrollX + b.x - 6}px;
            top: ${window.scrollY + b.y - 6}px;
            width: ${b.width + 12}px;
            height: ${b.height + 12}px;
            border: 3.5px solid ${col};
            border-radius: 18px;
            box-shadow: 0 0 0 5px ${col}33, 0 8px 25px ${col}40;
            pointer-events: none;
            z-index: 999999;
            box-sizing: border-box;
            ${style}
        `;

        if (num) {
            const badge = document.createElement('div');
            badge.style.cssText = `
                position: absolute;
                top: -16px;
                left: -16px;
                width: 34px;
                height: 34px;
                background: ${col};
                color: #ffffff;
                border: 3px solid #ffffff;
                border-radius: 50%;
                font-family: 'Inter', system-ui, -apple-system, sans-serif;
                font-weight: 900;
                font-size: 16px;
                display: flex;
                align-items: center;
                justify-content: center;
                box-shadow: 0 4px 14px rgba(0,0,0,0.35);
                z-index: 1000000;
            `;
            badge.innerText = num;
            overlay.appendChild(badge);
        }

        if (txt) {
            const callout = document.createElement('div');
            let posCss = 'top: -38px; left: 24px;';
            if (pos === 'bottom') posCss = 'bottom: -38px; left: 10px;';
            else if (pos === 'top-center') posCss = 'top: -38px; left: 50%; transform: translateX(-50%);';
            else if (pos === 'bottom-center') posCss = 'bottom: -38px; left: 50%; transform: translateX(-50%);';
            else if (pos === 'right') posCss = 'top: 50%; transform: translateY(-50%); left: calc(100% + 14px);';
            else if (pos === 'left') posCss = 'top: 50%; transform: translateY(-50%); right: calc(100% + 14px);';

            callout.style.cssText = `
                position: absolute;
                ${posCss}
                background: #111827;
                color: #ffffff;
                padding: 6px 12px;
                border-radius: 20px;
                font-family: 'Inter', system-ui, -apple-system, sans-serif;
                font-weight: 700;
                font-size: 11.5px;
                letter-spacing: 0.01em;
                box-shadow: 0 6px 20px rgba(0,0,0,0.3);
                border: 1.5px solid rgba(255,255,255,0.2);
                white-space: nowrap;
                z-index: 1000000;
                display: flex;
                align-items: center;
                gap: 5px;
            `;
            callout.innerHTML = txt;
            overlay.appendChild(callout);
        }

        document.body.appendChild(overlay);
    }, { b: box, num: number, txt: text, pos: position, col: color, style: customStyle });
}

async function clearAnnotations(page) {
    await page.evaluate(() => {
        document.querySelectorAll('.cashier-guide-annotation').forEach(el => el.remove());
        const toast = document.getElementById('toast-container');
        if (toast) toast.innerHTML = '';
    });
}

async function capture() {
    console.log("🚀 Launching Chromium...");
    const browser = await chromium.launch({ headless: true });

    // =========================================================================
    // ЧАСТЬ 1: МОБИЛЬНЫЙ ПУТЬ САМОРЕГИСТРАЦИИ ГОСТЯ (ЦЕЛЕВОЙ СЦЕНАРИЙ)
    // =========================================================================
    console.log("📱 Снятие скриншотов мобильной регистрации гостя по QR...");
    const mobileContext = await browser.newContext({
        viewport: { width: 390, height: 844 },
        deviceScaleFactor: 2,
        isMobile: true,
        hasTouch: true
    });
    const mobilePage = await mobileContext.newPage();
    const guestIndexPath = path.join(projectRoot, 'guest-registration', 'index.html');
    await mobilePage.goto(`file://${guestIndexPath}`);
    await mobilePage.waitForTimeout(600);

    // Делаем фон чисто белым без серых полей вокруг и убираем лишние зазоры
    await mobilePage.evaluate(() => {
        document.documentElement.style.setProperty('background', '#ffffff', 'important');
        document.body.style.setProperty('background', '#ffffff', 'important');
        document.documentElement.style.setProperty('--bg', '#ffffff');
        const style = document.createElement('style');
        style.innerHTML = `
            html, body { background: #ffffff !important; margin: 0 !important; }
            .app { max-width: 375px !important; margin: 0 auto !important; padding: 8px 8px 90px 8px !important; background: #ffffff !important; }
            .card { box-shadow: 0 4px 16px rgba(0,0,0,0.05) !important; border: 1.5px solid #e2e8f0 !important; background: #ffffff !important; }
        `;
        document.head.appendChild(style);
    });

    // Скриншот 1: Мобильный экран - Ввод телефона + согласие + СМС
    console.log("📸 [QR 1.1] Телефон + СМС на смартфоне гостя...");
    await mobilePage.locator('#phone-input').fill('(926) 777-88-99');
    await mobilePage.locator('#check-mandatory').check();
    await mobilePage.waitForTimeout(200);
    await mobilePage.locator('#btn-send-sms').click();
    await mobilePage.waitForTimeout(300);
    await mobilePage.locator('#sms-input').fill('1234');
    await mobilePage.waitForTimeout(200);

    await addAnnotation(mobilePage, {
        locator: mobilePage.locator('#btn-verify'),
        number: '1',
        text: '👉 Гость вводит СМС и нажимает «Подтвердить»',
        position: 'top',
        color: '#ff5a28'
    });
    await mobilePage.screenshot({ path: path.join(outputDir, 'qr_1_guest_mobile_phone.png') });
    await clearAnnotations(mobilePage);

    // Переходим на Шаг 2 (ФИО + Дети)
    console.log("📸 [QR 1.2] ФИО и Дети на смартфоне гостя...");
    await mobilePage.locator('#btn-verify').click();
    await mobilePage.waitForTimeout(1000);

    await mobilePage.locator('#fio-input').fill('Кузнецов Дмитрий Павлович');
    
    // Заполняем первого ребенка
    const mChild1 = mobilePage.locator('.child-card').first();
    await mChild1.locator('.child-name-input').fill('Артём');
    await mobilePage.evaluate(() => {
        const d1 = document.querySelectorAll('.child-dob-input')[0];
        if (d1) {
            d1.type = 'date';
            d1.value = '2019-04-10';
        }
    });

    // Добавляем второго ребенка
    await mobilePage.locator('#btn-add-child').click();
    await mobilePage.waitForTimeout(300);
    const mChild2 = mobilePage.locator('.child-card').nth(1);
    await mChild2.locator('.child-name-input').fill('Полина');
    await mobilePage.evaluate(() => {
        const d2 = document.querySelectorAll('.child-dob-input')[1];
        if (d2) {
            d2.type = 'date';
            d2.value = '2021-08-22';
        }
    });

    await addAnnotation(mobilePage, {
        locator: mobilePage.locator('#btn-step2'),
        number: '2',
        text: '👉 Гость нажимает «Зарегистрироваться»',
        position: 'top',
        color: '#ff5a28'
    });
    await mobilePage.screenshot({ path: path.join(outputDir, 'qr_2_guest_mobile_children.png') });
    await clearAnnotations(mobilePage);

    // Переходим на Шаг 3 (Экран успеха с QR)
    console.log("📸 [QR 1.3] Экран успеха на смартфоне гостя...");
    await mobilePage.locator('#btn-step2').click();
    await mobilePage.waitForTimeout(1200);

    const qrFrame = mobilePage.locator('.card-qr');
    await addAnnotation(mobilePage, {
        locator: qrFrame,
        number: '3',
        text: '🎉 Гость зарегистрирован и подходит к кассе!',
        position: 'top',
        color: '#22c55e'
    });
    await mobilePage.screenshot({ path: path.join(outputDir, 'qr_3_guest_mobile_success.png') });
    await clearAnnotations(mobilePage);

    await mobileContext.close();

    // =========================================================================
    // ЧАСТЬ 2: ДЕСКТОПНАЯ КАССОВАЯ ПАНЕЛЬ RFID (ПОЛНОРАЗМЕРНЫЕ СКРИНШОТЫ)
    // =========================================================================
    console.log("💻 Снятие скриншотов десктопной RFID-панели кассира...");
    const desktopPage = await browser.newPage();
    await desktopPage.setViewportSize({ width: 1280, height: 860 });

    const indexPath = path.join(projectRoot, 'index.html');
    await desktopPage.goto(`file://${indexPath}`);
    await desktopPage.waitForTimeout(600);

    // Включаем POS токены, чтобы кнопки были активными
    await desktopPage.evaluate(() => {
        window.posTokens = 5;
        // Добавляем семью Кузнецовых, которая только что зарегистрировалась по QR
        const kuznetsov = {
            id: 99999,
            fio: 'Кузнецов Дмитрий Павлович',
            phone: '+7 (926) 777-88-99',
            createdAt: '12.08.2026',
            expanded: true,
            isQrRegistered: true,
            children: [
                { id: 201, name: 'Артём', dob: '2019-04-10', uid: 'CH-881', status: 'Нет', rfid: null },
                { id: 202, name: 'Полина', dob: '2021-08-22', uid: 'CH-882', status: 'Нет', rfid: null }
            ]
        };
        accounts.unshift(kuznetsov);
        if (typeof renderAccounts === 'function') renderAccounts();
    });
    await desktopPage.waitForTimeout(400);

    // Скриншот QR 1.4: Кассир видит семью после QR регистрации
    console.log("📸 [QR 1.4] Появление семьи в панели кассира после QR...");
    const artemRow = desktopPage.locator('tr.child-list-row').filter({ hasText: 'Артём' }).first();
    const bindArtemBtn = artemRow.locator('button').filter({ hasText: /привязать/i }).first();
    
    await addAnnotation(desktopPage, {
        locator: bindArtemBtn,
        number: '4',
        text: '👉 Нажмите «Привязать» для выдачи браслетов',
        position: 'top-center',
        color: '#ff5a28'
    });
    await desktopPage.screenshot({ path: path.join(outputDir, 'qr_4_cashier_family_arrived.png'), clip: { x: 50, y: 140, width: 1180, height: 490 } });
    await clearAnnotations(desktopPage);

    // ==========================================
    // СЦЕНАРИЙ 1 (ЗАПАСНОЙ): РУЧНОЙ ВВОД НА КАССЕ
    // ==========================================
    console.log("📸 [Сценарий 1.1] Главный экран -> Клик Добавить аккаунт...");
    await addAnnotation(desktopPage, {
        locator: desktopPage.locator('#add-account-btn'),
        number: '1',
        text: '👉 Если у гостя нет телефона: нажмите «Добавить аккаунт»',
        position: 'bottom',
        color: '#ff5a28'
    });
    await desktopPage.screenshot({ path: path.join(outputDir, 's1_1_click_add_account.png'), clip: { x: 50, y: 30, width: 1180, height: 420 } });
    await clearAnnotations(desktopPage);

    console.log("📸 [Сценарий 1.2] Форма добавления аккаунта...");
    await desktopPage.locator('#add-account-btn').click();
    await desktopPage.waitForTimeout(400);

    await desktopPage.locator('#new-account-fio').fill('Иванов Иван Иванович');
    await desktopPage.locator('#new-account-phone').click();
    await desktopPage.locator('#new-account-phone').fill('(999) 123-45-67');
    await desktopPage.waitForTimeout(300);

    await desktopPage.locator('#send-sms-btn').click();
    await desktopPage.waitForTimeout(300);

    await desktopPage.locator('#sms-code-input').fill('1234');
    await desktopPage.waitForTimeout(200);
    await desktopPage.locator('#verify-sms-btn').click();
    await desktopPage.waitForTimeout(400);

    await clearAnnotations(desktopPage);

    const child1 = desktopPage.locator('.child-input-group').first();
    await child1.locator('.child-name').fill('Максим');
    await child1.locator('.child-dob').fill('2018-05-12');

    await desktopPage.locator('#add-child-field-btn').click();
    await desktopPage.waitForTimeout(200);
    const child2 = desktopPage.locator('.child-input-group').nth(1);
    await child2.locator('.child-name').fill('Алиса');
    await child2.locator('.child-dob').fill('2020-09-03');

    await addAnnotation(desktopPage, {
        locator: desktopPage.locator('#submit-add-account'),
        number: '2',
        text: '👉 Нажмите «Создать аккаунт»',
        position: 'top',
        color: '#ff5a28'
    });
    await desktopPage.locator('#add-account-modal-content').screenshot({ path: path.join(outputDir, 's1_2_fill_form.png') });
    await clearAnnotations(desktopPage);

    console.log("📸 [Сценарий 1.3] Модалка привязки 1-го ребенка...");
    await desktopPage.evaluate(() => {
        const modal = document.getElementById('add-account-modal');
        modal.classList.add('hidden');
        
        const ivanovParent = {
            id: 99991,
            fio: 'Иванов Иван Иванович',
            phone: '+7 (999) 123-45-67',
            createdAt: '12.08.2026',
            expanded: true,
            children: [
                { id: 101, name: 'Максим', dob: '2018-05-12', uid: 'CH-991', status: 'Нет', rfid: null },
                { id: 102, name: 'Алиса', dob: '2020-09-03', uid: 'CH-992', status: 'Нет', rfid: null }
            ]
        };
        
        accounts.unshift(ivanovParent);
        renderAccounts();
        startChainBinding(99991);
    });
    await desktopPage.waitForTimeout(600);

    await addAnnotation(desktopPage, {
        locator: desktopPage.locator('#rfid-code-input'),
        number: '3',
        text: '👉 Приложите 1-й браслет к сканеру на кассе',
        position: 'top',
        color: '#5123d4'
    });
    await desktopPage.locator('#rfid-modal-content').screenshot({ path: path.join(outputDir, 's1_3_bind_first_child.png') });
    await clearAnnotations(desktopPage);

    console.log("📸 [Сценарий 1.4] Оверлей успеха 1-го ребенка...");
    await desktopPage.locator('#rfid-generate-btn').click();
    await desktopPage.waitForTimeout(150);
    await desktopPage.locator('#rfid-confirm-btn').click();
    await desktopPage.waitForTimeout(350);

    await addAnnotation(desktopPage, {
        locator: desktopPage.locator('#rfid-success-overlay'),
        number: '✓',
        text: 'Браслет привязан! Автопереход через 1.5 сек...',
        position: 'top',
        color: '#22c55e'
    });
    await desktopPage.locator('#rfid-modal-content').screenshot({ path: path.join(outputDir, 's1_4_success_first_child.png') });
    await clearAnnotations(desktopPage);

    console.log("📸 [Сценарий 1.5] Автопереход ко 2-му ребенку...");
    await desktopPage.waitForTimeout(1600);

    await addAnnotation(desktopPage, {
        locator: desktopPage.locator('#rfid-code-input'),
        number: '4',
        text: '👉 Приложите 2-й браслет к сканеру',
        position: 'top',
        color: '#5123d4'
    });
    await desktopPage.locator('#rfid-modal-content').screenshot({ path: path.join(outputDir, 's1_5_bind_second_child.png') });
    await clearAnnotations(desktopPage);

    await desktopPage.locator('#rfid-generate-btn').click();
    await desktopPage.waitForTimeout(150);
    await desktopPage.locator('#rfid-confirm-btn').click();
    await desktopPage.waitForTimeout(2200);

    console.log("📸 [Сценарий 1.6] Финальная таблица с привязанными браслетами...");
    await addAnnotation(desktopPage, {
        locator: desktopPage.locator('#accounts-table-body tr').first(),
        number: '5',
        text: '🎉 Браслеты привязаны! Выдайте их детям',
        position: 'bottom',
        color: '#22c55e'
    });
    await desktopPage.screenshot({ path: path.join(outputDir, 's1_6_done_table.png'), clip: { x: 50, y: 140, width: 1180, height: 490 } });
    await clearAnnotations(desktopPage);

    // ==========================================
    // СЦЕНАРИЙ 2: ПОВТОРНЫЙ ВИЗИТ (АВАТАР)
    // ==========================================
    console.log("📸 [Сценарий 2.1] Поиск родителя...");
    await desktopPage.locator('#search-input').fill('Смирнова');
    await desktopPage.waitForTimeout(400);

    await addAnnotation(desktopPage, {
        locator: desktopPage.locator('#search-input'),
        number: '1',
        text: '👉 Введите фамилию или телефон родителя',
        position: 'bottom',
        color: '#ff5a28'
    });
    await desktopPage.screenshot({ path: path.join(outputDir, 's2_1_search_parent.png'), clip: { x: 50, y: 140, width: 1180, height: 350 } });
    await clearAnnotations(desktopPage);

    console.log("📸 [Сценарий 2.2] Поиск Аватара ребенка...");
    const smirnovaRow = desktopPage.locator('#accounts-table-body tr').filter({ hasText: 'Смирнова Анна Юрьевна' }).first();
    const toggleBtn = smirnovaRow.locator('.toggle-children-btn');
    const toggleText = await toggleBtn.locator('span').innerText();
    if (toggleText.toLowerCase().includes('открыть') || toggleText.toLowerCase().includes('show')) {
        await toggleBtn.click();
        await desktopPage.waitForTimeout(400);
    }

    const egorRow = desktopPage.locator('tr.child-list-row').filter({ hasText: 'Егор' }).first();
    const avatarBadge = egorRow.locator('.status-btn, span').first();
    const newBraceletBtn = egorRow.locator('.child-bind-btn').first();

    await addAnnotation(desktopPage, {
        locator: avatarBadge,
        number: '2',
        text: 'У ребенка сохранен игровой профиль (Аватар)',
        position: 'top',
        color: '#5123d4'
    });
    await addAnnotation(desktopPage, {
        locator: newBraceletBtn,
        number: '3',
        text: '👉 Нажмите «Новый браслет»',
        position: 'bottom',
        color: '#ff5a28'
    });
    await desktopPage.screenshot({ path: path.join(outputDir, 's2_2_find_avatar.png'), clip: { x: 50, y: 180, width: 1180, height: 420 } });
    await clearAnnotations(desktopPage);

    console.log("📸 [Сценарий 2.3] Окно привязки нового браслета для Егора...");
    await newBraceletBtn.click();
    await desktopPage.waitForTimeout(400);

    await addAnnotation(desktopPage, {
        locator: desktopPage.locator('#rfid-code-input'),
        number: '4',
        text: '👉 Приложите новый чистый браслет к сканеру',
        position: 'top',
        color: '#5123d4'
    });
    await desktopPage.locator('#rfid-modal-content').screenshot({ path: path.join(outputDir, 's2_3_bind_modal.png') });
    await clearAnnotations(desktopPage);

    console.log("📸 [Сценарий 2.4] Тост успешной привязки повторного визита...");
    await desktopPage.locator('#rfid-generate-btn').click();
    await desktopPage.waitForTimeout(150);
    await desktopPage.locator('#rfid-confirm-btn').click();
    await desktopPage.waitForTimeout(300);

    await addAnnotation(desktopPage, {
        locator: desktopPage.locator('#toast-container'),
        number: '✓',
        text: 'Игровой прогресс и очки перенесены на новый браслет!',
        position: 'bottom',
        color: '#22c55e'
    });
    await desktopPage.screenshot({ path: path.join(outputDir, 's2_4_repeat_success.png'), clip: { x: 50, y: 30, width: 1180, height: 450 } });
    await clearAnnotations(desktopPage);

    console.log("📸 [Сценарий 2.5] История визитов в карточке Аватара...");
    await desktopPage.waitForTimeout(600);
    await clearAnnotations(desktopPage);
    await desktopPage.evaluate(() => {
        showAvatarInfo('Егор');
        const visitsContent = document.getElementById('visits-accordion-content');
        const visitsIcon = document.getElementById('visits-accordion-icon');
        if (visitsContent) {
            visitsContent.innerHTML = `
                <div class="flex items-center gap-3 py-1.5 border-b border-gray-100 dark:border-gray-700">
                    <div class="w-2 h-2 rounded-full bg-green-500"></div>
                    <span class="text-xs font-bold text-gray-800 dark:text-slate-200">12.08.2026, 14:30</span>
                    <span class="text-[10px] bg-green-50 text-green-600 px-2 py-0.5 rounded-full font-bold ml-auto">Сегодня</span>
                </div>
                <div class="flex items-center gap-3 py-1.5 border-b border-gray-100 dark:border-gray-700">
                    <div class="w-2 h-2 rounded-full bg-brand-orange"></div>
                    <span class="text-xs font-bold text-gray-700 dark:text-slate-300">05.08.2026, 11:15</span>
                </div>
                <div class="flex items-center gap-3 py-1.5">
                    <div class="w-2 h-2 rounded-full bg-gray-400"></div>
                    <span class="text-xs font-bold text-gray-700 dark:text-slate-300">20.07.2026, 16:45</span>
                </div>
            `;
            visitsContent.classList.remove('hidden');
        }
        if (visitsIcon) visitsIcon.classList.add('rotate-180');
        const countBadge = document.getElementById('visits-count-badge');
        if (countBadge) countBadge.textContent = '3';
    });
    await desktopPage.waitForTimeout(400);

    const visitsContent = desktopPage.locator('#visits-accordion-content');
    await addAnnotation(desktopPage, {
        locator: visitsContent,
        number: '5',
        text: 'Новый визит добавлен в историю посещений',
        position: 'top',
        color: '#5123d4'
    });
    await desktopPage.locator('#avatar-info-modal-content').screenshot({ path: path.join(outputDir, 's2_5_visit_history.png') });
    await clearAnnotations(desktopPage);

    await desktopPage.evaluate(() => {
        animateModalClose('avatar-info-modal');
    });
    await desktopPage.waitForTimeout(400);

    // ==========================================
    // СЦЕНАРИЙ 3: ВЫДАЧА ПРИЗОВ
    // ==========================================
    console.log("📸 [Сценарий 3.1] Кнопка Проверить браслет в шапке...");
    await addAnnotation(desktopPage, {
        locator: desktopPage.locator('#check-bracelet-global-btn'),
        number: '1',
        text: '👉 Нажмите «Проверить браслет»',
        position: 'bottom',
        color: '#ff5a28'
    });
    await desktopPage.screenshot({ path: path.join(outputDir, 's3_1_check_header_btn.png'), clip: { x: 50, y: 20, width: 1180, height: 160 } });
    await clearAnnotations(desktopPage);

    console.log("📸 [Сценарий 3.2] Модалка проверки браслета...");
    await desktopPage.locator('#check-bracelet-global-btn').click();
    await desktopPage.waitForTimeout(400);

    await addAnnotation(desktopPage, {
        locator: desktopPage.locator('#check-bracelet-input'),
        number: '2',
        text: '👉 Приложите браслет ребенка к сканеру',
        position: 'top',
        color: '#ff5a28'
    });
    await desktopPage.locator('#check-bracelet-modal-content').screenshot({ path: path.join(outputDir, 's3_2_check_modal.png') });
    await clearAnnotations(desktopPage);

    console.log("📸 [Сценарий 3.3] Карточка наград Аватара...");
    await desktopPage.evaluate(() => {
        animateModalClose('check-bracelet-modal');
        showAvatarInfo('Егор');
    });
    await desktopPage.waitForTimeout(500);

    const issueBtnLocator = desktopPage.locator('#avatar-prizes-container div').filter({ hasText: 'Стикерпак' }).locator('button').first();
    await addAnnotation(desktopPage, {
        locator: issueBtnLocator,
        number: '3',
        text: '👉 Нажмите «ВЫДАТЬ» и вручите подарок!',
        position: 'top',
        color: '#5123d4'
    });
    await desktopPage.locator('#avatar-info-modal-content').screenshot({ path: path.join(outputDir, 's3_3_avatar_prizes.png') });
    await clearAnnotations(desktopPage);

    await desktopPage.evaluate(() => {
        animateModalClose('avatar-info-modal');
    });
    await desktopPage.waitForTimeout(400);

    // ==========================================
    // СЦЕНАРИЙ 4: ОЧИСТКА БРАСЛЕТОВ
    // ==========================================
    console.log("📸 [Сценарий 4.1] Кнопка Очистить браслет...");
    await addAnnotation(desktopPage, {
        locator: desktopPage.locator('#clear-bracelet-global-btn'),
        number: '1',
        text: '👉 Нажмите «Очистить браслет»',
        position: 'bottom',
        color: '#4b5563'
    });
    await desktopPage.screenshot({ path: path.join(outputDir, 's4_1_clear_header_btn.png'), clip: { x: 750, y: 20, width: 480, height: 160 } });
    await clearAnnotations(desktopPage);

    console.log("📸 [Сценарий 4.2] Окно потоковой очистки...");
    await desktopPage.locator('#clear-bracelet-global-btn').click();
    await desktopPage.waitForTimeout(400);

    await addAnnotation(desktopPage, {
        locator: desktopPage.locator('#clear-bracelet-input'),
        number: '2',
        text: '👉 Прикладывайте браслеты по очереди',
        position: 'top-center',
        color: '#4b5563'
    });
    await desktopPage.locator('#clear-bracelet-modal-content').screenshot({ path: path.join(outputDir, 's4_2_clear_stream_modal.png') });
    await clearAnnotations(desktopPage);

    await browser.close();
    console.log("🎉 Все скриншоты (мобильные + десктопные) успешно сгенерированы в docs/screenshots/cashier_guide/!");
}

capture().catch(err => {
    console.error("❌ Ошибка при снятии скриншотов:", err);
    process.exit(1);
});
