// E2E check of the config-driven international form (Dubai, Oman, Lima, locked RU) and its settings preview.
// Run: python3 -m http.server 4173 --bind 127.0.0.1  (from the repo root), then: node scripts/test_international_form.mjs
import { chromium } from 'playwright';

const B = process.env.HP_FORM_URL || 'http://127.0.0.1:4173/guest-registration/';
const browser = await chromium.launch({ headless: true });
const out = [];
const assert = (v, m) => { if (!v) throw Error(m); };
try {
    const ctx = await browser.newContext({ viewport: { width: 390, height: 844 } });
    const page = await ctx.newPage();
    const errors = [];
    page.on('pageerror', (e) => errors.push(e.message));
    const step = async (n) => {
        await page.waitForFunction((n) => document.querySelector('#step-' + n).classList.contains('step-visible'), n);
        await page.waitForTimeout(450);
    };

    // Dubai: English by default, UAE code, Arabic available.
    await page.goto(B + 'international.html?park=dubai');
    assert(await page.locator('html').getAttribute('lang') === 'en', 'Dubai should open in English');
    assert((await page.locator('#country-code-text').textContent()) === '+971', 'Dubai should default to +971');
    assert(!(await page.locator('#park-pill').count()), 'Park name pill should be gone');
    const hero = await page.locator('#step-1 .card-hero-image').boundingBox();
    assert(hero.width >= 150, 'Rocky should be large: ' + hero.width);
    const box = await page.locator('#check-mandatory').boundingBox();
    const txt = await page.locator('#consent-text').boundingBox();
    assert(Math.abs((box.y + box.height / 2) - (txt.y + txt.height / 2)) < 2, 'Consent text should be centred on the checkbox');

    // Country picker: rounded list with search by name and by code.
    await page.click('#country-button');
    assert(await page.locator('#country-popover').isVisible(), 'Picker should open');
    const radius = await page.locator('#country-popover').evaluate((el) => getComputedStyle(el).borderRadius);
    out.push('picker radius: ' + radius + ', first rows: ' + (await page.locator('.country-option .cn').allTextContents()).slice(0, 3).join(' | '));
    await page.screenshot({ path: 'output/playwright/intl-picker.png' });
    await page.fill('#country-search', 'per');
    out.push('search "per": ' + (await page.locator('.country-option .cn').allTextContents()).join(' | '));
    await page.fill('#country-search', '+51');
    assert((await page.locator('.country-option .cn').first().textContent()) === 'Peru', 'Search by code should find Peru');
    await page.fill('#country-search', 'zzzz');
    assert(await page.locator('.country-empty').isVisible(), 'Empty search should say nothing found');
    await page.fill('#country-search', 'oma');
    await page.keyboard.press('Enter');
    assert((await page.locator('#country-code-text').textContent()) === '+968', 'Enter should pick Oman');
    assert(await page.locator('#country-popover').isHidden(), 'Picker should close after choosing');
    await page.click('#country-button');
    await page.click('.country-option[data-country="AE"]');
    assert((await page.locator('#country-code-text').textContent()) === '+971', 'Click should pick UAE');
    await page.click('#country-button');
    await page.mouse.click(5, 5);
    assert(await page.locator('#country-popover').isHidden(), 'Outside click should close the picker');
    assert(await page.locator('#lang-switch button').count() === 2, 'Dubai should offer two languages');
    await page.fill('#phone-input', '0501234567');
    await page.check('#check-mandatory');
    assert(await page.locator('#btn-step1').isEnabled(), 'UAE mobile number should be accepted');
    out.push('AE formatted: ' + await page.inputValue('#phone-input'));
    await page.screenshot({ path: 'output/playwright/intl-dubai-en.png' });
    await page.click('#lang-switch button[lang=ar]');
    assert(await page.locator('html').getAttribute('dir') === 'rtl', 'Arabic should switch to RTL');
    assert(await page.inputValue('#phone-input') !== '', 'Phone should survive a language switch');
    await page.screenshot({ path: 'output/playwright/intl-dubai-ar.png' });
    await page.click('#country-button');
    await page.fill('#country-search', 'عمان');
    out.push('AR search "عمان": ' + (await page.locator('.country-option .cn').allTextContents()).slice(0, 2).join(' | '));
    await page.screenshot({ path: 'output/playwright/intl-picker-ar.png' });
    await page.keyboard.press('Escape');
    await page.click('#btn-step1'); await step(2);
    await page.click('#btn-step2');
    out.push('AR error: ' + await page.locator('#fio-error').innerText());
    await page.fill('#fio-input', 'سارة أحمد');
    await page.fill('.child-name-input', 'عمر');
    await page.fill('.child-dob-input', '31022018');
    await page.click('#btn-step2');
    assert(await page.locator('.child-err').isVisible(), '31 February should be rejected');
    await page.fill('.child-dob-input', '01012099');
    await page.click('#btn-step2');
    assert(await page.locator('.child-err').isVisible(), 'Future date should be rejected');
    await page.fill('.child-dob-input', '');
    await page.locator('.child-dob-input').pressSequentially('12042018');
    assert(await page.inputValue('.child-dob-input') === '12/04/2018', 'Date should be masked as DD/MM/YYYY');
    await page.screenshot({ path: 'output/playwright/intl-dubai-ar-family.png' });
    await page.click('#btn-step2'); await step(3);
    const payload = await page.evaluate(() => window.HP_LAST_REGISTRATION);
    out.push('payload: ' + JSON.stringify(payload));
    assert(payload.phone === '+971501234567' && payload.lang === 'ar' && payload.parkId === 'dubai', 'Payload should carry E.164 phone, park and language');
    await page.screenshot({ path: 'output/playwright/intl-dubai-ar-done.png' });
    await page.goto(B + 'international.html?park=dubai');
    assert(await page.locator('html').getAttribute('lang') === 'ar', 'Chosen language should be remembered per park');

    // Birth date placeholder follows the form language, whatever the browser locale is.
    await page.goto(B + 'international.html?park=lima&lang=es');
    await page.fill('#phone-input', '912345678'); await page.check('#check-mandatory'); await page.click('#btn-step1'); await step(2);
    out.push('ES date placeholder: ' + await page.getAttribute('.child-dob-input', 'placeholder'));
    await page.click('#lang-switch button[lang=en]');
    out.push('EN date placeholder: ' + await page.getAttribute('.child-dob-input', 'placeholder'));
    await page.evaluate(() => localStorage.removeItem('hp-lang:lima'));
    await page.goto(B + 'international.html?park=seligerskaya');
    await page.fill('#phone-input', '9161234567'); await page.check('#check-mandatory'); await page.click('#btn-step1'); await step(2);
    await page.locator('.child-dob-input').pressSequentially('12042018');
    assert(await page.inputValue('.child-dob-input') === '12.04.2018', 'Russian date should use dots');
    await page.screenshot({ path: 'output/playwright/intl-dob-ru.png' });

    // Oman: English by default, Oman code, landlines rejected, pasted +971 switches country.
    await page.goto(B + 'international.html?park=oman&lang=en');
    assert((await page.locator('#country-code-text').textContent()) === '+968', 'Oman should default to +968');
    await page.check('#check-mandatory');
    await page.fill('#phone-input', '24123456');
    assert(await page.locator('#btn-step1').isDisabled(), 'Oman landline should be rejected');
    await page.locator('#phone-input').blur();
    out.push('OM error: ' + await page.locator('#phone-error').innerText());
    await page.fill('#phone-input', '92123456');
    assert(await page.locator('#btn-step1').isEnabled(), 'Oman mobile should be accepted');
    await page.fill('#phone-input', '+971 50 123 4567');
    out.push('paste +971 -> ' + await page.locator('#country-code-text').textContent() + ' / ' + await page.inputValue('#phone-input'));
    assert((await page.locator('#country-code-text').textContent()) === '+971', 'Pasted international number should switch country');

    // Lima: Spanish by default, Peru code, English available.
    await page.goto(B + 'international.html?park=lima');
    assert(await page.locator('html').getAttribute('lang') === 'es', 'Lima should open in Spanish');
    assert((await page.locator('#country-code-text').textContent()) === '+51', 'Lima should default to +51');
    await page.click('#country-button');
    out.push('ES first options: ' + (await page.locator('.country-option .cn').allTextContents()).slice(0, 3).join(' | '));
    await page.keyboard.press('Escape');
    await page.fill('#phone-input', '912345678');
    await page.check('#check-mandatory');
    assert(await page.locator('#btn-step1').isEnabled(), 'Peru mobile should be accepted');
    out.push('PE formatted: ' + await page.inputValue('#phone-input'));
    await page.screenshot({ path: 'output/playwright/intl-lima-es.png' });

    // Locked example: one language, +7 only.
    await page.goto(B + 'international.html?park=seligerskaya');
    assert(await page.locator('#lang-switch').isHidden(), 'Single-language park should hide the switch');
    assert(await page.locator('#country-button').isDisabled(), 'Locked park should disable the country picker');
    await page.check('#check-mandatory');
    await page.fill('#phone-input', '+44 7911 123456');
    assert(await page.locator('#btn-step1').isDisabled(), 'Foreign number should be rejected when the code is locked');
    await page.fill('#phone-input', '9161234567');
    assert(await page.locator('#btn-step1').isEnabled(), 'Russian mobile should be accepted');

    await page.goto(B + 'international.html?park=nowhere');
    out.push('unknown park: ' + await page.locator('.success-title').innerText());

    for (const [park, l] of [['dubai', 'ar'], ['lima', 'es'], ['oman', 'en']]) {
        for (const w of [320, 390]) {
            await page.setViewportSize({ width: w, height: 844 });
            await page.goto(B + 'international.html?park=' + park + '&lang=' + l);
            assert(!await page.evaluate(() => document.documentElement.scrollWidth > innerWidth), 'Horizontal overflow: ' + park + ' ' + w);
        }
    }

    // Settings preview reacts to config changes.
    const s = await ctx.newPage();
    s.on('pageerror', (e) => errors.push('settings: ' + e.message));
    await s.setViewportSize({ width: 1100, height: 900 });
    await s.goto(B + 'international-settings.html');
    await s.waitForTimeout(700);
    const f = s.frameLocator('#frame');
    assert(await f.locator('html').getAttribute('lang') === 'en', 'Preview should start with Dubai in English');
    await s.selectOption('#park', 'lima'); await s.waitForTimeout(300);
    assert(await f.locator('html').getAttribute('lang') === 'es', 'Preview should switch to Lima');
    await s.selectOption('#default-lang', 'en'); await s.waitForTimeout(300);
    assert(await f.locator('html').getAttribute('lang') === 'en', 'Preview should follow the default language');
    await s.uncheck('#selectable');
    await s.selectOption('#default-country', 'CL'); await s.waitForTimeout(300);
    assert((await f.locator('#country-code-text').textContent()) === '+56', 'Preview should follow the default country');
    assert(await f.locator('#country-button').isDisabled(), 'Preview should lock the country picker');

    // Consent text: custom text with a link, then back to default.
    await s.selectOption('#park', 'dubai'); await s.waitForTimeout(300);
    const arArea = s.locator('.consent-block').nth(1).locator('textarea');
    await arArea.click(); await s.waitForTimeout(300);
    assert(await f.locator('html').getAttribute('lang') === 'ar', 'Editing Arabic consent should show Arabic preview');
    const enBlock = s.locator('.consent-block').first();
    await enBlock.locator('textarea').fill('I agree to the visit rules.');
    await enBlock.locator('textarea').evaluate((el) => el.setSelectionRange(15, 26));
    await enBlock.locator('input[type=url]').fill('https://example.com/rules');
    await enBlock.locator('.btn-link').click(); await s.waitForTimeout(200);
    assert((await enBlock.locator('textarea').inputValue()) === 'I agree to the [visit rules](https://example.com/rules).', 'Selected words should become a link');
    await enBlock.locator('textarea').focus(); await s.waitForTimeout(300);
    assert((await f.locator('#consent-text a').getAttribute('href')) === 'https://example.com/rules', 'Preview should render the custom link');
    assert((await s.locator('#json').textContent()).includes('consentText'), 'Config should contain the custom consent text');
    await s.screenshot({ path: 'output/playwright/intl-settings-consent.png', fullPage: true });
    await enBlock.locator('.btn-reset').click(); await s.waitForTimeout(300);
    assert(!(await s.locator('#json').textContent()).includes('consentText'), 'Reset should drop the custom text');
    assert((await f.locator('#consent-text').innerText()).startsWith('I accept the park rules'), 'Preview should return to default text');
    await s.screenshot({ path: 'output/playwright/intl-settings.png' });

    assert(!errors.length, errors.join('; '));
    out.push('ALL CHECKS PASSED');
} finally {
    await browser.close();
}
console.log(out.join('\n'));

