// Текст согласия: обычный текст, ссылки пишутся как [текст ссылки](https://адрес).
// Используется формой и страницей настроек, чтобы оба показывали одно и то же.
window.HP_CONSENT = (() => {
    const SAFE_URL = /^(https?:\/\/|mailto:|tel:)/i;

    function localized(value, lang, fallbackLang) {
        if (!value || typeof value === 'string') return value || '';
        return value[lang] || value[fallbackLang] || Object.values(value)[0] || '';
    }

    // Текст по умолчанию из словаря языка со ссылками парка.
    function defaultText(park, lang) {
        const dict = window.HP_I18N[lang] || window.HP_I18N.en;
        const fallback = park.languages && park.languages.default;
        const urls = {
            rulesUrl: localized(park.legal && park.legal.rulesUrl, lang, fallback) || '#',
            privacyUrl: localized(park.legal && park.legal.privacyUrl, lang, fallback) || '#'
        };
        return (dict.consent || window.HP_I18N.en.consent).replace(/\{(\w+)\}/g, (_, k) => urls[k] || '#');
    }

    // Текст, который увидит гость: свой текст парка или текст по умолчанию.
    function text(park, lang) {
        const custom = park.consentText && park.consentText[lang];
        return custom && custom.trim() ? custom : defaultText(park, lang);
    }

    function escapeHtml(v) {
        return String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    function toHtml(source) {
        return escapeHtml(source).replace(/\[([^\]]+)\]\(([^)\s]+)\)/g, (match, label, url) =>
            SAFE_URL.test(url) ? '<a href="' + url + '" target="_blank" rel="noopener noreferrer">' + label + '</a>' : label);
    }

    return { defaultText, text, toHtml };
})();

