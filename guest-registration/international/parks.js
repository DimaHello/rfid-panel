// Настройки парков для международной формы регистрации.
// В рабочей версии этот объект отдаёт API глобального сервера (см. docs/Архитектура_международной_формы.md),
// а редактирует его админка. Здесь он лежит файлом, чтобы прототип работал без сервера.
//
// languages.available — языки, между которыми может переключаться гость (коды из international/i18n.js).
// languages.default   — язык, который гость видит при первом открытии.
// phone.selectable    — может ли гость сменить код страны.
// phone.defaultCountry — страна телефона по умолчанию (ISO 3166-1 alpha-2: AE, OM, PE, RU…).
// phone.priorityCountries — страны в начале списка, остальные идут по алфавиту.
// phone.allowedCountries — необязательно: ограничить список стран; если не задан, доступны все.
// legal.*Url — ссылка строкой или объектом по языкам: { en: '…', ar: '…' }.
// consentText — необязательно: свой текст согласия по языкам. Ссылки: [текст](https://адрес).
//   Если для языка текста нет, берётся текст из словаря со ссылками legal.*Url.

window.HP_DEFAULT_PARK = 'dubai';

window.HP_PARKS = {
    dubai: {
        id: 'dubai',
        adminName: 'Дубай (ОАЭ)',
        name: { en: 'Hello Park Dubai', ar: 'Hello Park دبي' },
        languages: { available: ['en', 'ar'], default: 'en' },
        phone: {
            selectable: true,
            defaultCountry: 'AE',
            priorityCountries: ['AE', 'OM', 'SA', 'QA', 'KW', 'BH', 'IN', 'PK', 'PH', 'GB', 'RU']
        },
        legal: { rulesUrl: 'https://hello-park.com', privacyUrl: 'https://hello-park.com' }
    },
    oman: {
        id: 'oman',
        adminName: 'Оман',
        name: { en: 'Hello Park Oman', ar: 'Hello Park عُمان' },
        languages: { available: ['en', 'ar'], default: 'en' },
        phone: {
            selectable: true,
            defaultCountry: 'OM',
            priorityCountries: ['OM', 'AE', 'SA', 'QA', 'KW', 'BH', 'IN', 'PK', 'GB']
        },
        legal: { rulesUrl: 'https://hello-park.com', privacyUrl: 'https://hello-park.com' }
    },
    lima: {
        id: 'lima',
        adminName: 'Лима (Перу)',
        name: { es: 'Hello Park Lima', en: 'Hello Park Lima' },
        languages: { available: ['es', 'en'], default: 'es' },
        phone: {
            selectable: true,
            defaultCountry: 'PE',
            priorityCountries: ['PE', 'CL', 'CO', 'EC', 'BO', 'AR', 'BR', 'VE', 'US', 'ES']
        },
        legal: { rulesUrl: 'https://hello-park.com', privacyUrl: 'https://hello-park.com' }
    },
    seligerskaya: {
        id: 'seligerskaya',
        adminName: 'Селигерская (пример: один язык, код закреплён)',
        name: { ru: 'Hello Park Селигерская' },
        languages: { available: ['ru'], default: 'ru' },
        phone: { selectable: false, defaultCountry: 'RU', priorityCountries: [] },
        legal: {
            rulesUrl: 'https://hello-park.ru/seligerskaya/visiting',
            privacyUrl: 'https://hello-park.ru/seligerskaya/policy'
        }
    }
};

