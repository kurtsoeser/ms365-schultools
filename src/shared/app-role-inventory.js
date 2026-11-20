/**
 * Lesende IT-Sicht: was in diesem Browser für jede Schul-App eingerichtet ist.
 * Schreibt nichts zurück. Die App-Rolle selbst entsteht erst beim Öffnen der App.
 */

/** @typedef {{ getItem: (key: string) => string|null }} StorageLike */

/**
 * @typedef {{
 *   id: string,
 *   title: string,
 *   family: string,
 *   href: string,
 *   keys: string[],
 *   note?: string
 * }} AppPermissionSlot
 */

/** @type {AppPermissionSlot[]} */
export const APP_PERMISSION_SLOTS = [
    { id: 'schulaktivitaeten', title: 'Schulaktivitäten', family: 'Genehmigung', href: 'schulaktivitaeten-planer.html', keys: ['ms365-schulaktivitaeten-perms-v1', 'ms365-akt-setup-v1'] },
    { id: 'freistellung', title: 'Freistellungen Schüler', family: 'Genehmigung', href: 'freistellung-planer.html', keys: ['ms365-freistellung-perms-v1', 'ms365-freistellung-setup-v1'] },
    { id: 'lehrer-freistellung', title: 'Freistellungen Lehrkräfte', family: 'Genehmigung', href: 'lehrer-freistellung-planer.html', keys: ['ms365-lfr-perms-v1', 'ms365-lfr-setup-v1'] },
    { id: 'schularbeiten', title: 'Schularbeiten', family: 'Register', href: 'schularbeiten-planer.html', keys: ['ms365-schularbeiten-perms-v1'] },
    { id: 'projektwochen', title: 'Projektwochen', family: 'Buchung', href: 'projektwochen.html', keys: ['ms365-schularbeiten-perms-v1'], note: 'Liest die Schularbeiten-Gruppen mit.' },
    { id: 'srdp', title: 'sRDP-Anmeldung', family: 'Wahl', href: 'srdp-anmeldung.html', keys: ['ms365-srdp-list-perms-v1', 'ms365-srdp-anmeldung-v1'] },
    { id: 'wahlfach', title: 'Wahlfach-Zuteilung', family: 'Wahl', href: 'wahlfach-zuteilung.html', keys: ['ms365-schularbeiten-perms-v1'], note: 'Zugang über die Schularbeiten-Admin-Gruppe.' },
    { id: 'einsammler', title: 'Einsammler', family: 'Ablauf', href: 'einsammler.html', keys: ['ms365-einsammler-settings-v1', 'ms365-einsammler-sp-settings-v1'] },
    { id: 'elternbrief', title: 'Elternbrief', family: 'Ablauf', href: 'elternbrief.html', keys: ['ms365-elternbrief-settings-v1', 'ms365-elternbrief-sp-settings-v1'] },
    { id: 'krankmeldung', title: 'Krankmeldung', family: 'Ablauf', href: 'krankmeldung.html', keys: ['ms365-krankmeldung-settings-v1', 'ms365-krankmeldung-sp-settings-v1'] },
    { id: 'kollegium', title: 'Kollegiumsumfragen', family: 'Ablauf', href: 'kollegium-umfragen.html', keys: ['ms365-kollegium-umfragen-settings-v1', 'ms365-kollegium-umfragen-sp-settings-v1'] },
    { id: 'ausgaben', title: 'Ausgaben', family: 'Betrieb', href: 'ausgaben.html', keys: ['ms365-ausgaben-settings-v1', 'ms365-ausgaben-sp-settings-v1'] },
    { id: 'aufgaben', title: 'Aufgaben', family: 'Betrieb', href: 'aufgaben.html', keys: ['ms365-aufgaben-sp-settings-v1'] },
    { id: 'technik', title: 'Technik-Meldungen', family: 'Betrieb', href: 'technik-meldungen.html', keys: ['ms365-technik-meldungen-sp-settings-v1'] },
    { id: 'raumplan', title: 'Raumplan', family: 'Betrieb', href: 'raumplan.html', keys: ['ms365-raumplan-sp-settings-v1'] },
    { id: 'termine', title: 'Kalender und Termine', family: 'Kalender', href: 'termine-hub.html', keys: ['ms365-termine-hub-settings-v1'] },
    { id: 'teams-antrag', title: 'Team-Antrag', family: 'Antrag', href: 'teams-antrag.html', keys: ['ms365-teams-antrag-settings-v1'], note: 'Freigeber-Gruppe und Freigeber-Mails.' },
    { id: 'schulsprecherwahl', title: 'Schulsprecherwahl', family: 'Wahl', href: 'schulsprecherwahl.html', keys: [], note: 'Wahlleitung kommt vom Wahl-Dienst, nicht aus einer Planer-Matrix.' },
    { id: 'elternsprechtag', title: 'Elternsprechtag', family: 'Buchung', href: 'elternsprechtag-bookings.html', keys: [], note: 'Buchungskonto, keine Genehmiger-Rolle.' },
    { id: 'diplomarbeiten', title: 'Diplomarbeiten', family: 'Unterricht', href: 'diplomarbeiten.html', keys: [], note: 'Betriebsmodus, keine Rollenmatrix.' }
];

const KEY_LABELS = {
    groupLehrerName: 'Lehrkräfte-Gruppe',
    groupSchuelerName: 'Schüler-Gruppe',
    groupDirektionName: 'Direktion-Gruppe',
    groupAdminName: 'Verwaltung-Gruppe',
    groupKvName: 'KV-Gruppe',
    emailDirektion: 'Genehmiger',
    adminEmails: 'Admin-Mails',
    lehrerUsers: 'Lehrkräfte (Personen)',
    direktionUsers: 'Direktion (Personen)',
    adminUsers: 'Verwaltung (Personen)',
    schuelerUsers: 'Schüler (Personen)'
};

/**
 * @param {unknown} raw
 * @returns {string[]}
 */
export function summarizePermissionRecord(raw) {
    const o = raw && typeof raw === 'object' ? /** @type {Record<string, unknown>} */ (raw) : {};
    /** @type {string[]} */
    const lines = [];
    Object.keys(o).forEach((key) => {
        const value = o[key];
        const label = KEY_LABELS[key] || humanKey(key);
        if (typeof value === 'string') {
            const text = value.trim();
            if (!text) return;
            if (/id$/i.test(key) && !text.includes('@')) return;
            if (/name|email|mail/i.test(key) || text.includes('@')) lines.push(label + ': ' + text);
            return;
        }
        if (!Array.isArray(value) || !value.length) return;
        const people = value
            .map((item) => {
                if (typeof item === 'string') return item.trim();
                if (!item || typeof item !== 'object') return '';
                const row = /** @type {Record<string, unknown>} */ (item);
                return String(row.displayName || row.name || row.mail || row.email || '').trim();
            })
            .filter(Boolean);
        if (people.length) lines.push(label + ': ' + people.join(', '));
    });
    return lines;
}

/**
 * @param {string} key
 */
function humanKey(key) {
    return String(key || '')
        .replace(/([A-Z])/g, ' $1')
        .replace(/[_-]+/g, ' ')
        .trim();
}

/**
 * @param {StorageLike} storage
 * @param {string[]} keys
 */
export function readSlotLines(storage, keys) {
    /** @type {string[]} */
    const lines = [];
    (keys || []).forEach((key) => {
        let parsed = null;
        try {
            const raw = storage.getItem(key);
            if (!raw) return;
            parsed = JSON.parse(raw);
        } catch {
            return;
        }
        summarizePermissionRecord(parsed).forEach((line) => {
            if (!lines.includes(line)) lines.push(line);
        });
    });
    return lines;
}

/**
 * @param {StorageLike} storage
 */
export function readAppRoleInventory(storage) {
    let audience = {};
    try {
        audience = JSON.parse(storage.getItem('ms365-dashboard-audience-groups-v1') || '{}') || {};
    } catch {
        audience = {};
    }
    const principals = summarizePermissionRecord(audience);
    const apps = APP_PERMISSION_SLOTS.map((slot) => {
        const lines = readSlotLines(storage, slot.keys);
        return {
            id: slot.id,
            title: slot.title,
            family: slot.family,
            href: slot.href,
            note: slot.note || '',
            lines,
            configured: lines.length > 0
        };
    });
    return { principals, apps };
}
