/**
 * Domänenmodell Raumplan (Stammliste + Plan-Viewer).
 */

export const STORAGE_KEY = 'ms365-raumplan-v1';
export const DEMO_SEED_TAG = 'raumplan-demo-v1';

export const RAUM_TYP = {
    KLASSE: 'klasse',
    FACHRAUM: 'fachraum',
    BUERO: 'buero',
    SPORTHALLE: 'sporthalle',
    SONSTIGES: 'sonstiges'
};

export const RAUM_TYP_LABEL = {
    klasse: 'Klasse',
    fachraum: 'Fachraum',
    buero: 'Büro',
    sporthalle: 'Sporthalle',
    sonstiges: 'Sonstiges'
};

/**
 * @returns {string}
 */
export function newId(prefix) {
    const p = String(prefix || 'rp').trim() || 'rp';
    try {
        if (typeof crypto !== 'undefined' && crypto.randomUUID) {
            return p + '-' + crypto.randomUUID().replace(/-/g, '').slice(0, 12);
        }
    } catch {
        /* ignore */
    }
    return p + '-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

/**
 * @param {unknown} raw
 */
export function normalizeRaum(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const typRaw = String(o.typ || o.Typ || RAUM_TYP.SONSTIGES)
        .trim()
        .toLowerCase();
    const kap = Number(o.kapazitaet ?? o.Kapazitaet);
    return {
        id: String(o.id || o.RaumId || '').trim() || newId('raum'),
        code: String(o.code || o.Code || '')
            .trim()
            .toUpperCase(),
        name: String(o.name || o.Name || o.Title || '').trim(),
        gebaeude: String(o.gebaeude || o.Gebaeude || '').trim(),
        stockwerk: String(o.stockwerk || o.Stockwerk || '').trim(),
        kapazitaet: Number.isFinite(kap) && kap >= 0 ? Math.round(kap) : 0,
        typ: Object.values(RAUM_TYP).includes(typRaw) ? typRaw : RAUM_TYP.SONSTIGES,
        notiz: String(o.notiz || o.Notiz || '').trim(),
        seedTag: String(o.seedTag || o.SeedTag || '').trim(),
        spoItemId: String(o.spoItemId || o.itemId || '').trim()
    };
}

/**
 * @param {unknown} item
 */
export function isDemoSeedItem(item) {
    return String((item && item.seedTag) || '') === DEMO_SEED_TAG;
}

export function emptyState() {
    return {
        raeume: [],
        planImageDataUrl: '',
        planPdfName: '',
        localDemoOnly: false,
        updatedAt: new Date().toISOString()
    };
}

/**
 * @param {unknown} raw
 */
export function normalizeState(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    return {
        raeume: Array.isArray(o.raeume) ? o.raeume.map(normalizeRaum) : [],
        planImageDataUrl: String(o.planImageDataUrl || '').trim(),
        planPdfName: String(o.planPdfName || '').trim(),
        localDemoOnly: !!o.localDemoOnly,
        updatedAt: String(o.updatedAt || '').trim() || new Date().toISOString()
    };
}
