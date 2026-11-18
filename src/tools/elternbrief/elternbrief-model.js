/**
 * Domänenmodell Elternbrief (eigenes Tool, nicht Einsammler-Modus).
 */

export const STORAGE_KEY = 'ms365-elternbrief-v1';

/** Seed-Tag für Demo-Zeilen (nicht nach SharePoint pushen). */
export const ELTERNBRIEF_DEMO_SEED_TAG = 'elternbrief-demo-v1';

export const STATUS = {
    ENTWURF: 'entwurf',
    BEREIT: 'bereit',
    GESENDET: 'gesendet',
    GESCHLOSSEN: 'geschlossen',
    ARCHIVIERT: 'archiviert'
};

export const STATUS_LABEL = {
    entwurf: 'Entwurf',
    bereit: 'Bereit zum Versand',
    gesendet: 'Gesendet / Links aktiv',
    geschlossen: 'Geschlossen',
    archiviert: 'Archiviert'
};

export const ABSENDER_MODUS = {
    EIGEN: 'eigenes_konto',
    ZENTRAL: 'zentrales_postfach'
};

export const ABSENDER_MODUS_LABEL = {
    eigenes_konto: 'Eigenes Konto (Lehrkraft)',
    zentrales_postfach: 'Zentrales Elternbrief-Postfach'
};

export const LESE_STATUS = {
    OFFEN: 'offen',
    GELESEN: 'gelesen',
    ABGELAUFEN: 'abgelaufen'
};

export const LESE_STATUS_LABEL = {
    offen: 'Offen',
    gelesen: 'Gelesen',
    abgelaufen: 'Abgelaufen'
};

export const UPLOAD_MAX_BYTES = 6 * 1024 * 1024;

/**
 * @param {unknown} value
 * @returns {string}
 */
export function toIsoDateOnly(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
        return value.toISOString().slice(0, 10);
    }
    const s = String(value || '').trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
    return '';
}

/**
 * @returns {string}
 */
export function newId(prefix) {
    const p = String(prefix || 'eb').trim() || 'eb';
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
export function normalizeBrief(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    return {
        id: String(o.id || o.BriefId || '').trim() || newId('eb'),
        betreff: String(o.betreff || o.Betreff || o.titel || o.Title || '').trim(),
        inhalt: String(o.inhalt || o.InhaltHtml || o.textInhalt || '').trim(),
        vorlageId: String(o.vorlageId || o.VorlageId || '').trim(),
        klasseCodes: String(o.klasseCodes || o.KlasseCodes || '')
            .trim()
            .replace(/,/g, ';'),
        absenderModus: String(o.absenderModus || o.AbsenderModus || ABSENDER_MODUS.EIGEN)
            .trim()
            .toLowerCase(),
        absenderEmail: String(o.absenderEmail || o.AbsenderEmail || '')
            .trim()
            .toLowerCase(),
        absenderName: String(o.absenderName || o.AbsenderName || '').trim(),
        anhangName: String(o.anhangName || o.AnhangName || '').trim(),
        anhangMime: String(o.anhangMime || o.AnhangMime || '').trim(),
        anhangSize: Number(o.anhangSize || 0) || 0,
        anhangKey: String(o.anhangKey || '').trim(),
        anhangUrl: String(o.anhangUrl || o.AnhangUrl || '').trim(),
        faelligAm: toIsoDateOnly(o.faelligAm || o.FaelligAm),
        lesebestaetigung: o.lesebestaetigung !== false && o.Lesebestaetigung !== false,
        status: String(o.status || o.Status || STATUS.ENTWURF)
            .trim()
            .toLowerCase(),
        erstellerEmail: String(o.erstellerEmail || o.ErstellerEmail || '')
            .trim()
            .toLowerCase(),
        erstellerName: String(o.erstellerName || o.ErstellerName || '').trim(),
        gesendetAm: String(o.gesendetAm || o.GesendetAm || '').trim(),
        quelle: String(o.quelle || o.Quelle || 'ui')
            .trim()
            .toLowerCase(),
        seedTag: String(o.seedTag || o.SeedTag || '').trim(),
        spoItemId: String(o.spoItemId || o.itemId || '').trim(),
        createdAt: String(o.createdAt || '').trim() || new Date().toISOString()
    };
}

/**
 * @param {unknown} raw
 */
export function normalizeEmpfaenger(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    return {
        id: String(o.id || o.LeseId || '').trim() || newId('el'),
        briefId: String(o.briefId || o.BriefId || '').trim(),
        schuelerName: String(o.schuelerName || o.SchuelerName || '').trim(),
        klasseCode: String(o.klasseCode || o.KlasseCode || '').trim(),
        elternEmail: String(o.elternEmail || o.ElternEmail || '')
            .trim()
            .toLowerCase(),
        status: String(o.status || o.Status || LESE_STATUS.OFFEN)
            .trim()
            .toLowerCase(),
        gelesenAm: String(o.gelesenAm || o.GelesenAm || '').trim(),
        quelle: String(o.quelle || o.Quelle || 'ui')
            .trim()
            .toLowerCase(),
        seedTag: String(o.seedTag || o.SeedTag || '').trim(),
        spoItemId: String(o.spoItemId || o.itemId || '').trim()
    };
}

/**
 * @param {unknown} item
 */
export function isDemoSeedItem(item) {
    if (!item || typeof item !== 'object') return false;
    if (String(item.quelle || '').toLowerCase() === 'demo') return true;
    return String(item.seedTag || '') === ELTERNBRIEF_DEMO_SEED_TAG;
}

export function emptyState() {
    return {
        briefe: [],
        empfaenger: [],
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
        briefe: Array.isArray(o.briefe) ? o.briefe.map(normalizeBrief) : [],
        empfaenger: Array.isArray(o.empfaenger) ? o.empfaenger.map(normalizeEmpfaenger) : [],
        localDemoOnly: !!o.localDemoOnly,
        updatedAt: String(o.updatedAt || '').trim() || new Date().toISOString()
    };
}
