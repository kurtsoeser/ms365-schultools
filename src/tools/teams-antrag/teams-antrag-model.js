/**
 * Domänenmodell Teams-Antrag (local-first).
 */

export const STORAGE_KEY = 'ms365-teams-antrag-v1';

export const STATUS = {
    ENTWURF: 'entwurf',
    BEANTRAGT: 'beantragt',
    IN_PRUEFUNG: 'in_pruefung',
    GENEHMIGT: 'genehmigt',
    ANGELEGT: 'angelegt',
    ABGELEHNT: 'abgelehnt',
    ZURUECKGESTELLT: 'zurueckgestellt',
    ANLAGE_FEHLGESCHLAGEN: 'anlage_fehlgeschlagen'
};

export const STATUS_LABEL = {
    entwurf: 'Entwurf',
    beantragt: 'Beantragt',
    in_pruefung: 'In Prüfung',
    genehmigt: 'Genehmigt',
    angelegt: 'Angelegt',
    abgelehnt: 'Abgelehnt',
    zurueckgestellt: 'Zurückgestellt',
    anlage_fehlgeschlagen: 'Anlage fehlgeschlagen'
};

export const ART = {
    PROJEKT: 'projekt',
    AG: 'ag',
    FACHGRUPPE: 'fachgruppe',
    KLASSE: 'klasse',
    DIPLOM_VWA: 'diplom_vwa',
    SONSTIGES: 'sonstiges'
};

export const ART_LABEL = {
    projekt: 'Projekt',
    ag: 'Arbeitsgemeinschaft (AG)',
    fachgruppe: 'Fachgruppe',
    klasse: 'Klasse / Jahrgang',
    diplom_vwa: 'Diplomarbeit / VWA',
    sonstiges: 'Sonstiges'
};

/**
 * @returns {string}
 */
export function newId(prefix) {
    const p = String(prefix || 'ta').trim() || 'ta';
    try {
        if (typeof crypto !== 'undefined' && crypto.randomUUID) {
            return p + '-' + crypto.randomUUID().replace(/-/g, '').slice(0, 12);
        }
    } catch {
        /* ignore */
    }
    return p + '-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

/**
 * @param {unknown} raw
 * @returns {string[]}
 */
export function parseMailList(raw) {
    return String(raw || '')
        .split(/[\n;,]+/)
        .map((s) => s.trim().toLowerCase())
        .filter((s) => s && s.includes('@'));
}

/**
 * @param {unknown} draft
 */
export function normalizeAntrag(draft) {
    const d = draft && typeof draft === 'object' ? draft : {};
    const status = String(d.status || STATUS.ENTWURF).trim().toLowerCase();
    const art = String(d.art || ART.SONSTIGES).trim().toLowerCase();
    return {
        id: String(d.id || newId('ta')),
        title: String(d.title || '').trim().slice(0, 120),
        zweck: String(d.zweck || '').trim(),
        art: Object.prototype.hasOwnProperty.call(ART_LABEL, art) ? art : ART.SONSTIGES,
        namensvorschlag: String(d.namensvorschlag || '').trim().slice(0, 120),
        besitzerMails: parseMailList(d.besitzerMails).join('\n'),
        mitgliederMails: parseMailList(d.mitgliederMails).join('\n'),
        schuljahr: String(d.schuljahr || '').trim().slice(0, 20),
        status: Object.prototype.hasOwnProperty.call(STATUS_LABEL, status) ? status : STATUS.ENTWURF,
        antragstellerEmail: String(d.antragstellerEmail || '')
            .trim()
            .toLowerCase()
            .slice(0, 255),
        antragstellerName: String(d.antragstellerName || '').trim().slice(0, 120),
        finalDisplayName: String(d.finalDisplayName || '').trim().slice(0, 120),
        finalMailNickname: String(d.finalMailNickname || '')
            .trim()
            .toLowerCase()
            .replace(/[^a-z0-9-]/g, '')
            .slice(0, 64),
        teamId: String(d.teamId || '').trim(),
        teamWebUrl: String(d.teamWebUrl || '').trim(),
        ablehnungsgrund: String(d.ablehnungsgrund || '').trim(),
        notizIt: String(d.notizIt || '').trim(),
        quelle: String(d.quelle || 'ui').trim() || 'ui',
        createdAt: String(d.createdAt || new Date().toISOString()),
        updatedAt: String(d.updatedAt || new Date().toISOString()),
        genehmigtAt: String(d.genehmigtAt || ''),
        angelegtAt: String(d.angelegtAt || ''),
        lastError: String(d.lastError || '').trim()
    };
}

/**
 * @returns {{ antraege: object[], updatedAt: string }}
 */
export function emptyState() {
    return { antraege: [], updatedAt: new Date().toISOString() };
}

/**
 * @param {unknown} raw
 */
export function normalizeState(raw) {
    const r = raw && typeof raw === 'object' ? raw : {};
    const antraege = Array.isArray(r.antraege) ? r.antraege.map(normalizeAntrag) : [];
    return {
        antraege,
        updatedAt: String(r.updatedAt || new Date().toISOString())
    };
}
