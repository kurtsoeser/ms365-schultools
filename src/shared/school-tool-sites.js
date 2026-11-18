/**
 * Schulweite Planer-/Tools-Sites und zentrale Postfächer (Stammdaten → Tools).
 * Auflösung: Tool-localStorage (Override) → setup.planerSiteUrls[kind] → planerDefaultSiteUrl → Intranet.
 */

/** @typedef {'schularbeiten'|'projektwochen'|'freistellung'|'lehrerFreistellung'|'aktivitaeten'|'wahlfach'|'srdp'|'krankmeldung'|'einsammler'|'elternbrief'|'technik'|'aufgaben'|'ausgaben'|'raumplan'|'kollegium'|'stammdatenPortal'} SchoolToolSiteKind */

/** @type {Record<SchoolToolSiteKind, string>} */
export const SCHOOL_TOOL_SITE_KEYS = {
    schularbeiten: 'ms365-sa-site-url',
    projektwochen: 'ms365-pw-site-url',
    freistellung: 'ms365-freistellung-planer-site-v1',
    lehrerFreistellung: 'ms365-lfr-planer-site-v1',
    aktivitaeten: 'ms365-akt-planer-site-v1',
    wahlfach: '', // in ms365-wahlfach-zuteilung-settings-v1.siteUrl
    srdp: 'ms365-srdp-planer-site-v1',
    krankmeldung: '', // in ms365-krankmeldung-sp-settings-v1.siteUrl
    einsammler: '', // in ms365-einsammler-sp-settings-v1.siteUrl
    elternbrief: '', // in ms365-elternbrief-sp-settings-v1.siteUrl
    technik: '', // in ms365-technik-meldungen-sp-settings-v1.siteUrl
    aufgaben: '', // in ms365-aufgaben-sp-settings-v1.siteUrl
    ausgaben: '', // in ms365-ausgaben-sp-settings-v1.siteUrl
    raumplan: '', // in ms365-raumplan-sp-settings-v1.siteUrl
    kollegium: '', // in ms365-kollegium-umfragen-sp-settings-v1.siteUrl
    stammdatenPortal: '' // in ms365-schueler-stammdaten-portal-v1.spSiteUrl (Gateway schreibt App-Only)
};

/** @type {{ id: SchoolToolSiteKind, label: string }[]} */
export const SCHOOL_TOOL_SITE_KINDS = [
    { id: 'schularbeiten', label: 'Schularbeiten-Planer' },
    { id: 'projektwochen', label: 'Projektwochen' },
    { id: 'freistellung', label: 'Freistellungen Schüler' },
    { id: 'lehrerFreistellung', label: 'Freistellungen Lehrkräfte' },
    { id: 'aktivitaeten', label: 'Schulaktivitäten' },
    { id: 'wahlfach', label: 'Wahlfach-Zuteilung' },
    { id: 'srdp', label: 'sRDP-Anmeldung' },
    { id: 'krankmeldung', label: 'Krankmeldung' },
    { id: 'einsammler', label: 'Einsammler' },
    { id: 'elternbrief', label: 'Elternbrief' },
    { id: 'technik', label: 'Technik-Meldungen' },
    { id: 'aufgaben', label: 'Aufgaben' },
    { id: 'ausgaben', label: 'Ausgabenverwaltung' },
    { id: 'raumplan', label: 'Raumplan' },
    { id: 'kollegium', label: 'Kollegiumsumfragen' },
    { id: 'stammdatenPortal', label: 'Schüler-Stammdaten-Portal' }
];

const WAHLFACH_SETTINGS_KEY = 'ms365-wahlfach-zuteilung-settings-v1';
const KRANKMELDUNG_SP_SETTINGS_KEY = 'ms365-krankmeldung-sp-settings-v1';
const EINSAMMLER_SP_SETTINGS_KEY = 'ms365-einsammler-sp-settings-v1';
const ELTERNBRIEF_SP_SETTINGS_KEY = 'ms365-elternbrief-sp-settings-v1';
const TECHNIK_SP_SETTINGS_KEY = 'ms365-technik-meldungen-sp-settings-v1';
const AUFGABEN_SP_SETTINGS_KEY = 'ms365-aufgaben-sp-settings-v1';
const AUSGABEN_SP_SETTINGS_KEY = 'ms365-ausgaben-sp-settings-v1';
const RAUMPLAN_SP_SETTINGS_KEY = 'ms365-raumplan-sp-settings-v1';
const KOLLEGIUM_SP_SETTINGS_KEY = 'ms365-kollegium-umfragen-sp-settings-v1';
const ELTERNBRIEF_SETTINGS_KEY = 'ms365-elternbrief-settings-v1';
const STAMMDATEN_PORTAL_SETTINGS_KEY = 'ms365-schueler-stammdaten-portal-v1';

function trimUrl(u) {
    return String(u || '').trim();
}

function readSetup(setup) {
    if (setup && typeof setup === 'object') return setup;
    try {
        if (
            typeof window !== 'undefined' &&
            window.ms365AppDataV2 &&
            typeof window.ms365AppDataV2.getSetup === 'function'
        ) {
            return window.ms365AppDataV2.getSetup() || {};
        }
    } catch {
        /* ignore */
    }
    return {};
}

function localStorageTrim(key) {
    if (!key) return '';
    try {
        return String(localStorage.getItem(key) || '').trim();
    } catch {
        return '';
    }
}

function readWahlfachSiteFromStorage() {
    try {
        const raw = localStorage.getItem(WAHLFACH_SETTINGS_KEY);
        if (!raw) return '';
        const o = JSON.parse(raw);
        return o && o.siteUrl ? String(o.siteUrl).trim() : '';
    } catch {
        return '';
    }
}

function readKrankmeldungSiteFromStorage() {
    try {
        const raw = localStorage.getItem(KRANKMELDUNG_SP_SETTINGS_KEY);
        if (!raw) return '';
        const o = JSON.parse(raw);
        return o && o.siteUrl ? String(o.siteUrl).trim() : '';
    } catch {
        return '';
    }
}

function readEinsammlerSiteFromStorage() {
    try {
        const raw = localStorage.getItem(EINSAMMLER_SP_SETTINGS_KEY);
        if (!raw) return '';
        const o = JSON.parse(raw);
        return o && o.siteUrl ? String(o.siteUrl).trim() : '';
    } catch {
        return '';
    }
}

function readElternbriefSiteFromStorage() {
    try {
        const raw = localStorage.getItem(ELTERNBRIEF_SP_SETTINGS_KEY);
        if (!raw) return '';
        const o = JSON.parse(raw);
        return o && o.siteUrl ? String(o.siteUrl).trim() : '';
    } catch {
        return '';
    }
}

function readTechnikSiteFromStorage() {
    try {
        const raw = localStorage.getItem(TECHNIK_SP_SETTINGS_KEY);
        if (!raw) return '';
        const o = JSON.parse(raw);
        return o && o.siteUrl ? String(o.siteUrl).trim() : '';
    } catch {
        return '';
    }
}

function readAufgabenSiteFromStorage() {
    try {
        const raw = localStorage.getItem(AUFGABEN_SP_SETTINGS_KEY);
        if (!raw) return '';
        const o = JSON.parse(raw);
        return o && o.siteUrl ? String(o.siteUrl).trim() : '';
    } catch {
        return '';
    }
}

function readKollegiumSiteFromStorage() {
    try {
        const raw = localStorage.getItem(KOLLEGIUM_SP_SETTINGS_KEY);
        if (!raw) return '';
        const o = JSON.parse(raw);
        return o && o.siteUrl ? String(o.siteUrl).trim() : '';
    } catch {
        return '';
    }
}

function readAusgabenSiteFromStorage() {
    try {
        const raw = localStorage.getItem(AUSGABEN_SP_SETTINGS_KEY);
        if (!raw) return '';
        const o = JSON.parse(raw);
        return o && o.siteUrl ? String(o.siteUrl).trim() : '';
    } catch {
        return '';
    }
}

function readRaumplanSiteFromStorage() {
    try {
        const raw = localStorage.getItem(RAUMPLAN_SP_SETTINGS_KEY);
        if (!raw) return '';
        const o = JSON.parse(raw);
        return o && o.siteUrl ? String(o.siteUrl).trim() : '';
    } catch {
        return '';
    }
}

function readStammdatenPortalSiteFromStorage() {
    try {
        const raw = localStorage.getItem(STAMMDATEN_PORTAL_SETTINGS_KEY);
        if (!raw) return '';
        const o = JSON.parse(raw);
        return o && o.spSiteUrl ? String(o.spSiteUrl).trim() : '';
    } catch {
        return '';
    }
}

function writeStammdatenPortalSiteUrl(url) {
    if (!url) return;
    try {
        const raw = localStorage.getItem(STAMMDATEN_PORTAL_SETTINGS_KEY);
        const o = raw ? JSON.parse(raw) || {} : {};
        if (String(o.spSiteUrl || '').trim()) return;
        o.spSiteUrl = url;
        if (!o.spListName) o.spListName = 'SchuelerAnmeldungen';
        localStorage.setItem(STAMMDATEN_PORTAL_SETTINGS_KEY, JSON.stringify(o));
    } catch {
        /* ignore */
    }
}

function writeSettingsSiteUrl(settingsKey, url) {
    if (!settingsKey || !url) return;
    try {
        const raw = localStorage.getItem(settingsKey);
        const o = raw ? JSON.parse(raw) || {} : {};
        if (String(o.siteUrl || '').trim()) return;
        o.siteUrl = url;
        localStorage.setItem(settingsKey, JSON.stringify(o));
    } catch {
        /* ignore */
    }
}

/**
 * @param {object} [setup]
 * @returns {{ planerDefaultSiteUrl: string, planerSiteUrls: Record<string, string>, schoolMailboxes: { elternbriefZentral: string, elternbriefName: string, schulkalender: string } }}
 */
export function normalizeSchoolToolSitesSetup(setup) {
    const x = setup && typeof setup === 'object' ? setup : {};
    const urlsIn = x.planerSiteUrls && typeof x.planerSiteUrls === 'object' ? x.planerSiteUrls : {};
    /** @type {Record<string, string>} */
    const planerSiteUrls = {};
    SCHOOL_TOOL_SITE_KINDS.forEach(function (k) {
        planerSiteUrls[k.id] = trimUrl(urlsIn[k.id]);
    });
    const mb = x.schoolMailboxes && typeof x.schoolMailboxes === 'object' ? x.schoolMailboxes : {};
    return {
        planerDefaultSiteUrl: trimUrl(x.planerDefaultSiteUrl),
        planerSiteUrls: planerSiteUrls,
        schoolMailboxes: {
            elternbriefZentral: String(mb.elternbriefZentral || '')
                .trim()
                .toLowerCase(),
            elternbriefName: String(mb.elternbriefName || 'Elternbrief Schule').trim() || 'Elternbrief Schule',
            schulkalender: String(mb.schulkalender || '')
                .trim()
                .toLowerCase()
        }
    };
}

/**
 * Effektive Site-URL für ein Tool (ohne DOM).
 * @param {SchoolToolSiteKind|string} kind
 * @param {object} [setup]
 * @param {{ ignoreLocal?: boolean }} [opts]
 */
export function resolveSchoolToolSiteUrl(kind, setup, opts) {
    const id = String(kind || '').trim();
    const s = normalizeSchoolToolSitesSetup(readSetup(setup));
    const full = readSetup(setup);
    const ignoreLocal = !!(opts && opts.ignoreLocal);

    if (!ignoreLocal) {
        if (id === 'wahlfach') {
            const w = readWahlfachSiteFromStorage();
            if (w) return w;
        } else if (id === 'krankmeldung') {
            const k = readKrankmeldungSiteFromStorage();
            if (k) return k;
        } else if (id === 'einsammler') {
            const e = readEinsammlerSiteFromStorage();
            if (e) return e;
        } else if (id === 'elternbrief') {
            const eb = readElternbriefSiteFromStorage();
            if (eb) return eb;
        } else if (id === 'technik') {
            const t = readTechnikSiteFromStorage();
            if (t) return t;
        } else if (id === 'aufgaben') {
            const a = readAufgabenSiteFromStorage();
            if (a) return a;
        } else if (id === 'ausgaben') {
            const ag = readAusgabenSiteFromStorage();
            if (ag) return ag;
        } else if (id === 'raumplan') {
            const rp = readRaumplanSiteFromStorage();
            if (rp) return rp;
        } else if (id === 'kollegium') {
            const ko = readKollegiumSiteFromStorage();
            if (ko) return ko;
        } else if (id === 'stammdatenPortal') {
            const sp = readStammdatenPortalSiteFromStorage();
            if (sp) return sp;
        } else {
            const key = SCHOOL_TOOL_SITE_KEYS[id];
            const local = key ? localStorageTrim(key) : '';
            if (local) return local;
        }
    }

    if (s.planerSiteUrls[id]) return s.planerSiteUrls[id];
    if (s.planerDefaultSiteUrl) return s.planerDefaultSiteUrl;
    const intranet = trimUrl(full.intranetSiteUrl);
    if (intranet) return intranet;
    return '';
}

/**
 * Alle effektiven URLs (für UI/Datenlandkarte).
 * @param {object} [setup]
 * @param {{ ignoreLocal?: boolean }} [opts]
 */
export function resolveAllSchoolToolSiteUrls(setup, opts) {
    /** @type {Record<string, string>} */
    const out = {};
    SCHOOL_TOOL_SITE_KINDS.forEach(function (k) {
        out[k.id] = resolveSchoolToolSiteUrl(k.id, setup, opts);
    });
    out.default = normalizeSchoolToolSitesSetup(setup).planerDefaultSiteUrl || trimUrl(readSetup(setup).intranetSiteUrl);
    return out;
}

/**
 * Primär-Site-Vorschlag: Default → Intranet → Hub.
 * @param {object} [setup]
 */
export function suggestPrimarySiteUrl(setup) {
    const full = readSetup(setup);
    const s = normalizeSchoolToolSitesSetup(full);
    if (s.planerDefaultSiteUrl) return s.planerDefaultSiteUrl;
    return trimUrl(full.intranetSiteUrl) || trimUrl(full.schoolIntranetSiteUrl) || '';
}

/**
 * Speichert Setup-Felder und spiegelt zentrale Postfächer in Tool-Settings.
 * Schreibt keine Tool-localStorage-Sites (Override bleibt Tool-Sache), außer applyLocalEmpty.
 * @param {{
 *   planerDefaultSiteUrl?: string,
 *   planerSiteUrls?: Record<string, string>,
 *   schoolMailboxes?: { elternbriefZentral?: string, elternbriefName?: string, schulkalender?: string },
 *   applyLocalEmpty?: boolean
 * }} patch
 */
export function patchSchoolToolSitesSetup(patch) {
    const p = patch && typeof patch === 'object' ? patch : {};
    const cur = normalizeSchoolToolSitesSetup(readSetup());
    const nextUrls = Object.assign({}, cur.planerSiteUrls);
    if (p.planerSiteUrls && typeof p.planerSiteUrls === 'object') {
        SCHOOL_TOOL_SITE_KINDS.forEach(function (k) {
            if (Object.prototype.hasOwnProperty.call(p.planerSiteUrls, k.id)) {
                nextUrls[k.id] = trimUrl(p.planerSiteUrls[k.id]);
            }
        });
    }
    const nextMb = Object.assign({}, cur.schoolMailboxes, p.schoolMailboxes || {});
    nextMb.elternbriefZentral = String(nextMb.elternbriefZentral || '')
        .trim()
        .toLowerCase();
    nextMb.elternbriefName =
        String(nextMb.elternbriefName || 'Elternbrief Schule').trim() || 'Elternbrief Schule';
    nextMb.schulkalender = String(nextMb.schulkalender || '')
        .trim()
        .toLowerCase();

    const payload = {
        planerDefaultSiteUrl:
            p.planerDefaultSiteUrl != null ? trimUrl(p.planerDefaultSiteUrl) : cur.planerDefaultSiteUrl,
        planerSiteUrls: nextUrls,
        schoolMailboxes: nextMb
    };

    try {
        if (
            typeof window !== 'undefined' &&
            window.ms365AppDataV2 &&
            typeof window.ms365AppDataV2.patchSetup === 'function'
        ) {
            window.ms365AppDataV2.patchSetup(payload);
        }
    } catch {
        /* ignore */
    }

    syncElternbriefMailbox(nextMb);
    syncSchulkalenderHint(nextMb.schulkalender);

    if (p.applyLocalEmpty) {
        applyDefaultToEmptyToolStorage(payload);
    }

    return payload;
}

/**
 * Leere Tool-localStorage mit effektiver URL füllen (kein Überschreiben gesetzter Overrides).
 * @param {ReturnType<typeof normalizeSchoolToolSitesSetup>} [normalized]
 */
export function applyDefaultToEmptyToolStorage(normalized) {
    const setup = normalized || normalizeSchoolToolSitesSetup(readSetup());
    SCHOOL_TOOL_SITE_KINDS.forEach(function (k) {
        const url = resolveSchoolToolSiteUrl(k.id, setup, { ignoreLocal: true });
        if (!url) return;
        if (k.id === 'wahlfach') {
            if (readWahlfachSiteFromStorage()) return;
            writeSettingsSiteUrl(WAHLFACH_SETTINGS_KEY, url);
            return;
        }
        if (k.id === 'krankmeldung') {
            if (readKrankmeldungSiteFromStorage()) return;
            writeSettingsSiteUrl(KRANKMELDUNG_SP_SETTINGS_KEY, url);
            return;
        }
        if (k.id === 'einsammler') {
            if (readEinsammlerSiteFromStorage()) return;
            writeSettingsSiteUrl(EINSAMMLER_SP_SETTINGS_KEY, url);
            return;
        }
        if (k.id === 'elternbrief') {
            if (readElternbriefSiteFromStorage()) return;
            writeSettingsSiteUrl(ELTERNBRIEF_SP_SETTINGS_KEY, url);
            return;
        }
        if (k.id === 'technik') {
            if (readTechnikSiteFromStorage()) return;
            writeSettingsSiteUrl(TECHNIK_SP_SETTINGS_KEY, url);
            return;
        }
        if (k.id === 'aufgaben') {
            if (readAufgabenSiteFromStorage()) return;
            writeSettingsSiteUrl(AUFGABEN_SP_SETTINGS_KEY, url);
            return;
        }
        if (k.id === 'ausgaben') {
            if (readAusgabenSiteFromStorage()) return;
            writeSettingsSiteUrl(AUSGABEN_SP_SETTINGS_KEY, url);
            return;
        }
        if (k.id === 'raumplan') {
            if (readRaumplanSiteFromStorage()) return;
            writeSettingsSiteUrl(RAUMPLAN_SP_SETTINGS_KEY, url);
            return;
        }
        if (k.id === 'kollegium') {
            if (readKollegiumSiteFromStorage()) return;
            writeSettingsSiteUrl(KOLLEGIUM_SP_SETTINGS_KEY, url);
            return;
        }
        if (k.id === 'stammdatenPortal') {
            if (readStammdatenPortalSiteFromStorage()) return;
            writeStammdatenPortalSiteUrl(url);
            return;
        }
        const key = SCHOOL_TOOL_SITE_KEYS[k.id];
        if (!key) return;
        if (localStorageTrim(key)) return;
        try {
            localStorage.setItem(key, url);
        } catch {
            /* ignore */
        }
    });
}

function syncElternbriefMailbox(mb) {
    const email = String((mb && mb.elternbriefZentral) || '')
        .trim()
        .toLowerCase();
    if (!email) return;
    try {
        const raw = localStorage.getItem(ELTERNBRIEF_SETTINGS_KEY);
        const o = raw ? JSON.parse(raw) || {} : {};
        o.zentralPostfach = email;
        if (mb.elternbriefName) o.zentralName = String(mb.elternbriefName).trim();
        if (!o.absenderDefault) o.absenderDefault = 'zentrales_postfach';
        localStorage.setItem(ELTERNBRIEF_SETTINGS_KEY, JSON.stringify(o));
    } catch {
        /* ignore */
    }
}

function syncSchulkalenderHint(email) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!em) return;
    try {
        const raw = localStorage.getItem('ms365-kalender-postfaecher-v1');
        const o = raw ? JSON.parse(raw) || {} : {};
        const items = Array.isArray(o.items) ? o.items : [];
        const has = items.some(function (it) {
            return String((it && (it.primarySmtp || it.email || it.identity)) || '')
                .trim()
                .toLowerCase() === em;
        });
        if (!has) {
            items.push({
                id: 'stammdaten-schulkalender',
                presetId: 'schulkalender',
                displayName: 'Schulkalender',
                alias: em.split('@')[0] || 'schulkalender',
                primarySmtp: em,
                fromStammdaten: true
            });
            o.items = items;
            localStorage.setItem('ms365-kalender-postfaecher-v1', JSON.stringify(o));
        }
    } catch {
        /* ignore */
    }
}

/**
 * Zentrales Elternbrief-Postfach: Tool-Setting oder Stammdaten.
 */
export function resolveElternbriefZentralPostfach(settingsZentral) {
    const fromTool = String(settingsZentral || '')
        .trim()
        .toLowerCase();
    if (fromTool) return fromTool;
    return normalizeSchoolToolSitesSetup(readSetup()).schoolMailboxes.elternbriefZentral;
}

if (typeof window !== 'undefined') {
    window.ms365ResolveSchoolToolSiteUrl = resolveSchoolToolSiteUrl;
    window.ms365PatchSchoolToolSitesSetup = patchSchoolToolSitesSetup;
}
