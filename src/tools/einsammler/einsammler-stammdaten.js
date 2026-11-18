/**
 * Stammdaten-Anbindung für Einsammler (Klassen + Empfängerzeilen).
 * Local-first: Live-Hydrate optional; 403/Fehler wipe’t den lokalen Cache nicht.
 */

import { hydrateStammdatenFromLiveSite } from '../../shared/school-stammdaten-spo-lists.js';
import { resolveSchoolToolSiteUrl } from '../../shared/school-tool-sites.js';

const STAMM_READ_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.Read.All'
];
const STAMM_RW_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All'
];

/** @type {{ classes: Array<{ code: string, name: string }>, students: Array<object> }|null} */
let hydratedBundle = null;

/** @type {{ ok: boolean, error: string, source: string, classCount: number, studentCount: number, siteUrl: string }} */
let stammdatenMeta = {
    ok: false,
    error: '',
    source: 'local',
    classCount: 0,
    studentCount: 0,
    siteUrl: ''
};

/**
 * @returns {{ ok: boolean, error: string, source: string, classCount: number, studentCount: number, siteUrl: string }}
 */
export function getEinsammlerStammdatenMeta() {
    return { ...stammdatenMeta };
}

/**
 * @returns {{ classes: Array<{ code: string, name: string }>, students: Array<object> }}
 */
function loadLocalStammdatenBundle() {
    /** @type {Array<{ code: string, name: string }>} */
    const classes = [];
    const classSeen = new Set();
    /** @type {object[]} */
    const students = [];
    const studentSeen = new Set();

    const pushClass = (c) => {
        if (!c || typeof c !== 'object') return;
        const code = String(c.code || c.name || c.klasse || '').trim();
        if (!code) return;
        const key = code.toLowerCase();
        if (classSeen.has(key)) return;
        classSeen.add(key);
        classes.push({
            code,
            name: String(c.name || code).trim()
        });
    };

    const pushStudent = (s) => {
        if (!s || typeof s !== 'object') return;
        const name = String(s.name || '').trim();
        const klasse = String(s.klasse || s.class || s.classCode || '').trim();
        const email = String(s.email || s.mail || '')
            .trim()
            .toLowerCase();
        if (!name && !klasse && !email) return;
        const key = (email || name.toLowerCase()) + '|' + klasse.toLowerCase();
        if (studentSeen.has(key)) return;
        studentSeen.add(key);

        /** @type {Array<{ name: string, email: string }>} */
        let parents = [];
        if (Array.isArray(s.parentPairs)) {
            parents = s.parentPairs
                .map((p) => ({
                    name: String((p && p.name) || '').trim(),
                    email: String((p && p.email) || '')
                        .trim()
                        .toLowerCase()
                }))
                .filter((p) => p.email.includes('@'));
        } else if (Array.isArray(s.parents)) {
            parents = s.parents
                .map((p) => ({
                    name: String((p && p.name) || '').trim(),
                    email: String((p && p.email) || '')
                        .trim()
                        .toLowerCase()
                }))
                .filter((p) => p.email.includes('@'));
        }
        students.push({ name, klasse, email, parents });
    };

    try {
        if (typeof window !== 'undefined' && typeof window.ms365TenantSettingsLoad === 'function') {
            const core = window.ms365TenantSettingsLoad();
            const data = (core && core.data) || core || {};
            (Array.isArray(data.classes) ? data.classes : []).forEach(pushClass);
            (Array.isArray(data.students) ? data.students : []).forEach(pushStudent);
        }
    } catch {
        /* ignore */
    }

    try {
        if (
            typeof window !== 'undefined' &&
            window.ms365StammdatenCanonical &&
            typeof window.ms365StammdatenCanonical.listAllClasses === 'function'
        ) {
            (window.ms365StammdatenCanonical.listAllClasses() || []).forEach(pushClass);
        }
    } catch {
        /* ignore */
    }

    classes.sort((a, b) => a.code.localeCompare(b.code, 'de'));
    return { classes, students };
}

/**
 * @param {object[]} rawStudents
 */
function mapHydratedStudents(rawStudents) {
    return (Array.isArray(rawStudents) ? rawStudents : []).map((s) => {
        const name = String((s && (s.name || s.displayName)) || '').trim();
        const klasse = String((s && (s.klasse || s.class || s.classCode)) || '').trim();
        const email = String((s && (s.email || s.mail)) || '')
            .trim()
            .toLowerCase();
        /** @type {Array<{ name: string, email: string }>} */
        let parents = [];
        if (Array.isArray(s && s.parentPairs)) {
            parents = s.parentPairs
                .map((p) => ({
                    name: String((p && p.name) || '').trim(),
                    email: String((p && p.email) || '')
                        .trim()
                        .toLowerCase()
                }))
                .filter((p) => p.email.includes('@'));
        } else if (Array.isArray(s && s.parents)) {
            parents = s.parents
                .map((p) => ({
                    name: String((p && p.name) || '').trim(),
                    email: String((p && p.email) || '')
                        .trim()
                        .toLowerCase()
                }))
                .filter((p) => p.email.includes('@'));
        }
        return { name, klasse, email, parents };
    });
}

/**
 * @returns {{ classes: Array<{ code: string, name: string }>, students: Array<object> }}
 */
export function loadEinsammlerStammdaten() {
    if (hydratedBundle) {
        return {
            classes: hydratedBundle.classes.slice(),
            students: hydratedBundle.students.slice()
        };
    }
    const local = loadLocalStammdatenBundle();
    stammdatenMeta = {
        ok: local.classes.length > 0,
        error: '',
        source: 'local',
        classCount: local.classes.length,
        studentCount: local.students.length,
        siteUrl: String(stammdatenMeta.siteUrl || '')
    };
    return local;
}

/**
 * @param {{ getGraphToken: Function }} api
 */
async function acquireStammdatenReadToken(api) {
    if (typeof window !== 'undefined' && typeof window.ms365AuthAcquireTokenSilent === 'function') {
        try {
            return await window.ms365AuthAcquireTokenSilent(STAMM_READ_SCOPES);
        } catch {
            /* weiter */
        }
        try {
            return await window.ms365AuthAcquireTokenSilent(STAMM_RW_SCOPES);
        } catch {
            /* weiter */
        }
    }
    try {
        return await api.getGraphToken(STAMM_READ_SCOPES);
    } catch {
        return await api.getGraphToken(STAMM_RW_SCOPES);
    }
}

function friendlyStammdatenError(msg) {
    const s = String(msg || '');
    if (/403|accessdenied|forbidden/i.test(s)) {
        return 'Kein Zugriff auf Stammdaten-Listen. Lokale Klassen bleiben nutzbar.';
    }
    if (/nicht angemeldet|anmelden/i.test(s)) {
        return 'Bitte anmelden – sonst bleiben nur lokale Stammdaten.';
    }
    return s.slice(0, 160);
}

/**
 * Live-Stammdaten laden; bei Fehler/403 lokalen Stand behalten.
 * @param {{ siteKind?: string }} [opts] z. B. `elternbrief` (gleiche Ablauf-Familie)
 * @returns {Promise<ReturnType<typeof getEinsammlerStammdatenMeta>>}
 */
export async function hydrateEinsammlerStammdaten(opts) {
    const local = loadLocalStammdatenBundle();
    const siteKind = String((opts && opts.siteKind) || 'einsammler').trim() || 'einsammler';
    const siteUrl = String(resolveSchoolToolSiteUrl(siteKind) || '').trim();
    stammdatenMeta = {
        ok: local.classes.length > 0,
        error: '',
        source: 'local',
        classCount: local.classes.length,
        studentCount: local.students.length,
        siteUrl
    };

    if (!siteUrl) {
        hydratedBundle = null;
        return getEinsammlerStammdatenMeta();
    }

    const api = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!api || typeof api.getGraphToken !== 'function' || typeof api.resolveSiteFromWebUrl !== 'function') {
        stammdatenMeta.error = '';
        hydratedBundle = null;
        return getEinsammlerStammdatenMeta();
    }

    try {
        const tok = await acquireStammdatenReadToken(api);
        const site = await api.resolveSiteFromWebUrl(tok, siteUrl);
        const siteId = site && site.id ? String(site.id) : '';
        if (!siteId) {
            stammdatenMeta.error = 'Site nicht gefunden';
            stammdatenMeta.source = 'local-fallback';
            hydratedBundle = null;
            return getEinsammlerStammdatenMeta();
        }

        const result = await hydrateStammdatenFromLiveSite({
            token: tok,
            siteId,
            siteWebUrl: siteUrl,
            graph: api,
            local: {
                teachers: [],
                classes: local.classes,
                students: local.students
            }
        });

        const mergedClasses = Array.isArray(result.stammdaten && result.stammdaten.classes)
            ? result.stammdaten.classes
                  .map((c) => ({
                      code: String((c && (c.code || c.name || c.klasse)) || '').trim(),
                      name: String((c && (c.name || c.code || c.klasse)) || '').trim()
                  }))
                  .filter((c) => c.code)
            : local.classes;
        const mergedStudents = mapHydratedStudents(
            (result.stammdaten && result.stammdaten.students) || local.students
        );
        const liveClasses = (result.live && result.live.classes) || [];
        const usedLive = liveClasses.length > 0;

        // Nie leeren Stand übernehmen, wenn lokal schon Klassen da waren.
        if (mergedClasses.length || local.classes.length) {
            hydratedBundle = {
                classes: (mergedClasses.length ? mergedClasses : local.classes).slice().sort((a, b) =>
                    a.code.localeCompare(b.code, 'de')
                ),
                students: (mergedStudents.length ? mergedStudents : local.students).slice()
            };
        } else {
            hydratedBundle = null;
        }

        const bundle = hydratedBundle || local;
        stammdatenMeta = {
            ok: usedLive || bundle.classes.length > 0,
            error: usedLive ? '' : friendlyStammdatenError(result.error || ''),
            source: usedLive ? 'sharepoint' : result.ok ? 'local-fallback' : 'local-fallback',
            classCount: bundle.classes.length,
            studentCount: bundle.students.length,
            siteUrl
        };
        return getEinsammlerStammdatenMeta();
    } catch (e) {
        const msg = e && e.message ? e.message : String(e);
        // 403/Fehler: lokalen Cache behalten, kein Wipe.
        hydratedBundle = null;
        stammdatenMeta = {
            ok: local.classes.length > 0,
            error: friendlyStammdatenError(msg),
            source: 'local-fallback',
            classCount: local.classes.length,
            studentCount: local.students.length,
            siteUrl
        };
        return getEinsammlerStammdatenMeta();
    }
}

/**
 * Empfängerzeilen für gewählte Klassen: Name;Klasse;Eltern-E-Mail
 * Pro Erziehungsberechtigte:r mit E-Mail eine Zeile; sonst eine Zeile ohne Mail.
 * @param {string[]} klasseCodes
 * @param {object[]} [students]
 */
export function buildEmpfaengerLinesForKlassen(klasseCodes, students) {
    const wanted = new Set(
        (klasseCodes || []).map((c) => String(c || '').trim().toLowerCase()).filter(Boolean)
    );
    if (!wanted.size) return [];
    const list = Array.isArray(students) ? students : loadEinsammlerStammdaten().students;
    /** @type {string[]} */
    const lines = [];
    const seen = new Set();
    const pushLine = (name, klasse, email) => {
        const line = name + ';' + klasse + ';' + String(email || '').trim().toLowerCase();
        const key = line.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        lines.push(line);
    };
    for (const s of list) {
        const k = String(s.klasse || '').trim();
        if (!wanted.has(k.toLowerCase())) continue;
        const name = String(s.name || '').trim() || 'Schüler:in';
        const parents = Array.isArray(s.parents) ? s.parents : [];
        const parentMails = parents
            .map((p) => String((p && p.email) || '').trim().toLowerCase())
            .filter((em) => em.includes('@'));
        if (parentMails.length) {
            parentMails.forEach((em) => pushLine(name, k, em));
        } else {
            pushLine(name, k, '');
        }
    }
    return lines;
}

/**
 * @param {string} text
 * @returns {string[]}
 */
export function parseKlasseCodesField(text) {
    return String(text || '')
        .split(/[;,\n]+/)
        .map((s) => s.trim())
        .filter(Boolean);
}
