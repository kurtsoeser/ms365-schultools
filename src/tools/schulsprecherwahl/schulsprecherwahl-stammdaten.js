/**
 * Stammdaten-Pools für Schulwahlen (Schüler / Lehrer).
 * Local-first: Live-Hydrate optional über Schul-Site; 403 wipe’t den Cache nicht.
 * Stimmen liegen NICHT in SharePoint – Site nur für Personenpools.
 */

import { hydrateStammdatenFromLiveSite } from '../../shared/school-stammdaten-spo-lists.js';
import { resolveSchoolToolSiteUrl } from '../../shared/school-tool-sites.js';
import { loadStudentPool as loadLocalStudents } from './schulsprecherwahl-students.js';
import { loadTeacherPool as loadLocalTeachers } from './schulsprecherwahl-teachers.js';

const STAMM_READ_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.Read.All'
];
const STAMM_RW_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All'
];

/** @type {{ name: string, className: string, email: string }[]|null} */
let hydratedStudents = null;
/** @type {{ name: string, className: string, email: string, code: string }[]|null} */
let hydratedTeachers = null;

/** @type {{ ok: boolean, error: string, source: string, studentCount: number, teacherCount: number, siteUrl: string }} */
let stammdatenMeta = {
    ok: false,
    error: '',
    source: 'local',
    studentCount: 0,
    teacherCount: 0,
    siteUrl: ''
};

/**
 * Site für Personenpools: Default-/Intranet-Site, nicht „Stimmen-Liste“.
 * @returns {string}
 */
export function resolveElectionStammdatenSiteUrl() {
    // Bekannte Kind-Keys fallen auf planerDefault/Intranet zurück, wenn leer.
    const fromDefault = String(resolveSchoolToolSiteUrl('aktivitaeten') || '').trim();
    if (fromDefault) return fromDefault;
    try {
        if (
            typeof window !== 'undefined' &&
            window.ms365AppDataV2 &&
            typeof window.ms365AppDataV2.getSetup === 'function'
        ) {
            const setup = window.ms365AppDataV2.getSetup() || {};
            return String(setup.planerDefaultSiteUrl || setup.intranetSiteUrl || '').trim();
        }
    } catch {
        /* ignore */
    }
    return '';
}

/**
 * @returns {{ ok: boolean, error: string, source: string, studentCount: number, teacherCount: number, siteUrl: string }}
 */
export function getSchulsprecherwahlStammdatenMeta() {
    return { ...stammdatenMeta };
}

/**
 * Effektiver Schüler-Pool (hydrate → lokal).
 * @returns {{ name: string, className: string, email: string }[]}
 */
export function loadStudentPoolEffective() {
    if (hydratedStudents && hydratedStudents.length) return hydratedStudents.slice();
    return loadLocalStudents();
}

/**
 * Effektiver Lehrer-Pool (hydrate → lokal).
 * @returns {{ name: string, className: string, email: string, code: string }[]}
 */
export function loadTeacherPoolEffective() {
    if (hydratedTeachers && hydratedTeachers.length) return hydratedTeachers.slice();
    return loadLocalTeachers();
}

/**
 * @param {string} email
 */
export function findStudentByEmailEffective(email) {
    const key = String(email || '')
        .trim()
        .toLowerCase();
    if (!key) return null;
    return loadStudentPoolEffective().find((s) => s.email === key) || null;
}

/**
 * @param {string} email
 */
export function findTeacherByEmailEffective(email) {
    const key = String(email || '')
        .trim()
        .toLowerCase();
    if (!key) return null;
    return loadTeacherPoolEffective().find((t) => t.email === key) || null;
}

/**
 * @param {string} query
 * @param {number} [limit]
 */
export function searchStudentPoolEffective(query, limit = 40) {
    const q = String(query || '')
        .trim()
        .toLowerCase();
    const all = loadStudentPoolEffective();
    if (!q) return all.slice(0, limit);
    return all
        .filter(
            (s) =>
                s.name.toLowerCase().includes(q) ||
                s.className.toLowerCase().includes(q) ||
                s.email.includes(q)
        )
        .slice(0, limit);
}

/**
 * @param {string} query
 * @param {number} [limit]
 */
export function searchTeacherPoolEffective(query, limit = 40) {
    const q = String(query || '')
        .trim()
        .toLowerCase();
    const all = loadTeacherPoolEffective();
    if (!q) return all.slice(0, limit);
    return all
        .filter(
            (t) =>
                t.name.toLowerCase().includes(q) ||
                t.className.toLowerCase().includes(q) ||
                String(t.code || '')
                    .toLowerCase()
                    .includes(q) ||
                t.email.includes(q)
        )
        .slice(0, limit);
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

/**
 * @param {string} msg
 */
function friendlyStammdatenError(msg) {
    const s = String(msg || '');
    if (/403|accessdenied|forbidden/i.test(s)) {
        return 'Kein Zugriff auf Stammdaten-Listen. Lokale Pools bleiben nutzbar.';
    }
    if (/nicht angemeldet|anmelden|401|unauthorized/i.test(s)) {
        return 'Bitte anmelden – sonst bleiben nur lokale Stammdaten.';
    }
    return s.slice(0, 160);
}

/**
 * @param {unknown[]} list
 * @returns {{ name: string, className: string, email: string }[]}
 */
function mapStudents(list) {
    return (Array.isArray(list) ? list : [])
        .map((s) => {
            const o = s && typeof s === 'object' ? s : {};
            return {
                name: String(o.name || o.displayName || '').trim(),
                className: String(o.klasse || o.className || o.class || o.classCode || '').trim(),
                email: String(o.email || o.mail || o.upn || '')
                    .trim()
                    .toLowerCase()
            };
        })
        .filter((s) => s.name || s.email)
        .sort((a, b) => {
            const kc = a.className.localeCompare(b.className, 'de');
            if (kc) return kc;
            return a.name.localeCompare(b.name, 'de');
        });
}

/**
 * @param {unknown[]} list
 * @returns {{ name: string, className: string, email: string, code: string }[]}
 */
function mapTeachers(list) {
    return (Array.isArray(list) ? list : [])
        .map((t) => {
            const o = t && typeof t === 'object' ? t : {};
            const code = String(o.code || o.kuerzel || o.shortName || o.id || '').trim();
            return {
                name: String(o.name || o.displayName || '').trim(),
                className: code,
                email: String(o.email || o.mail || o.upn || '')
                    .trim()
                    .toLowerCase(),
                code
            };
        })
        .filter((t) => t.name || t.email)
        .sort((a, b) => a.name.localeCompare(b.name, 'de'));
}

/**
 * Live-Pools laden; bei Fehler/403 lokalen Stand behalten (kein Wipe).
 * @returns {Promise<ReturnType<typeof getSchulsprecherwahlStammdatenMeta>>}
 */
export async function hydrateSchulsprecherwahlStammdaten() {
    const localStudents = loadLocalStudents();
    const localTeachers = loadLocalTeachers();
    const siteUrl = resolveElectionStammdatenSiteUrl();

    stammdatenMeta = {
        ok: localStudents.length > 0 || localTeachers.length > 0,
        error: '',
        source: 'local',
        studentCount: localStudents.length,
        teacherCount: localTeachers.length,
        siteUrl
    };
    hydratedStudents = null;
    hydratedTeachers = null;

    if (!siteUrl) {
        return getSchulsprecherwahlStammdatenMeta();
    }

    const api = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!api || typeof api.getGraphToken !== 'function' || typeof api.resolveSiteFromWebUrl !== 'function') {
        return getSchulsprecherwahlStammdatenMeta();
    }

    try {
        const tok = await acquireStammdatenReadToken(api);
        const site = await api.resolveSiteFromWebUrl(tok, siteUrl);
        const siteId = site && site.id ? String(site.id) : '';
        if (!siteId) {
            stammdatenMeta.error = 'Site nicht gefunden';
            stammdatenMeta.source = 'local-fallback';
            return getSchulsprecherwahlStammdatenMeta();
        }

        const result = await hydrateStammdatenFromLiveSite({
            token: tok,
            siteId,
            siteWebUrl: siteUrl,
            graph: api,
            local: {
                teachers: localTeachers,
                classes: [],
                students: localStudents
            }
        });

        const liveStudents = mapStudents((result.stammdaten && result.stammdaten.students) || []);
        const liveTeachers = mapTeachers((result.stammdaten && result.stammdaten.teachers) || []);
        const usedLive = liveStudents.length > 0 || liveTeachers.length > 0;

        // Nie leeren Stand übernehmen, wenn lokal schon Daten da waren.
        hydratedStudents = liveStudents.length ? liveStudents : null;
        hydratedTeachers = liveTeachers.length ? liveTeachers : null;

        const students = loadStudentPoolEffective();
        const teachers = loadTeacherPoolEffective();

        stammdatenMeta = {
            ok: usedLive || students.length > 0 || teachers.length > 0,
            error: usedLive ? '' : friendlyStammdatenError(result.error || ''),
            source: usedLive ? 'sharepoint' : 'local-fallback',
            studentCount: students.length,
            teacherCount: teachers.length,
            siteUrl
        };
        return getSchulsprecherwahlStammdatenMeta();
    } catch (e) {
        const msg = e && e.message ? e.message : String(e);
        // 403/Fehler: lokalen Cache behalten, kein Wipe.
        hydratedStudents = null;
        hydratedTeachers = null;
        stammdatenMeta = {
            ok: localStudents.length > 0 || localTeachers.length > 0,
            error: friendlyStammdatenError(msg),
            source: 'local-fallback',
            studentCount: localStudents.length,
            teacherCount: localTeachers.length,
            siteUrl
        };
        return getSchulsprecherwahlStammdatenMeta();
    }
}
