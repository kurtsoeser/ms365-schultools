/**
 * Stammdaten-Anbindung Technik-Meldungen (Meta + local-first).
 * Live-Hydrate optional; 403/Fehler wipe’t lokalen Cache nicht.
 */

import { hydrateStammdatenFromLiveSite } from '../../shared/school-stammdaten-spo-lists.js';
import { resolveTechnikSiteUrl } from './technik-meldungen-settings.js';

const STAMM_READ_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.Read.All'
];

/** @type {{ ok: boolean, error: string, source: string, classCount: number, teacherCount: number, siteUrl: string }} */
let stammdatenMeta = {
    ok: false,
    error: '',
    source: 'local',
    classCount: 0,
    teacherCount: 0,
    siteUrl: ''
};

/**
 * @returns {{ ok: boolean, error: string, source: string, classCount: number, teacherCount: number, siteUrl: string }}
 */
export function getTechnikStammdatenMeta() {
    return { ...stammdatenMeta };
}

function readLocalCounts() {
    let classCount = 0;
    let teacherCount = 0;
    try {
        if (
            typeof window !== 'undefined' &&
            window.ms365AppDataV2 &&
            typeof window.ms365AppDataV2.getSetup === 'function'
        ) {
            const setup = window.ms365AppDataV2.getSetup() || {};
            const classes = Array.isArray(setup.classes) ? setup.classes : [];
            const teachers = Array.isArray(setup.teachers) ? setup.teachers : [];
            classCount = classes.length;
            teacherCount = teachers.length;
        }
    } catch {
        /* ignore */
    }
    return { classCount, teacherCount };
}

/**
 * Soft-Hydrate: Fehler → Meta setzen, keine Wipe der App-Daten.
 * @param {{ force?: boolean }} [opts]
 */
export async function hydrateTechnikStammdaten(opts = {}) {
    const siteUrl = resolveTechnikSiteUrl();
    const local = readLocalCounts();
    stammdatenMeta = {
        ok: local.classCount > 0 || local.teacherCount > 0,
        error: '',
        source: 'local',
        classCount: local.classCount,
        teacherCount: local.teacherCount,
        siteUrl
    };

    if (!siteUrl) {
        stammdatenMeta.error = siteUrl ? '' : 'Keine Tool-Site hinterlegt (System / Listen-Assistent).';
        return getTechnikStammdatenMeta();
    }

    try {
        const getToken =
            typeof window !== 'undefined' && window.ms365SpoGraphShared
                ? window.ms365SpoGraphShared.getGraphToken
                : typeof window !== 'undefined' && typeof window.ms365GetGraphToken === 'function'
                  ? window.ms365GetGraphToken
                  : null;
        if (typeof getToken !== 'function') {
            return getTechnikStammdatenMeta();
        }
        const result = await hydrateStammdatenFromLiveSite({
            siteUrl,
            getToken: () => getToken(STAMM_READ_SCOPES),
            force: !!opts.force
        });
        if (result && result.ok) {
            const nextLocal = readLocalCounts();
            stammdatenMeta = {
                ok: true,
                error: '',
                source: 'live',
                classCount: nextLocal.classCount || local.classCount,
                teacherCount: nextLocal.teacherCount || local.teacherCount,
                siteUrl
            };
        } else {
            const err = String((result && result.error) || 'Stammdaten nicht erreichbar.');
            const denied = /403|Forbidden|Zugriff/i.test(err);
            stammdatenMeta = {
                ...stammdatenMeta,
                ok: stammdatenMeta.ok,
                error: denied ? '' : err,
                source: stammdatenMeta.ok ? 'local' : 'error',
                siteUrl
            };
        }
    } catch (e) {
        const msg = String((e && e.message) || e || '');
        const denied = /403|Forbidden|Zugriff/i.test(msg);
        stammdatenMeta = {
            ...stammdatenMeta,
            error: denied ? '' : msg,
            source: stammdatenMeta.ok ? 'local' : 'error',
            siteUrl
        };
    }
    return getTechnikStammdatenMeta();
}
