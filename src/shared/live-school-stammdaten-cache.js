/**
 * Live-Stammdaten (Lehrer/Klassen) aus der Schul-Tool-Site für Browser ohne IT-Lokalstand.
 * Portal-Persona und Planer lesen denselben Cache, sobald SharePoint geliefert hat.
 */

const CACHE_KEY = 'ms365-live-school-stammdaten-v1';

/**
 * @typedef {{
 *   teachers?: object[],
 *   classes?: object[],
 *   siteUrl?: string,
 *   source?: string,
 *   at?: number
 * }} LiveSchoolStammdatenCache
 */

/**
 * @returns {LiveSchoolStammdatenCache}
 */
export function loadLiveSchoolStammdatenCache() {
    try {
        if (typeof sessionStorage === 'undefined') return {};
        const raw = sessionStorage.getItem(CACHE_KEY);
        if (!raw) return {};
        const data = JSON.parse(raw);
        if (!data || typeof data !== 'object') return {};
        return {
            teachers: Array.isArray(data.teachers) ? data.teachers : [],
            classes: Array.isArray(data.classes) ? data.classes : [],
            siteUrl: String(data.siteUrl || '').trim(),
            source: String(data.source || '').trim(),
            at: Number(data.at) || 0
        };
    } catch {
        return {};
    }
}

/**
 * @param {LiveSchoolStammdatenCache|null|undefined} payload
 */
export function saveLiveSchoolStammdatenCache(payload) {
    const p = payload && typeof payload === 'object' ? payload : {};
    const teachers = Array.isArray(p.teachers) ? p.teachers : [];
    const classes = Array.isArray(p.classes) ? p.classes : [];
    if (!teachers.length && !classes.length) return loadLiveSchoolStammdatenCache();
    const next = {
        teachers,
        classes,
        siteUrl: String(p.siteUrl || '').trim(),
        source: String(p.source || 'sharepoint').trim() || 'sharepoint',
        at: Date.now()
    };
    try {
        if (typeof sessionStorage !== 'undefined') {
            sessionStorage.setItem(CACHE_KEY, JSON.stringify(next));
        }
    } catch {
        /* ignore */
    }
    return next;
}

/**
 * Lehrer/Klassen zusammenführen: Live gewinnt bei Inhalt, sonst lokal.
 * @param {{ teachers?: unknown[], classes?: unknown[] }|null|undefined} local
 * @param {{ teachers?: unknown[], classes?: unknown[] }|null|undefined} live
 */
export function mergeSchoolStammdatenLocalAndLive(local, live) {
    const loc = local || {};
    const liv = live || {};
    const liveTeachers = Array.isArray(liv.teachers) ? liv.teachers : [];
    const liveClasses = Array.isArray(liv.classes) ? liv.classes : [];
    const localTeachers = Array.isArray(loc.teachers) ? loc.teachers : [];
    const localClasses = Array.isArray(loc.classes) ? loc.classes : [];
    return {
        teachers: liveTeachers.length ? liveTeachers : localTeachers,
        classes: liveClasses.length ? liveClasses : localClasses
    };
}

/**
 * Nur Demo-/Platzhalter-Lehrer (z. B. „Lehrer01 DEMO“) – darf SharePoint-Volliste nicht blockieren.
 * @param {unknown} teachers
 * @returns {boolean}
 */
export function isSparseOrDemoOnlyTeacherList(teachers) {
    const list = Array.isArray(teachers) ? teachers : [];
    if (!list.length) return true;
    if (list.length >= 8) return false;
    let demoish = 0;
    for (let i = 0; i < list.length; i++) {
        const t = list[i] || {};
        const em = String(t.email || t.upn || t.mail || '')
            .trim()
            .toLowerCase();
        const name = String(t.name || t.Title || '').trim();
        const code = String(t.code || '').trim();
        if (
            /\.demo@/i.test(em) ||
            /@ms365\.schule$/i.test(em) ||
            /\bDEMO\b/i.test(name) ||
            /^L\d{2}$/i.test(code)
        ) {
            demoish++;
        }
    }
    // Alles Demo, oder winzige Liste (1–2) die keine echte Kollegiumsliste sein kann
    return demoish === list.length || list.length <= 2;
}

/**
 * Die reichere Lehrerliste wählen (Live gewinnt gegen Demo-/Mini-Stand).
 * @param {unknown} local
 * @param {unknown} live
 * @returns {object[]}
 */
export function preferRicherTeacherList(local, live) {
    const loc = Array.isArray(local) ? local : [];
    const liv = Array.isArray(live) ? live : [];
    if (!liv.length) return loc;
    if (!loc.length) return liv;
    if (liv.length > loc.length) return liv;
    if (isSparseOrDemoOnlyTeacherList(loc) && !isSparseOrDemoOnlyTeacherList(liv)) return liv;
    return loc;
}

/**
 * Live-Stand in Tenant-Settings schreiben, wenn lokal leer/dünn/nur Demo
 * oder Live klar mehr Lehrkräfte hat.
 *
 * @param {object[]} teachers
 * @param {object[]} [classes]
 * @returns {{ saved: boolean, teacherCount: number, classCount: number }}
 */
export function mergeLiveTeachersIntoLocalTenantIfEmpty(teachers, classes) {
    const list = Array.isArray(teachers) ? teachers : [];
    const classList = Array.isArray(classes) ? classes : [];
    if (!list.length && !classList.length) {
        return { saved: false, teacherCount: 0, classCount: 0 };
    }
    if (typeof window === 'undefined' || typeof window.ms365TenantSettingsLoad !== 'function') {
        return { saved: false, teacherCount: list.length, classCount: classList.length };
    }
    try {
        const cur = window.ms365TenantSettingsLoad() || {};
        const existingT = Array.isArray(cur.teachers) ? cur.teachers : [];
        const existingC = Array.isArray(cur.classes) ? cur.classes : [];
        const nextTeachers = preferRicherTeacherList(existingT, list);
        const nextClasses =
            classList.length > existingC.length || !existingC.length ? classList : existingC;
        const tChanged = nextTeachers.length !== existingT.length || nextTeachers !== existingT;
        const cChanged = nextClasses.length !== existingC.length;
        if (!tChanged && !cChanged) {
            return {
                saved: false,
                teacherCount: existingT.length,
                classCount: existingC.length
            };
        }
        if (typeof window.ms365TenantSettingsSave !== 'function') {
            return { saved: false, teacherCount: nextTeachers.length, classCount: nextClasses.length };
        }
        window.ms365TenantSettingsSave({
            ...cur,
            teachers: nextTeachers,
            classes: nextClasses
        });
        try {
            window.dispatchEvent(
                new CustomEvent('ms365-tenant-settings-changed', {
                    detail: { source: 'live-spo-stammdaten' }
                })
            );
        } catch {
            /* ignore */
        }
        return {
            saved: true,
            teacherCount: nextTeachers.length,
            classCount: nextClasses.length
        };
    } catch {
        return { saved: false, teacherCount: list.length, classCount: classList.length };
    }
}

/**
 * Effektive Lehrerliste für Persona/Tools: Tenant → App-Data → Live-Cache.
 * @returns {object[]}
 */
export function loadEffectiveTeachersForPersona() {
    /** @type {object[]} */
    let teachers = [];
    try {
        if (typeof window !== 'undefined' && typeof window.ms365TenantSettingsLoad === 'function') {
            const s = window.ms365TenantSettingsLoad() || {};
            if (Array.isArray(s.teachers) && s.teachers.length) teachers = s.teachers;
        }
    } catch {
        /* ignore */
    }
    if (!teachers.length) {
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getCore === 'function') {
                const core = window.ms365AppDataV2.getCore() || {};
                if (Array.isArray(core.teachers) && core.teachers.length) teachers = core.teachers;
            }
        } catch {
            /* ignore */
        }
    }
    if (!teachers.length) {
        const live = loadLiveSchoolStammdatenCache();
        if (Array.isArray(live.teachers) && live.teachers.length) teachers = live.teachers;
    }
    return teachers;
}
