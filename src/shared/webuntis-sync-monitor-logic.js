/**
 * WebUntis-/Kursteam-Sync-Monitor (pure Logic, Analyse 04 A3).
 * Liest Snapshot aus webuntis-teams-creator-state-v1 / Kursteams-State.
 */

export const KURSTEAM_STORAGE_KEY = 'webuntis-teams-creator-state-v1';

function norm(v) {
    return String(v == null ? '' : v)
        .trim()
        .toLowerCase();
}

function normCode(v) {
    return norm(v).replace(/\s+/g, '');
}

/**
 * @param {unknown} raw
 * @returns {{ ok: boolean, reason?: string, state?: object }}
 */
export function parseKursteamState(raw) {
    if (raw == null || raw === '') {
        return { ok: false, reason: 'empty' };
    }
    let obj = raw;
    if (typeof raw === 'string') {
        try {
            obj = JSON.parse(raw);
        } catch {
            return { ok: false, reason: 'invalid-json' };
        }
    }
    if (!obj || typeof obj !== 'object') return { ok: false, reason: 'invalid' };
    const has =
        Array.isArray(obj.rawData) ||
        Array.isArray(obj.filteredData) ||
        Array.isArray(obj.teamsData) ||
        obj.kind === 'ms365-kursteams-state';
    if (!has) return { ok: false, reason: 'not-kursteam' };
    return { ok: true, state: obj };
}

/**
 * @param {object} state
 * @param {{ teachers?: { code?: string, email?: string, name?: string }[], classes?: { code?: string, name?: string }[] }} [stammdaten]
 */
export function buildSyncMonitorReport(state, stammdaten) {
    const st = state && typeof state === 'object' ? state : {};
    const teams = Array.isArray(st.teamsData) ? st.teamsData : [];
    const raw = Array.isArray(st.rawData) ? st.rawData : [];
    const filtered = Array.isArray(st.filteredData) ? st.filteredData : [];
    const mapping =
        st.teacherEmailMapping && typeof st.teacherEmailMapping === 'object' ? st.teacherEmailMapping : {};
    const teachers = Array.isArray(stammdaten && stammdaten.teachers) ? stammdaten.teachers : [];
    const classes = Array.isArray(stammdaten && stammdaten.classes) ? stammdaten.classes : [];

    const teacherByCode = new Map();
    teachers.forEach(function (t) {
        const c = normCode(t && t.code);
        if (c) teacherByCode.set(c, t);
    });

    const missingOwner = [];
    const missingEmail = [];
    const invalidTeam = [];
    const ready = [];

    teams.forEach(function (t, idx) {
        const name = String((t && (t.teamName || t.displayName)) || '').trim() || 'Team #' + (idx + 1);
        const besitzer = String((t && t.besitzer) || '').trim();
        const mail = String((t && t.gruppenmail) || '').trim();
        const code = normCode(t && t.lehrerCode);
        const row = {
            index: idx,
            teamName: name,
            lehrerCode: String((t && t.lehrerCode) || '').trim(),
            fach: String((t && t.fach) || '').trim(),
            originalClass: String((t && t.originalClass) || '').trim(),
            gruppenmail: mail,
            besitzer: besitzer,
            isValid: !!(t && t.isValid)
        };
        if (!besitzer) missingOwner.push(row);
        if (!mail) missingEmail.push(row);
        if (t && t.isValid === false) invalidTeam.push(row);
        if (besitzer && mail && t && t.isValid !== false) ready.push(row);

        if (code && !mapping[String((t && t.lehrerCode) || '').toUpperCase().trim()]) {
            const stTeacher = teacherByCode.get(code);
            if (!stTeacher || !String(stTeacher.email || '').trim()) {
                /* counted via missingOwner / mapping below */
            }
        }
    });

    const codesInData = new Set();
    (filtered.length ? filtered : raw).forEach(function (r) {
        const k = String((r && r.lehrer) || '')
            .toUpperCase()
            .trim();
        if (k) codesInData.add(k);
    });

    const teachersWithoutMatch = [];
    codesInData.forEach(function (code) {
        const mapped = mapping[code];
        const fromStamm = teacherByCode.get(normCode(code));
        const email = String(mapped || (fromStamm && fromStamm.email) || '').trim();
        if (!email || email.indexOf('@') === -1) {
            teachersWithoutMatch.push({
                code: code,
                name: (fromStamm && fromStamm.name) || '',
                hint: mapped ? 'Mapping ungültig' : fromStamm ? 'E-Mail in Stammdaten fehlt' : 'Kürzel nicht in Stammdaten'
            });
        }
    });

    const classCodesInTeams = new Set();
    teams.forEach(function (t) {
        const oc = String((t && t.originalClass) || '').trim();
        if (oc) {
            oc.split(/[,;/|]+/).forEach(function (part) {
                const c = normCode(part);
                if (c) classCodesInTeams.add(c);
            });
        }
    });

    const classesWithoutTeam = [];
    classes.forEach(function (c) {
        const code = normCode(c && (c.code || c.name));
        if (!code) return;
        if (!classCodesInTeams.has(code)) {
            classesWithoutTeam.push({
                code: String((c && (c.code || c.name)) || '').trim(),
                name: String((c && c.name) || '').trim()
            });
        }
    });

    const savedAt = String(st.savedAt || st.updatedAt || st.lastSavedAt || '').trim();
    const entryMode = String(st.kursteamEntryMode || '').trim() || (raw.length ? 'webuntis' : 'unset');
    const yearPrefix = String(st.yearPrefix || '').trim();

    return {
        savedAt: savedAt,
        entryMode: entryMode,
        yearPrefix: yearPrefix,
        counts: {
            rawRows: raw.length,
            filteredRows: filtered.length,
            teams: teams.length,
            ready: ready.length,
            missingOwner: missingOwner.length,
            missingEmail: missingEmail.length,
            invalid: invalidTeam.length,
            teachersWithoutMatch: teachersWithoutMatch.length,
            classesWithoutTeam: classesWithoutTeam.length
        },
        missingOwner: missingOwner,
        missingEmail: missingEmail,
        invalidTeam: invalidTeam,
        teachersWithoutMatch: teachersWithoutMatch,
        classesWithoutTeam: classesWithoutTeam,
        actionable: missingOwner.length + teachersWithoutMatch.length + invalidTeam.length
    };
}

export function loadKursteamStateFromStorage(storage) {
    const store = storage || (typeof localStorage !== 'undefined' ? localStorage : null);
    if (!store || typeof store.getItem !== 'function') {
        return { ok: false, reason: 'no-storage' };
    }
    return parseKursteamState(store.getItem(KURSTEAM_STORAGE_KEY));
}

const api = {
    KURSTEAM_STORAGE_KEY: KURSTEAM_STORAGE_KEY,
    parseKursteamState: parseKursteamState,
    buildSyncMonitorReport: buildSyncMonitorReport,
    loadKursteamStateFromStorage: loadKursteamStateFromStorage
};

if (typeof window !== 'undefined') {
    window.ms365WebuntisSyncMonitor = api;
}

export default api;
