/**
 * Querschnitts-Übersicht: gematchte Gruppen vs. lokale Stammlisten.
 * @file
 */
import { normEmailList } from './membership-reconcile.js';

const SCAN_STORAGE_KEY = 'ms365-hygiene-scan-v2';

function normStr(v) {
    return String(v ?? '').trim();
}

function normCode(v) {
    return normStr(v).toUpperCase();
}

function collectTeacherEmails(teachers) {
    return normEmailList(
        (Array.isArray(teachers) ? teachers : []).map(function (r) {
            return r && r.email;
        })
    );
}

function collectStudentEmails(students) {
    return normEmailList(
        (Array.isArray(students) ? students : []).map(function (r) {
            return r && r.email;
        })
    );
}

function collectAdminEmails(admin) {
    return normEmailList(
        (Array.isArray(admin) ? admin : []).map(function (r) {
            return r && r.email;
        })
    );
}

function emailsForClassCode(students, classCode, className) {
    const code = normCode(classCode);
    const name = normStr(className).toLowerCase();
    const seen = new Set();
    const out = [];
    (Array.isArray(students) ? students : []).forEach(function (s) {
        const k = normStr(s && s.klasse);
        if (!k) return;
        const match = (code && normCode(k) === code) || (name && k.toLowerCase() === name);
        if (!match) return;
        const em = String((s && s.email) || '')
            .trim()
            .toLowerCase();
        if (!em || em.indexOf('@') === -1 || seen.has(em)) return;
        seen.add(em);
        out.push(em);
    });
    return out;
}

function deriveClassStableNick(cls) {
    if (!cls) return '';
    if (typeof globalThis !== 'undefined' && typeof globalThis.ms365DeriveClassStableMailNickname === 'function') {
        const d = String(globalThis.ms365DeriveClassStableMailNickname(cls.year || '', cls.code || '') || '')
            .trim()
            .replace(/[^a-zA-Z0-9]/g, '')
            .toLowerCase()
            .slice(0, 60);
        if (d) return d;
    }
    const y = normStr(cls.year);
    const yy = /^\d{4}$/.test(y) ? y : '';
    const tail = normCode(cls.code)
        .replace(/[^0-9A-Za-z]/g, '')
        .toLowerCase()
        .slice(0, 24);
    if (yy && tail) return ('jg' + yy + tail).toLowerCase().slice(0, 60);
    if (tail) return ('jg' + tail).toLowerCase().slice(0, 60);
    return '';
}

function sanitizeStableNick(raw) {
    return String(raw || '')
        .trim()
        .replace(/[^a-zA-Z0-9]/g, '')
        .toLowerCase()
        .slice(0, 60);
}

/**
 * Gleiche Zuordnung wie im Werkzeug Jahrgangsgruppen (Code + Abschlussjahr, sonst Alias).
 * @param {object|null} cls
 * @param {object[]} classTeams
 * @returns {object|null}
 */
export function findClassTeamForClass(cls, classTeams) {
    if (!cls) return null;
    const teams = Array.isArray(classTeams) ? classTeams : [];
    const code = normCode(cls.code);
    const year = normStr(cls.year);
    if (code) {
        for (let i = 0; i < teams.length; i++) {
            const team = teams[i];
            if (normCode(team && team.classCode) !== code) continue;
            if (year && team.abschlussJahr && String(team.abschlussJahr) !== year) continue;
            return team;
        }
    }
    const nick = deriveClassStableNick(cls);
    if (nick) {
        for (let k = 0; k < teams.length; k++) {
            if (sanitizeStableNick(teams[k] && teams[k].stableMailNickname) === nick) return teams[k];
        }
    }
    return null;
}

function normalizeClassTeamsFromContainer(container) {
    const raw =
        container && container.core && Array.isArray(container.core.classTeams) ? container.core.classTeams : [];
    if (
        typeof globalThis !== 'undefined' &&
        globalThis.ms365AppDataV2 &&
        typeof globalThis.ms365AppDataV2.normalizeCoreClassTeams === 'function'
    ) {
        return globalThis.ms365AppDataV2.normalizeCoreClassTeams(raw);
    }
    return raw;
}

/**
 * @param {object|null} container app-data-v2 container
 * @param {object|null} settings tenant settings snapshot
 * @returns {object[]}
 */
export function buildHygieneTargets(container, settings) {
    const s = settings && typeof settings === 'object' ? settings : {};
    const setup = container && container.setup ? container.setup : {};
    const matched = setup.matched && typeof setup.matched === 'object' ? setup.matched : {};
    const classTeams = normalizeClassTeamsFromContainer(container);
    const classes = Array.isArray(s.classes) ? s.classes : [];
    const students = Array.isArray(s.students) ? s.students : [];
    const targets = [];

    function pushTarget(row) {
        if (!row || !row.id) return;
        targets.push(row);
    }

    const schuelerGid = matched.schuelerGroupId ? String(matched.schuelerGroupId).trim() : '';
    pushTarget({
        id: 'slg-schueler',
        category: 'sammelgruppe',
        label: 'Schüler:innen (Sammelgruppe)',
        groupId: schuelerGid || null,
        listCount: collectStudentEmails(s.students).length,
        toolHref: 'schueler-lehrer-gruppen.html',
        reviewHint: 'Schüler:innen wählen → Mitglieder vergleichen'
    });

    const lehrerGid = matched.lehrerGroupId ? String(matched.lehrerGroupId).trim() : '';
    pushTarget({
        id: 'slg-lehrer',
        category: 'sammelgruppe',
        label: 'Lehrer:innen (Sammelgruppe)',
        groupId: lehrerGid || null,
        listCount: collectTeacherEmails(s.teachers).length,
        toolHref: 'schueler-lehrer-gruppen.html',
        reviewHint: 'Lehrer:innen wählen → Mitglieder vergleichen'
    });

    const vwGid = matched.verwaltungGroupId ? String(matched.verwaltungGroupId).trim() : '';
    pushTarget({
        id: 'verwaltung',
        category: 'sammelgruppe',
        label: 'Verwaltung (Sammelgruppe)',
        groupId: vwGid || null,
        listCount: collectAdminEmails(s.admin).length,
        toolHref: 'verwaltung.html',
        reviewHint: 'Sammelgruppe → Mitglieder vergleichen'
    });

    classes.forEach(function (cls) {
        if (!cls) return;
        const code = normCode(cls.code);
        if (!code) return;
        const team = findClassTeamForClass(cls, classTeams);
        const gid = team && team.graphGroupId ? String(team.graphGroupId).trim() : '';
        const labelParts = [cls.name || cls.code || code];
        if (cls.year) labelParts.push('Abschluss ' + cls.year);
        pushTarget({
            id: 'klasse-' + code,
            category: 'klasse',
            label: 'Klasse ' + labelParts.join(' · '),
            groupId: gid || null,
            listCount: emailsForClassCode(students, code, cls.name).length,
            toolHref: 'jahrgangsgruppen.html',
            reviewHint: 'Klasse wählen → Mitglieder vergleichen',
            classCode: code
        });
    });

    return targets;
}

/**
 * @param {object} target
 * @param {number|null} groupCount
 * @returns {'unmatched'|'unknown'|'ok'|'mismatch'|'empty-list'}
 */
export function hygieneStatusForTarget(target, groupCount) {
    const listN = typeof target.listCount === 'number' ? target.listCount : 0;
    const gid = target.groupId ? String(target.groupId).trim() : '';
    if (!gid) return 'unmatched';
    if (!listN) return 'empty-list';
    if (groupCount === null || groupCount === undefined || groupCount < 0) return 'unknown';
    return listN === groupCount ? 'ok' : 'mismatch';
}

/**
 * @param {object[]} targets
 * @param {Record<string, number>} groupCountsById groupId → count
 * @returns {object}
 */
export function summarizeHygieneScan(targets, groupCountsById) {
    const counts = { ok: 0, mismatch: 0, unmatched: 0, unknown: 0, emptyList: 0, matched: 0 };
    const rows = (Array.isArray(targets) ? targets : []).map(function (t) {
        const gid = t.groupId ? String(t.groupId).trim() : '';
        const groupN =
            gid && groupCountsById && Object.prototype.hasOwnProperty.call(groupCountsById, gid)
                ? groupCountsById[gid]
                : null;
        const status = hygieneStatusForTarget(t, groupN);
        if (gid) counts.matched += 1;
        if (status === 'ok') counts.ok += 1;
        else if (status === 'mismatch') counts.mismatch += 1;
        else if (status === 'unmatched') counts.unmatched += 1;
        else if (status === 'unknown') counts.unknown += 1;
        else if (status === 'empty-list') counts.emptyList += 1;
        return Object.assign({}, t, {
            groupCount: groupN,
            status: status
        });
    });
    return { rows: rows, counts: counts };
}

export function loadHygieneScanCache() {
    try {
        const raw = localStorage.getItem(SCAN_STORAGE_KEY);
        if (!raw) return null;
        const o = JSON.parse(raw);
        if (!o || typeof o !== 'object') return null;
        return o;
    } catch {
        return null;
    }
}

/**
 * @param {object} payload
 */
export function saveHygieneScanCache(payload) {
    try {
        localStorage.setItem(
            SCAN_STORAGE_KEY,
            JSON.stringify(
                Object.assign({}, payload, {
                    savedAt: new Date().toISOString()
                })
            )
        );
    } catch {
        /* ignore */
    }
}

/**
 * @param {object} cfg
 * @param {() => object|null} cfg.loadContainer
 * @param {() => object|null} cfg.loadSettings
 * @param {() => Promise<string>} cfg.getGraphToken
 * @param {(token: string, groupId: string) => Promise<number>} cfg.fetchGroupMemberCount
 */
export async function runHygieneScan(cfg) {
    const loadContainer =
        cfg && typeof cfg.loadContainer === 'function' ? cfg.loadContainer : function () {
            return null;
        };
    const loadSettings =
        cfg && typeof cfg.loadSettings === 'function' ? cfg.loadSettings : function () {
            return null;
        };
    const container = loadContainer();
    const settings = loadSettings();
    const targets = buildHygieneTargets(container, settings);
    const matched = targets.filter(function (t) {
        return t.groupId;
    });
    const groupCountsById = {};
    if (matched.length && cfg && typeof cfg.getGraphToken === 'function' && typeof cfg.fetchGroupMemberCount === 'function') {
        const token = await cfg.getGraphToken();
        await Promise.all(
            matched.map(async function (t) {
                const gid = String(t.groupId || '').trim();
                if (!gid) return;
                try {
                    const n = await cfg.fetchGroupMemberCount(token, gid);
                    groupCountsById[gid] = typeof n === 'number' ? n : -1;
                } catch {
                    groupCountsById[gid] = -1;
                }
            })
        );
    }
    const summary = summarizeHygieneScan(targets, groupCountsById);
    const payload = {
        scannedAt: new Date().toISOString(),
        rows: summary.rows,
        counts: summary.counts
    };
    saveHygieneScanCache(payload);
    return payload;
}

const api = {
    buildHygieneTargets: buildHygieneTargets,
    findClassTeamForClass: findClassTeamForClass,
    hygieneStatusForTarget: hygieneStatusForTarget,
    summarizeHygieneScan: summarizeHygieneScan,
    loadHygieneScanCache: loadHygieneScanCache,
    saveHygieneScanCache: saveHygieneScanCache,
    runHygieneScan: runHygieneScan,
    SCAN_STORAGE_KEY: SCAN_STORAGE_KEY
};

if (typeof window !== 'undefined') {
    window.ms365MembershipHygiene = api;
}

export default api;
