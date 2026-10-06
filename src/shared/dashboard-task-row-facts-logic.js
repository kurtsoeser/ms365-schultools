/**
 * Logik für Kennzahlen auf Dashboard-Aufgaben-Kacheln.
 */
import { loadPlaybookState } from './playbook-store.js';
import {
    klassenChatsProvisionStatus,
    klassenChatsStatusHint,
    scanSummaryHygieneStatus,
    subjectCatalogLinkCounts,
    subjectCatalogLinkStatus,
} from './dashboard-tool-status-logic.js';
import {
    SCHULJAHRSTART_PLAYBOOK_STORAGE_KEY,
    schuljahresstartPlaybookProgress,
    tenantCountsFromSettings
} from './playbook-schuljahresstart-gates.js';

/** @typedef {{ tone?: string, chip?: string, desc?: string }} TaskRowFacts */

/**
 * @param {object|null} settings
 */
export function registerSnapshotLine(settings) {
    const s = settings && typeof settings === 'object' ? settings : {};
    const subjects = ((s.subjects) || []).length;
    const teachers = ((s.teachers) || []).length;
    const students = ((s.students) || []).length;
    const classes = ((s.classes) || []).length;
    const parts = [];
    if (teachers) parts.push(teachers + ' Lehrkräfte');
    if (students) parts.push(students + ' Schüler:innen');
    if (classes) parts.push(classes + ' Klassen');
    if (subjects) parts.push(subjects + ' Fächer');
    return parts.join(' · ');
}

/**
 * @param {object|null} container
 * @param {object|null} settings
 * @param {object|null} hygieneApi
 */
export function classTeamsLinkedCounts(container, settings, hygieneApi) {
    const classes = (settings && settings.classes) || [];
    const classTeams =
        container && container.core && Array.isArray(container.core.classTeams)
            ? container.core.classTeams
            : [];
    const total = classes.length;
    if (!total) return { linked: 0, total: 0 };
    let linked = 0;
    if (hygieneApi && typeof hygieneApi.countLinkedClassTeamsForClasses === 'function') {
        linked = hygieneApi.countLinkedClassTeamsForClasses(classes, classTeams).linked;
    }
    return { linked: linked, total: total };
}

/**
 * @param {string|null|undefined} status
 * @param {object|null} hygieneApi
 */
export function hygieneStateChip(status, hygieneApi) {
    if (!status || !hygieneApi) return '';
    if (status === 'ok') return 'Konsistent';
    if (status === 'mismatch' || status === 'empty-list') return 'Abweichung';
    if (status === 'unmatched') return 'Offen';
    if (status === 'unknown') return 'Prüfen';
    return '';
}

/**
 * Zahlenzeile (Mitte der Kachel) – ohne Statuswort.
 * @returns {string}
 */
export function hygieneTargetMetricLine(targetId, status, container, settings, hygieneApi) {
    if (!hygieneApi || typeof hygieneApi.buildHygieneTargets !== 'function') return '';
    const targets = hygieneApi.buildHygieneTargets(container, settings);
    const t = targets.find(function (row) {
        return row && row.id === targetId;
    });
    if (!t) return '';

    const listN = typeof t.listCount === 'number' ? t.listCount : 0;
    let groupN = null;
    if (typeof hygieneApi.loadHygieneScanCache === 'function') {
        const cache = hygieneApi.loadHygieneScanCache();
        const rows = cache && Array.isArray(cache.rows) ? cache.rows : [];
        const hit = rows.find(function (r) {
            return r && r.id === targetId;
        });
        if (hit && typeof hit.groupCount === 'number' && hit.groupCount >= 0) groupN = hit.groupCount;
    }

    if (!t.groupId) {
        return listN ? listN + ' im Register · Gruppe fehlt' : 'Noch nicht verknüpft';
    }
    if (groupN === null) {
        return listN ? listN + ' im Register · Abgleich offen' : 'Abgleich offen';
    }
    return listN + ' Register · ' + groupN + ' M365';
}

/** @deprecated Kompatibilität – volle Zeile; neue UI nutzt metric + chip getrennt */
export function hygieneTargetNumericHint(targetId, status, container, settings, hygieneApi) {
    const metric = hygieneTargetMetricLine(targetId, status, container, settings, hygieneApi);
    const chip = hygieneStateChip(status, hygieneApi);
    if (metric && chip) return metric + ' · ' + chip;
    return metric || chip;
}

function toneFromHygieneStatus(status, hygieneApi) {
    if (!status || !hygieneApi) return '';
    return hygieneApi.hygieneStatusDashboardTone(status);
}

/**
 * @param {object} opts
 * @returns {TaskRowFacts|null}
 */
export function resolveAggregateTaskRowFacts(opts) {
    const o = opts || {};
    const hygieneApi = o.hygieneApi;
    const settings = o.settings;
    const container = o.container;
    const byId = o.hygieneById || {};
    const links = o.links || [];

    if (o.hygieneId) {
        const st = byId[o.hygieneId] || null;
        const desc = hygieneTargetMetricLine(o.hygieneId, st, container, settings, hygieneApi);
        if (!desc) return null;
        return {
            tone: toneFromHygieneStatus(st, hygieneApi),
            chip: hygieneStateChip(st, hygieneApi),
            desc: desc
        };
    }

    if (o.aggregate === 'klassen') {
        const counts = classTeamsLinkedCounts(container, settings, hygieneApi);
        if (!counts.total) {
            return {
                tone: 'pending',
                chip: 'Offen',
                desc: 'Noch keine Klassen im Register'
            };
        }
        const tone = counts.linked >= counts.total ? 'ok' : counts.linked ? 'warn' : 'pending';
        return {
            tone: tone,
            chip: hygieneStateChip(counts.linked >= counts.total ? 'ok' : 'mismatch', hygieneApi) || 'Offen',
            desc: counts.linked + ' von ' + counts.total + ' Klassengruppen'
        };
    }

    if (o.aggregate === 'subjects') {
        const st = subjectCatalogLinkStatus(settings, links);
        const counts = subjectCatalogLinkCounts(settings, links);
        if (!counts.total) {
            return {
                tone: 'pending',
                chip: '',
                desc: 'Fächer/ARGEs im Register anlegen'
            };
        }
        return {
            tone: toneFromHygieneStatus(st, hygieneApi),
            chip: hygieneStateChip(st, hygieneApi),
            desc: counts.linked + ' von ' + counts.total + ' Fächer/ARGEs verknüpft'
        };
    }

    if (o.aggregate === 'klassenchats') {
        const st = klassenChatsProvisionStatus(container, settings);
        const classes = ((settings && settings.classes) || []).length;
        if (!classes) {
            return { tone: 'pending', chip: '', desc: 'Klassen im Register fehlen' };
        }
        const hint = klassenChatsStatusHint(container, settings);
        return {
            tone: toneFromHygieneStatus(st, hygieneApi),
            chip: hygieneStateChip(st, hygieneApi),
            desc: hint
        };
    }

    if (o.aggregate === 'scan') {
        const st = scanSummaryHygieneStatus(hygieneApi);
        const cache = hygieneApi && hygieneApi.loadHygieneScanCache ? hygieneApi.loadHygieneScanCache() : null;
        const c = cache && cache.counts ? cache.counts : {};
        const parts = [];
        if (c.ok) parts.push(c.ok + ' ok');
        if (c.mismatch) parts.push(c.mismatch + ' Abweichung');
        if (c.unmatched) parts.push(c.unmatched + ' offen');
        return {
            tone: toneFromHygieneStatus(st, hygieneApi),
            chip: parts.length ? hygieneStateChip(st, hygieneApi) : 'Prüfen',
            desc: parts.length ? parts.join(' · ') : 'Noch nicht geprüft'
        };
    }

    return null;
}

/**
 * @param {object} opts
 * @returns {TaskRowFacts|null}
 */
export function resolveToolTaskRowFacts(opts) {
    const toolId = String((opts && opts.toolId) || '').trim();
    const settings = opts && opts.settings;
    const register = registerSnapshotLine(settings);

    if (toolId === 'personen-verwaltung') {
        return {
            tone: '',
            chip: '',
            desc: register || 'Stammdaten im Schulregister'
        };
    }

    if (toolId === 'schueler-lifecycle') {
        const n = ((settings && settings.students) || []).length;
        return {
            tone: n ? 'ok' : 'pending',
            chip: n ? '' : 'Offen',
            desc: n ? n + ' Schüler:innen im Register' : 'Noch keine Schüler:innen'
        };
    }

    if (toolId === 'gaeste-verwalten') {
        const teachers = ((settings && settings.teachers) || []).length;
        const students = ((settings && settings.students) || []).length;
        return {
            tone: '',
            chip: '',
            desc:
                teachers + students
                    ? teachers + students + ' interne Konten · Gäste in M365 prüfen'
                    : 'Register befüllen, dann Gäste prüfen'
        };
    }

    if (toolId === 'lizenzverwaltung') {
        const teachers = ((settings && settings.teachers) || []).length;
        return {
            tone: '',
            chip: '',
            desc: teachers ? teachers + ' Lehrkräfte im Register' : 'Lehrkräfte-Stamm pflegen'
        };
    }

    if (toolId === 'namenskonvention-audit') {
        const teachers = ((settings && settings.teachers) || []).length;
        const students = ((settings && settings.students) || []).length;
        const n = teachers + students;
        return {
            tone: '',
            chip: '',
            desc: n ? n + ' Namen aus dem Register prüfbar' : 'Register befüllen'
        };
    }

    if (toolId === 'kursteams' || toolId === 'unterrichtsteams-katalog') {
        const classes = ((settings && settings.classes) || []).length;
        const subjects = ((settings && settings.subjects) || []).length;
        const parts = [];
        if (classes) parts.push(classes + ' Klassen');
        if (subjects) parts.push(subjects + ' Fächer');
        return {
            tone: '',
            chip: '',
            desc: parts.length ? parts.join(' · ') + ' im Register' : 'Stammdaten für Kursteams'
        };
    }

    if (toolId === 'playbook-schuljahresstart') {
        const st = loadPlaybookState(SCHULJAHRSTART_PLAYBOOK_STORAGE_KEY);
        const counts = tenantCountsFromSettings(settings);
        const prog = schuljahresstartPlaybookProgress(st);
        if (prog.done > 0 || counts.classes > 0) {
            return {
                tone: prog.done >= prog.total ? 'ok' : 'warn',
                chip: prog.done >= prog.total ? 'Fertig' : 'Offen',
                desc: prog.done + ' von ' + prog.total + ' Schritte'
            };
        }
        return null;
    }

    if (toolId === 'organisations-assistent') {
        const counts = tenantCountsFromSettings(settings);
        if (counts.classes || counts.students) {
            return {
                tone: 'pending',
                chip: '',
                desc: counts.classes + ' Klassen · ' + counts.students + ' Schüler:innen'
            };
        }
    }

    if (toolId === 'datenlandkarte') {
        return {
            tone: '',
            chip: '',
            desc: register || 'Datenquellen der Schule'
        };
    }

    return null;
}
