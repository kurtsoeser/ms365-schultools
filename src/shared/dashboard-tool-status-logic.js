/**
 * Live-Status für Dashboard (Aufgaben-Kacheln + Werkzeugkatalog).
 */
import { loadPlaybookState } from './playbook-store.js';
import {
    DASHBOARD_PLAYBOOKS,
    computePlaybookProgress,
    formatPlaybookProgressLabel
} from './dashboard-playbooks-catalog.js';
import {
    SCHULJAHRSTART_PLAYBOOK_STORAGE_KEY,
    schuljahresstartPlaybookProgress,
    tenantCountsFromSettings
} from './playbook-schuljahresstart-gates.js';

/** @typedef {{ tone: string, primary: string, secondary?: string }} DashboardToolStatus */

const HYGIENE_TOOL_ID = {
    'slg-schueler': 'slg-schueler',
    'slg-lehrer': 'slg-lehrer',
    verwaltung: 'verwaltung',
    klassenvorstaende: 'klassenvorstaende'
};

const PLAYBOOK_BY_TOOL_ID = Object.fromEntries(
    DASHBOARD_PLAYBOOKS.map(function (p) {
        const toolId = p.href.replace(/^tools\//, '').replace(/\.html$/, '');
        return [toolId, p];
    })
);

const SHAREPOINT_LIST_TOOLS = new Set([
    'sharepoint-intranet-hub',
    'sharepoint-liste-lehrer',
    'sharepoint-liste-stammdaten',
    'sharepoint-liste-schultermine',
    'sharepoint-liste-schularbeiten',
    'sharepoint-liste-schulaktivitaeten',
    'sharepoint-liste-projektwochen',
    'sharepoint-liste-srdp',
    'sharepoint-liste-vertretung',
    'playbook-intranet'
]);

export function subjectCatalogLinkStatus(s, links) {
    const subjectRows = (s && s.subjects) || [];
    const argeRows = (s && s.arges) || [];
    const items = [];
    subjectRows.forEach(function (row) {
        const code = row && typeof row === 'object' ? row.code : row;
        if (String(code || '').trim()) {
            items.push({ kind: 'subject', code: String(code).trim().toLowerCase() });
        }
    });
    argeRows.forEach(function (row) {
        const code = row && typeof row === 'object' ? row.code : row;
        if (String(code || '').trim()) {
            items.push({ kind: 'arge', code: String(code).trim().toLowerCase() });
        }
    });
    if (!items.length) return null;
    let linked = 0;
    items.forEach(function (item) {
        const hit = (Array.isArray(links) ? links : []).some(function (l) {
            return (
                l &&
                l.kind === item.kind &&
                l.graphGroupId &&
                String(l.code || '')
                    .trim()
                    .toLowerCase() === item.code
            );
        });
        if (hit) linked += 1;
    });
    if (linked === items.length) return 'ok';
    if (linked === 0) return 'unmatched';
    return 'mismatch';
}

/** @param {object|null} s @param {object[]} links */
export function subjectCatalogLinkCounts(s, links) {
    const subjectRows = (s && s.subjects) || [];
    const argeRows = (s && s.arges) || [];
    const items = [];
    subjectRows.forEach(function (row) {
        const code = row && typeof row === 'object' ? row.code : row;
        if (String(code || '').trim()) {
            items.push({ kind: 'subject', code: String(code).trim().toLowerCase() });
        }
    });
    argeRows.forEach(function (row) {
        const code = row && typeof row === 'object' ? row.code : row;
        if (String(code || '').trim()) {
            items.push({ kind: 'arge', code: String(code).trim().toLowerCase() });
        }
    });
    let linked = 0;
    items.forEach(function (item) {
        const hit = (Array.isArray(links) ? links : []).some(function (l) {
            return (
                l &&
                l.kind === item.kind &&
                l.graphGroupId &&
                String(l.code || '')
                    .trim()
                    .toLowerCase() === item.code
            );
        });
        if (hit) linked += 1;
    });
    return { linked: linked, total: items.length };
}

/**
 * @param {object|null} s
 * @param {object[]} links
 * @returns {string}
 */
export function subjectCatalogStatusHint(s, links) {
    const counts = subjectCatalogLinkCounts(s, links);
    if (!counts.total) return 'Noch keine Fächer/ARGEs im Register';
    const st = subjectCatalogLinkStatus(s, links);
    if (st === 'ok') return counts.linked + ' von ' + counts.total + ' verknüpft';
    if (st === 'unmatched') return 'Noch nicht verknüpft';
    return counts.linked + ' von ' + counts.total + ' verknüpft';
}

/**
 * @param {object|null} container
 * @param {object|null} settings
 * @returns {'ok'|'mismatch'|'unmatched'|null}
 */
export function klassenChatsProvisionStatus(container, settings) {
    const classes = (settings && settings.classes) || [];
    if (!classes.length) return null;
    const current = String((container && container.years && container.years.current) || '').trim();
    const byLabel =
        container && container.years && container.years.byLabel && typeof container.years.byLabel === 'object'
            ? container.years.byLabel
            : {};
    const bucket = current && byLabel[current] ? byLabel[current] : null;
    const items =
        bucket && bucket.classChats && Array.isArray(bucket.classChats.items) ? bucket.classChats.items : [];
    const chatKeys = new Set();
    items.forEach(function (it) {
        const k = String((it && it.klasse) || '')
            .trim()
            .toUpperCase();
        if (k) chatKeys.add(k);
    });
    let withChat = 0;
    classes.forEach(function (cls) {
        if (!cls) return;
        const code = String(cls.code || cls.name || '')
            .trim()
            .toUpperCase();
        if (code && chatKeys.has(code)) withChat += 1;
    });
    const total = classes.length;
    if (withChat >= total) return 'ok';
    if (withChat === 0) return 'unmatched';
    return 'mismatch';
}

/**
 * @param {object|null} container
 * @param {object|null} settings
 * @returns {string}
 */
export function klassenChatsStatusHint(container, settings) {
    const classes = (settings && settings.classes) || [];
    if (!classes.length) return 'Zuerst Klassen im Register';
    const st = klassenChatsProvisionStatus(container, settings);
    const current = String((container && container.years && container.years.current) || '').trim();
    const byLabel =
        container && container.years && container.years.byLabel && typeof container.years.byLabel === 'object'
            ? container.years.byLabel
            : {};
    const bucket = current && byLabel[current] ? byLabel[current] : null;
    const items =
        bucket && bucket.classChats && Array.isArray(bucket.classChats.items) ? bucket.classChats.items : [];
    const chatKeys = new Set();
    items.forEach(function (it) {
        const k = String((it && it.klasse) || '')
            .trim()
            .toUpperCase();
        if (k) chatKeys.add(k);
    });
    let withChat = 0;
    classes.forEach(function (cls) {
        if (!cls) return;
        const code = String(cls.code || cls.name || '')
            .trim()
            .toUpperCase();
        if (code && chatKeys.has(code)) withChat += 1;
    });
    const total = classes.length;
    if (st === 'ok') return withChat + ' von ' + total + ' Klassen mit Chat';
    if (st === 'unmatched') return 'Noch keine Chats angelegt';
    return withChat + ' von ' + total + ' Klassen mit Chat';
}

export function scanSummaryHygieneStatus(hygieneApi) {
    if (!hygieneApi || typeof hygieneApi.loadHygieneScanCache !== 'function') return null;
    const cache = hygieneApi.loadHygieneScanCache();
    if (!cache || !cache.counts) return null;
    const c = cache.counts;
    if ((c.mismatch || 0) > 0 || (c.emptyList || 0) > 0) return 'mismatch';
    if ((c.unmatched || 0) > 0) return 'unmatched';
    if ((c.ok || 0) > 0 && !(c.unknown || 0)) return 'ok';
    if ((c.unknown || 0) > 0) return 'unknown';
    return null;
}

export function formatHygieneSyncSecondary(hygieneApi) {
    if (!hygieneApi || typeof hygieneApi.loadHygieneScanCache !== 'function') return '';
    const cache = hygieneApi.loadHygieneScanCache();
    const iso = cache && (cache.scannedAt || cache.savedAt) ? String(cache.scannedAt || cache.savedAt) : '';
    return formatSyncSecondary(iso);
}

export function formatSyncSecondary(iso) {
    if (!iso) return '';
    try {
        const d = new Date(iso);
        if (isNaN(d.getTime())) return '';
        const now = new Date();
        const sameDay =
            d.getFullYear() === now.getFullYear() &&
            d.getMonth() === now.getMonth() &&
            d.getDate() === now.getDate();
        if (sameDay) return 'Letzter Sync: heute';
        return (
            'Letzter Sync: ' +
            d.toLocaleString('de-AT', {
                day: '2-digit',
                month: '2-digit',
                year: 'numeric',
                hour: '2-digit',
                minute: '2-digit'
            })
        );
    } catch {
        return '';
    }
}

function hygieneLabel(status, hygieneApi) {
    if (!status || !hygieneApi) return null;
    const primary = hygieneApi.hygieneStatusDashboardHint(status);
    if (!primary) return null;
    const tone = hygieneApi.hygieneStatusDashboardTone(status);
    const secondary =
        status === 'ok' || status === 'mismatch' ? formatHygieneSyncSecondary(hygieneApi) : '';
    return { tone, primary, secondary };
}

function classTeamsProgress(container, settings) {
    const classes = (settings && settings.classes) || [];
    const classTeams =
        container && container.core && Array.isArray(container.core.classTeams)
            ? container.core.classTeams
            : [];
    const hygieneApi = typeof window !== 'undefined' ? window.ms365MembershipHygiene : null;
    let matched = 0;
    if (hygieneApi && typeof hygieneApi.countLinkedClassTeamsForClasses === 'function') {
        matched = hygieneApi.countLinkedClassTeamsForClasses(classes, classTeams).linked;
    }
    const total = classes.length;
    if (!total) return null;
    const tone = matched >= total ? 'ok' : matched ? 'warn' : 'pending';
    return {
        tone,
        primary: matched + ' von ' + total + ' Klassengruppen verknüpft',
        secondary: ''
    };
}

function playbookStatusForDef(def) {
    const st = loadPlaybookState(def.storageKey);
    const prog = computePlaybookProgress(st, def.stepIds);
    const label = formatPlaybookProgressLabel(prog);
    let tone = 'pending';
    if (prog.status === 'complete') tone = 'ok';
    else if (prog.status === 'in-progress') tone = 'warn';
    else tone = 'pending';
    return { tone, primary: label, secondary: prog.status === 'in-progress' ? 'Playbook fortsetzen' : '' };
}

function sharepointListStatus(container) {
    const setup = container && container.setup ? container.setup : {};
    const url = setup.intranetSiteUrl ? String(setup.intranetSiteUrl).trim() : '';
    if (!url) {
        return {
            tone: 'pending',
            primary: 'Intranet-URL fehlt',
            secondary: 'Im Schulregister oder Intranet-Hub setzen'
        };
    }
    let secondary = '';
    try {
        const raw = localStorage.getItem('ms365-stammdaten-spo-sync-v1');
        const m = raw ? JSON.parse(raw) : null;
        if (m && m.at) secondary = formatSyncSecondary(String(m.at));
    } catch {
        /* ignore */
    }
    return {
        tone: 'ok',
        primary: 'Intranet verbunden',
        secondary: secondary || 'Listen im Werkzeug abgleichen'
    };
}

/**
 * @param {string} toolId
 * @param {{
 *   container: object|null,
 *   settings: object|null,
 *   hygieneById: Record<string, string>|null,
 *   hygieneApi: object|null,
 *   show: boolean
 * }} ctx
 * @returns {DashboardToolStatus|null}
 */
export function resolveDashboardToolStatus(toolId, ctx) {
    if (!ctx.show || !toolId) return null;
    const id = String(toolId).trim();
    const hygieneApi = ctx.hygieneApi;
    const byId = ctx.hygieneById || {};
    const s = ctx.settings;
    const container = ctx.container;

    const hygieneKey = HYGIENE_TOOL_ID[id];
    if (hygieneKey && hygieneApi) {
        return hygieneLabel(byId[hygieneKey] || null, hygieneApi);
    }

    if (id === 'jahrgang' || id === 'kursteams' || id === 'unterrichtsteams-katalog') {
        return classTeamsProgress(container, s);
    }

    if (id === 'arge-fachgruppen' && hygieneApi) {
        const links =
            container && container.setup && Array.isArray(container.setup.catalogLinks)
                ? container.setup.catalogLinks
                : [];
        const st = subjectCatalogLinkStatus(s, links);
        const base = hygieneLabel(st, hygieneApi);
        if (base) return base;
        if (st === 'ok') return { tone: 'ok', primary: 'Fächer/ARGEs verknüpft', secondary: '' };
        return { tone: 'pending', primary: 'Fächer/ARGEs noch nicht verknüpft', secondary: '' };
    }

    if (id === 'datenhygiene' && hygieneApi) {
        const st = scanSummaryHygieneStatus(hygieneApi);
        const base = hygieneLabel(st, hygieneApi);
        if (base) return base;
        return { tone: 'pending', primary: 'Noch nicht geprüft', secondary: 'Mit Microsoft 365 abgleichen' };
    }

    if (id === 'klassenchats') {
        const st = klassenChatsProvisionStatus(container, s);
        if (st === 'ok') {
            return { tone: 'ok', primary: 'Konsistent', secondary: '' };
        }
        if (st === 'mismatch') {
            return {
                tone: 'warn',
                primary: klassenChatsStatusHint(container, s),
                secondary: ''
            };
        }
        if (st === 'unmatched') {
            return { tone: 'unmatched', primary: 'Noch keine Chats', secondary: '' };
        }
        return null;
    }

    const playbookDef = PLAYBOOK_BY_TOOL_ID[id];
    if (playbookDef) {
        return playbookStatusForDef(playbookDef);
    }

    if (id === 'organisations-assistent' && typeof window !== 'undefined' && window.ms365Playbook) {
        const st = loadPlaybookState(SCHULJAHRSTART_PLAYBOOK_STORAGE_KEY);
        const counts = tenantCountsFromSettings(s);
        const prog = schuljahresstartPlaybookProgress(st);
        if (prog.done > 0 || counts.classes > 0) {
            return {
                tone: prog.done >= prog.total ? 'ok' : 'warn',
                primary: prog.done + ' von ' + prog.total + ' Playbook-Schritte',
                secondary: ''
            };
        }
    }

    if (SHAREPOINT_LIST_TOOLS.has(id)) {
        return sharepointListStatus(container);
    }

    if (id === 'webuntis-sync-monitor') {
        return classTeamsProgress(container, s);
    }

    return null;
}
