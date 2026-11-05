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
    if (!counts.total) return 'Noch keine Fächer/ARGEs in den Stammdaten';
    const st = subjectCatalogLinkStatus(s, links);
    if (st === 'ok') return counts.linked + ' von ' + counts.total + ' verknüpft';
    if (st === 'unmatched') return 'Noch nicht verknüpft';
    return counts.linked + ' von ' + counts.total + ' verknüpft';
}

/**
 * Klassenkürzel vergleichbar machen (1.A / 1A / 1-A).
 * @param {unknown} raw
 * @returns {string}
 */
export function normClassChatKey(raw) {
    return String(raw || '')
        .trim()
        .toUpperCase()
        .replace(/[\s.\u00A0_\-–—/]+/g, '');
}

/**
 * Ob ein classChats-Eintrag eine echte Teams-Ressource hat (Gruppenchat oder privater Kanal).
 * @param {object|null|undefined} it
 */
export function classChatItemHasResource(it) {
    if (!it || typeof it !== 'object') return false;
    if (String(it.chatId || '').trim()) return true;
    if (String(it.channelId || '').trim()) return true;
    if (String(it.id || '').trim() && String(it.klasse || it.classCode || '').trim()) return true;
    if (String(it.webUrl || '').trim() && String(it.klasse || it.classCode || '').trim()) return true;
    return false;
}

/**
 * Roh-Snapshot → Item-Liste (unterstützt { items }, { chats }, byKlasse-Map und reines Array).
 * @param {unknown} snap
 * @returns {object[]}
 */
export function classChatItemsFromSnapshot(snap) {
    if (!snap) return [];
    if (Array.isArray(snap)) {
        return snap
            .map(function (it) {
                if (!it || typeof it !== 'object') return it;
                if (it.klasse) return it;
                if (it.classCode) return Object.assign({}, it, { klasse: it.classCode });
                return it;
            })
            .filter(classChatItemHasResource);
    }
    if (typeof snap !== 'object') return [];
    let list = Array.isArray(snap.items)
        ? snap.items
        : Array.isArray(snap.chats)
          ? snap.chats
          : null;
    if (!list && snap.byKlasse && typeof snap.byKlasse === 'object') {
        list = Object.keys(snap.byKlasse).map(function (k) {
            const row = snap.byKlasse[k];
            if (!row || typeof row !== 'object') return { klasse: k };
            return Object.assign({}, row, { klasse: row.klasse || row.classCode || k });
        });
    }
    if (!list) {
        const reserved = {
            updatedAt: 1,
            yearPrefix: 1,
            namePattern: 1,
            targetKind: 1,
            lehrerTeamId: 1,
            lehrerTeamDisplayName: 1,
            items: 1,
            chats: 1,
            byKlasse: 1
        };
        const keys = Object.keys(snap).filter(function (k) {
            return !reserved[k];
        });
        const looksLikeMap =
            keys.length > 0 &&
            keys.every(function (k) {
                const v = snap[k];
                return (
                    v &&
                    typeof v === 'object' &&
                    !Array.isArray(v) &&
                    (v.chatId || v.channelId || v.webUrl || v.id || v.topic)
                );
            });
        if (looksLikeMap) {
            list = keys.map(function (k) {
                const row = snap[k];
                return Object.assign({}, row, { klasse: row.klasse || row.classCode || k });
            });
        } else {
            list = [];
        }
    }
    return list
        .map(function (it) {
            if (!it || typeof it !== 'object') return it;
            if (it.klasse) return it;
            if (it.classCode) return Object.assign({}, it, { klasse: it.classCode });
            return it;
        })
        .filter(classChatItemHasResource);
}

/**
 * Chat-/Kanal-IDs aus einem Organigramm-Node.
 * @param {object} n
 * @returns {{ chatId: string, channelId: string, webUrl: string, kind: string }}
 */
function resourceFromOrganigrammNode(n) {
    const meta = n && n.meta && typeof n.meta === 'object' ? n.meta : {};
    const bindings =
        n && n.m365 && n.m365.bindings && typeof n.m365.bindings === 'object' ? n.m365.bindings : {};
    const groupChat = bindings.groupChat && typeof bindings.groupChat === 'object' ? bindings.groupChat : {};
    const privateCh =
        bindings.privateChannel && typeof bindings.privateChannel === 'object'
            ? bindings.privateChannel
            : {};
    const chatId = String(
        meta.chatId || groupChat.id || groupChat.chatId || ''
    ).trim();
    const channelId = String(
        meta.channelId || privateCh.id || privateCh.channelId || ''
    ).trim();
    const webUrl = String(meta.webUrl || meta.chatWebUrl || groupChat.webUrl || privateCh.webUrl || '').trim();
    const kind = channelId && !chatId ? 'privateChannel' : 'groupChat';
    return { chatId: chatId, channelId: channelId, webUrl: webUrl, kind: kind };
}

/**
 * Chat-IDs aus dem Schul-Organigramm (meta.chatId / channelId an KlassenChat-Knoten).
 * @param {object|null} container
 * @param {object|null} [settings]
 * @returns {object[]}
 */
export function collectClassChatItemsFromOrganigramm(container, settings) {
    /** @type {object[]} */
    const orgs = [];
    if (container && container.core && container.core.schulOrganigramm) {
        orgs.push(container.core.schulOrganigramm);
    }
    if (
        settings &&
        settings.schulOrganigramm &&
        (!container ||
            !container.core ||
            settings.schulOrganigramm !== container.core.schulOrganigramm)
    ) {
        orgs.push(settings.schulOrganigramm);
    }
    /** @type {object[]} */
    const out = [];
    const seen = new Set();
    orgs.forEach(function (org) {
        const nodes = org && Array.isArray(org.nodes) ? org.nodes : [];
        nodes.forEach(function (n) {
            if (!n || typeof n !== 'object') return;
            const refKind = String(n.refKind || '')
                .trim()
                .toLowerCase();
            const klasse = String(n.refKey || (n.meta && n.meta.klasse) || '').trim();
            if (!klasse) return;
            if (refKind && refKind !== 'klassenchatclass' && refKind !== 'klassenchat') return;
            const res = resourceFromOrganigrammNode(n);
            if (!res.chatId && !res.channelId && !res.webUrl) return;
            const key = normClassChatKey(klasse) + '|' + res.kind;
            if (seen.has(key)) return;
            seen.add(key);
            out.push({
                klasse: klasse,
                chatId: res.chatId,
                channelId: res.channelId,
                kind: res.kind,
                topic: String((n.meta && (n.meta.chatTopic || n.meta.topic)) || n.label || '').trim(),
                webUrl: res.webUrl
            });
        });
    });
    return out;
}

/**
 * classChats-Items aus Year-Buckets, App-Data-API, Settings und Organigramm.
 * @param {object|null} container
 * @param {object|null} [settings]
 * @returns {{ items: object[], yearLabel: string, fromFallback: boolean }}
 */
export function collectClassChatItems(container, settings) {
    const byLabel =
        container && container.years && container.years.byLabel && typeof container.years.byLabel === 'object'
            ? container.years.byLabel
            : {};
    const current = String((container && container.years && container.years.current) || '').trim();

    function itemsFromBucket(bucket) {
        return classChatItemsFromSnapshot(bucket && bucket.classChats);
    }

    /** @type {object[]} */
    const merged = [];
    const seen = new Set();

    function pushAll(list) {
        (Array.isArray(list) ? list : []).forEach(function (it) {
            if (!classChatItemHasResource(it)) return;
            const klasse = String((it && (it.klasse || it.classCode)) || '').trim();
            if (!klasse) return;
            const kind = String(
                (it && it.kind) || (it && it.channelId ? 'privateChannel' : 'groupChat')
            );
            const key = normClassChatKey(klasse) + '|' + kind;
            if (seen.has(key)) return;
            seen.add(key);
            merged.push(Object.assign({}, it, { klasse: klasse }));
        });
    }

    // 1) App-Data-API (gleicher Pfad wie Werkzeug KlassenChats)
    try {
        if (
            typeof window !== 'undefined' &&
            window.ms365AppDataV2 &&
            typeof window.ms365AppDataV2.getClassChats === 'function'
        ) {
            const labels = Object.keys(byLabel);
            if (current && labels.indexOf(current) === -1) labels.unshift(current);
            if (!labels.length) labels.push('');
            labels.forEach(function (lab) {
                pushAll(classChatItemsFromSnapshot(window.ms365AppDataV2.getClassChats(lab || undefined)));
            });
        }
    } catch {
        /* ignore */
    }

    // 2) Container Year-Buckets
    if (current && byLabel[current]) pushAll(itemsFromBucket(byLabel[current]));
    Object.keys(byLabel).forEach(function (label) {
        if (label === current) return;
        pushAll(itemsFromBucket(byLabel[label]));
    });

    // 3) Settings (Tenant-Load / Organigramm reichern settings.classChats an)
    if (settings && settings.classChats) {
        pushAll(classChatItemsFromSnapshot(settings.classChats));
    }

    // 4) Organigramm-Meta (container + settings)
    pushAll(collectClassChatItemsFromOrganigramm(container, settings));

    // 5) Roh-localStorage (falls In-Memory-Cache und Year-Buckets auseinanderlaufen)
    if (!merged.length) {
        try {
            if (typeof localStorage !== 'undefined') {
                const raw = localStorage.getItem('ms365-schooltool-data-v2');
                const parsed = raw ? JSON.parse(raw) : null;
                const rawBy =
                    parsed && parsed.years && parsed.years.byLabel && typeof parsed.years.byLabel === 'object'
                        ? parsed.years.byLabel
                        : {};
                Object.keys(rawBy).forEach(function (lab) {
                    pushAll(itemsFromBucket(rawBy[lab]));
                });
                // Separater Klassenchats-Spiegel (überlebt Year-Bucket-Verluste)
                const mirrorRaw = localStorage.getItem('ms365-class-chats-mirror-v1');
                const mirror = mirrorRaw ? JSON.parse(mirrorRaw) : null;
                const mirrorBy =
                    mirror && mirror.byYear && typeof mirror.byYear === 'object' ? mirror.byYear : {};
                Object.keys(mirrorBy).forEach(function (lab) {
                    pushAll(classChatItemsFromSnapshot(mirrorBy[lab]));
                });
            }
        } catch {
            /* ignore */
        }
    }

    const fromFallback = merged.length > 0 && !(current && itemsFromBucket(byLabel[current]).length);
    return {
        items: merged,
        yearLabel: current,
        fromFallback: fromFallback
    };
}

/**
 * @param {object|null} container
 * @param {object|null} settings
 * @returns {{ total: number, withChat: number, resourceCount: number, classKeys: number, yearLabel: string, fromFallback: boolean }}
 */
export function klassenChatsCoverage(container, settings) {
    const classes = (settings && settings.classes) || [];
    const collected = collectClassChatItems(container, settings);
    const items = collected.items;
    const chatKeys = new Set();
    items.forEach(function (it) {
        const k = normClassChatKey((it && it.klasse) || '');
        if (k) chatKeys.add(k);
    });
    let withChat = 0;
    classes.forEach(function (cls) {
        if (!cls) return;
        const code = normClassChatKey(cls.code || cls.name || '');
        if (code && chatKeys.has(code)) withChat += 1;
    });
    return {
        total: classes.length,
        withChat: withChat,
        resourceCount: items.length,
        classKeys: chatKeys.size,
        yearLabel: collected.yearLabel,
        fromFallback: collected.fromFallback
    };
}

/**
 * @param {object|null} container
 * @param {object|null} settings
 * @returns {'ok'|'mismatch'|'unmatched'|null}
 */
export function klassenChatsProvisionStatus(container, settings) {
    const classes = (settings && settings.classes) || [];
    if (!classes.length) return null;
    const cov = klassenChatsCoverage(container, settings);
    // Ressourcen vorhanden, aber Klassenkürzel passen nicht → trotzdem nicht „keine Chats“
    if (cov.resourceCount === 0) return 'unmatched';
    if (cov.withChat >= cov.total) return 'ok';
    return 'mismatch';
}

/**
 * @param {object|null} container
 * @param {object|null} settings
 * @returns {string}
 */
export function klassenChatsStatusHint(container, settings) {
    const classes = (settings && settings.classes) || [];
    if (!classes.length) return 'Zuerst Klassen in den Stammdaten';
    const cov = klassenChatsCoverage(container, settings);
    const st = klassenChatsProvisionStatus(container, settings);
    if (st === 'unmatched') return 'Noch keine Chats angelegt';
    if (cov.withChat === 0 && cov.resourceCount > 0) {
        const n = cov.classKeys || cov.resourceCount;
        return n + ' Chat' + (n === 1 ? '' : 's') + ' angelegt';
    }
    return cov.withChat + ' von ' + cov.total + ' Klassen mit Chat';
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
    const setup = container && container.setup ? container.setup : {};
    const classGroupMatchByKey =
        setup.classGroupMatchByKey && typeof setup.classGroupMatchByKey === 'object'
            ? setup.classGroupMatchByKey
            : {};
    const hygieneApi = typeof window !== 'undefined' ? window.ms365MembershipHygiene : null;
    let matched = 0;
    if (hygieneApi && typeof hygieneApi.countLinkedClassTeamsForClasses === 'function') {
        matched = hygieneApi.countLinkedClassTeamsForClasses(classes, classTeams, classGroupMatchByKey)
            .linked;
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

function readSpoSyncAtForStatus() {
    try {
        if (
            typeof window !== 'undefined' &&
            window.ms365StammdatenSpoAutoSync &&
            typeof window.ms365StammdatenSpoAutoSync.getStatus === 'function'
        ) {
            const st = window.ms365StammdatenSpoAutoSync.getStatus();
            if (st && st.lastAt) return String(st.lastAt);
        }
    } catch {
        /* ignore */
    }
    try {
        let tid = '';
        if (typeof window !== 'undefined' && typeof window.ms365AuthGetAccountInfo === 'function') {
            const info = window.ms365AuthGetAccountInfo();
            tid = info && info.tenantId ? String(info.tenantId).trim() : '';
        }
        if (tid && typeof localStorage !== 'undefined') {
            const map = JSON.parse(localStorage.getItem('ms365-stammdaten-spo-sync-by-tenant-v2') || '{}');
            if (map && map[tid] && map[tid].at) return String(map[tid].at);
        }
    } catch {
        /* ignore */
    }
    try {
        if (typeof localStorage === 'undefined') return '';
        const raw = localStorage.getItem('ms365-stammdaten-spo-sync-v1');
        const m = raw ? JSON.parse(raw) : null;
        return m && m.at ? String(m.at) : '';
    } catch {
        return '';
    }
}

function sharepointListStatus(container) {
    const setup = container && container.setup ? container.setup : {};
    const url =
        (setup.intranetSiteUrl && String(setup.intranetSiteUrl).trim()) ||
        (setup.schoolIntranetSiteUrl && String(setup.schoolIntranetSiteUrl).trim()) ||
        '';
    if (!url) {
        return {
            tone: 'pending',
            primary: 'Intranet-URL fehlt',
            secondary: 'In den Stammdaten oder Intranet-Hub setzen'
        };
    }
    let secondary = '';
    const at = readSpoSyncAtForStatus();
    if (at) secondary = formatSyncSecondary(at);
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
        const hint = klassenChatsStatusHint(container, s);
        if (st === 'ok') {
            return { tone: 'ok', primary: hint, secondary: '' };
        }
        if (st === 'mismatch') {
            return {
                tone: 'warn',
                primary: hint,
                secondary: ''
            };
        }
        if (st === 'unmatched') {
            return { tone: 'unmatched', primary: hint || 'Noch keine Chats angelegt', secondary: '' };
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
