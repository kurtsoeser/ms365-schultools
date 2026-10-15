/**
 * Schul-Organigramm UI (Phase 1–6: Pflege, Migration, Filter/Druck/DnD).
 */
import {
    childrenOf,
    normalizeSchulOrganigramm,
    seedSchulOrganigrammFromStammdaten,
    resolveSchulOrganigrammLayout,
    migrateVerwaltungOrganigrammToSchul,
    hasVerwaltungOrganigrammLayout,
    schulOrganigrammStats,
    moveSchulOrganigrammNode,
    reorderSchulOrganigrammSections,
    nodeMatchesSchulOrgFilter,
    rebuildSchulOrganigrammFromStammdaten,
    applySchulOrganigrammToStammdaten,
    isRoleCapableSection,
    audienceIdFromSectionId,
    canDeleteSchulOrganigrammArea,
    isAreaNode,
    isArgeAreaNode,
    isArgeNode,
    isClassAreaNode,
    isClassNode,
    sammelgruppeKeyForArea,
    isSgaBoardArea,
    isSgaScopeArea,
    isJahrgangArea,
    mergeSchulOrganigrammLivePositions,
    SCHUL_ORG_BOARD_REF,
    SCHUL_ORG_SEC_LEHRER,
    SCHUL_ORG_SEC_SCHUELER,
    NODE,
    schulOrgGroupKindLabel,
    personHasM365UserLink,
    nodeM365LinkFlags
} from './schul-organigramm-logic.js';
import {
    createSchulOrganigrammEditor,
    previewPeopleForSchulOrgNode,
    SCHUL_ORG_PEOPLE_PREVIEW_LIMIT,
    findClass,
    findArge,
    findAdminRole,
    listPeopleForSchulOrgNode
} from './schul-organigramm-edit.js';
import { mountSchulOrganigrammLandkarte } from './schul-organigramm-canvas.js';
import { openSchulOrgProvisionDiffPanel } from './schul-organigramm-provision-ui.js';
import {
    collectSchulOrgPersonMemberships,
    searchSchulOrgDirectoryPeople
} from './schul-organigramm-person-focus.js';
import {
    mergeGraphFocusPeople,
    nodeUsesGraphFocusPeople,
    ensureLeadSeat
} from './schul-organigramm-focus-people.js';
import { resolveClassGraphGroupId } from '../../shared/membership-hygiene.js';
import '../../shared/schul-organigramm-stammdaten-sync.js';

const VIEW_MODE_KEY = 'ms365-schul-org-view-mode-v1';

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function loadSettings() {
    let settings = {};
    if (typeof window.ms365TenantSettingsLoad === 'function') {
        settings = window.ms365TenantSettingsLoad() || {};
    }
    // Kursteams / KlassenChats liegen in years.* (App-Data v2)
    try {
        const api = window.ms365AppDataV2;
        if (!settings.unterrichtsbelegung && api && typeof api.getUnterrichtsbelegung === 'function') {
            const snap = api.getUnterrichtsbelegung();
            if (snap && Array.isArray(snap.rows) && snap.rows.length) {
                settings = Object.assign({}, settings, { unterrichtsbelegung: snap });
            }
        }
        if (!settings.classChats && api && typeof api.getClassChats === 'function') {
            const chats = api.getClassChats();
            if (chats && typeof chats === 'object') {
                settings = Object.assign({}, settings, { classChats: chats });
            }
        }
    } catch {
        /* ignore */
    }
    return settings;
}

function getAppContainer() {
    if (typeof window.ms365AppDataV2?.getContainer === 'function') {
        return window.ms365AppDataV2.getContainer() || null;
    }
    return null;
}

function personInitials(name, email) {
    const s = String(name || '').trim();
    if (s) {
        const parts = s.split(/\s+/).filter(Boolean);
        if (parts.length >= 2) {
            return (parts[0].charAt(0) + parts[1].charAt(0)).toUpperCase();
        }
        return s.slice(0, 2).toUpperCase();
    }
    const em = String(email || '').trim();
    return em ? em.slice(0, 2).toUpperCase() : '?';
}

async function acquireOrgGraphToken() {
    try {
        if (typeof window.ms365AuthAcquireTokenSilent === 'function') {
            const t = await window.ms365AuthAcquireTokenSilent();
            if (t) return t;
        }
    } catch {
        /* silent fail → popup */
    }
    if (typeof window.ms365AuthAcquireToken === 'function') {
        return window.ms365AuthAcquireToken();
    }
    if (typeof window.ms365AuthAcquireTokenPopup === 'function') {
        return window.ms365AuthAcquireTokenPopup();
    }
    throw new Error('Anmeldung für Microsoft Graph nicht verfügbar.');
}

/**
 * Mitglieder + Besitzer aus MS365 für ARGE/Klasse (Fokus-Zoom).
 * @param {object} node
 */
async function fetchFocusPeopleForNode(node) {
    const settings = loadSettings();
    const local = listPeopleForSchulOrgNode(settings, node);
    if (!nodeUsesGraphFocusPeople(node)) return local;

    const localLead = (local || []).find(function (p) {
        return String((p && p.kind) || '').toLowerCase() === 'head';
    });
    const gid = resolveGroupIdForNode(node, settings);
    if (!gid) {
        return ensureLeadSeat(local, node, localLead || null);
    }

    const gug = window.ms365GraphUnifiedGroups;
    if (!gug || typeof gug.fetchGroupMembers !== 'function') {
        return ensureLeadSeat(local, node, localLead || null);
    }

    const token = await acquireOrgGraphToken();
    const [members, owners] = await Promise.all([
        gug.fetchGroupMembers(token, gid).catch(function () {
            return [];
        }),
        typeof gug.fetchGroupOwners === 'function'
            ? gug.fetchGroupOwners(token, gid).catch(function () {
                  return [];
              })
            : Promise.resolve([])
    ]);

    return mergeGraphFocusPeople({
        localPeople: local,
        members: members,
        owners: owners,
        node: node,
        teachers: Array.isArray(settings.teachers) ? settings.teachers : [],
        localLead: localLead
            ? { name: localLead.name, email: localLead.email }
            : null
    });
}

/**
 * Verknüpfte M365-Gruppen-ID für Kachel/Bereich (wenn in Stammdaten/Catalog hinterlegt).
 * @param {object} node
 * @param {object} settings
 * @returns {string}
 */
function resolveGroupIdForNode(node, settings) {
    if (!node) return '';
    const container = getAppContainer();
    const setup = container && container.setup ? container.setup : {};
    const matched = setup.matched && typeof setup.matched === 'object' ? setup.matched : {};
    const links = Array.isArray(setup.catalogLinks) ? setup.catalogLinks : [];

    function resolveArgeCatalogGroupId(argeNode) {
        if (!argeNode || !argeNode.refKey) return '';
        const key = String(argeNode.refKey || '')
            .trim()
            .toUpperCase();
        const arge = findArge(settings, argeNode.refKey);
        const nameKey = String((arge && arge.name) || '')
            .trim()
            .toLowerCase();
        const hit = links.find(function (l) {
            if (!l || l.kind !== 'arge') return false;
            const code = String(l.code || '')
                .trim()
                .toUpperCase();
            if (key && code === key) return true;
            if (nameKey && String(l.displayName || '').trim().toLowerCase() === nameKey) return true;
            return false;
        });
        return hit && hit.graphGroupId ? String(hit.graphGroupId).trim() : '';
    }

    if (isAreaNode(node) || node.type === NODE.SECTION) {
        const m365 = node.m365 && typeof node.m365 === 'object' ? node.m365 : null;
        if (m365) {
            const mode = String(m365.groupMode || 'none')
                .trim()
                .toLowerCase();
            // Explizit „Keine Gruppe“ → nie Fallback (Klasse/ARGE/Sammel)
            if (mode === 'none') return '';
            if (mode === 'match' || mode === 'create') {
                const ownId = String(m365.groupId || '').trim();
                // Nur echte Verknüpfung (vorhandene Gruppen-ID) – „anlegen“ ohne ID zählt nicht
                // Kein Fallthrough auf classTeams/Catalog, sonst bleibt eine gelöste Gruppe sichtbar
                return ownId;
            }
        }
        if (isArgeAreaNode(node)) {
            return resolveArgeCatalogGroupId(node);
        }
        // Schüler-Klassen ohne explizites m365: Klassengruppe aus Stammdaten/Teams
        // (Klassenchats/Eltern-Kacheln nie über classTeams auflösen)
        if (isClassAreaNode(node) && node.refKey) {
            const cls = findClass(settings, node.refKey);
            if (!cls) return '';
            const classTeams =
                container && container.core && Array.isArray(container.core.classTeams)
                    ? container.core.classTeams
                    : [];
            const matchMap =
                setup.classGroupMatchByKey && typeof setup.classGroupMatchByKey === 'object'
                    ? setup.classGroupMatchByKey
                    : {};
            return resolveClassGraphGroupId(cls, classTeams, matchMap) || '';
        }
        // Ohne m365 am Knoten: nur Lehrer-/Schüler-Sammelgruppen aus Tenant-Match
        // (Schulleitung/Verwaltung/Audience: Icon nur nach explizitem Match im Bereich-Dialog)
        const sammelKey = sammelgruppeKeyForArea(node);
        if (sammelKey === 'lehrer' || node.id === SCHUL_ORG_SEC_LEHRER) {
            return String(matched.lehrerGroupId || '').trim();
        }
        if (sammelKey === 'schueler' || node.id === SCHUL_ORG_SEC_SCHUELER) {
            return String(matched.schuelerGroupId || '').trim();
        }
        return '';
    }

    if (isClassNode(node) && node.refKey) {
        const cls = findClass(settings, node.refKey);
        if (!cls) return '';
        const classTeams =
            container && container.core && Array.isArray(container.core.classTeams)
                ? container.core.classTeams
                : [];
        const matchMap =
            setup.classGroupMatchByKey && typeof setup.classGroupMatchByKey === 'object'
                ? setup.classGroupMatchByKey
                : {};
        return resolveClassGraphGroupId(cls, classTeams, matchMap) || '';
    }

    if (isArgeNode(node) && node.refKey) {
        return resolveArgeCatalogGroupId(node);
    }

    if (node.type === NODE.KURSTEAM || node.type === 'kursteam') {
        const meta = node.meta || {};
        if (meta.graphGroupId) return String(meta.graphGroupId).trim();
        if (node.m365 && node.m365.groupId) return String(node.m365.groupId).trim();
        const mail = String(meta.gruppenmail || (node.m365 && node.m365.groupMail) || '')
            .trim()
            .toLowerCase();
        if (mail) {
            const hit = links.find(function (l) {
                if (!l) return false;
                const gm = String(l.mailNickname || l.groupMail || l.mail || '')
                    .trim()
                    .toLowerCase();
                const nick = String(l.mailNickname || '')
                    .trim()
                    .toLowerCase();
                if (gm && (gm === mail || mail.indexOf(gm) !== -1)) return true;
                if (nick && (nick === mail || mail.indexOf(nick) !== -1)) return true;
                return false;
            });
            if (hit && hit.graphGroupId) return String(hit.graphGroupId).trim();
        }
        return '';
    }

    if (node.type === NODE.ROLE || node.type === 'role') {
        if (node.m365 && node.m365.groupId) return String(node.m365.groupId).trim();
        if (node.refKind === 'adminRole' && node.refKey) {
            const role = findAdminRole(settings, node.refKey);
            if (role && role.m365ResourceId) return String(role.m365ResourceId).trim();
        }
        return '';
    }

    return '';
}

function hydrateOrgPhotos(root) {
    try {
        if (window.ms365GroupPhotoThumb && typeof window.ms365GroupPhotoThumb.hydrate === 'function') {
            window.ms365GroupPhotoThumb.hydrate(root);
        }
    } catch {
        /* ignore */
    }
    try {
        if (window.ms365UserPhotoThumb && typeof window.ms365UserPhotoThumb.hydrate === 'function') {
            window.ms365UserPhotoThumb.hydrate(root);
        }
    } catch {
        /* ignore */
    }
}

function groupThumbHtml(groupId, displayName) {
    const gid = String(groupId || '').trim();
    const name = String(displayName || '').trim();
    if (!gid) return '';
    const ini = personInitials(name, '');
    return (
        '<span class="gd-group-photo-thumb gd-group-photo-thumb--list schul-org__group-thumb" data-gd-group-photo="' +
        escapeHtml(gid) +
        '" data-gd-group-photo-name="' +
        escapeHtml(name) +
        '" aria-hidden="true">' +
        '<img alt="" hidden />' +
        '<span class="gd-group-photo-thumb__initials">' +
        escapeHtml(ini) +
        '</span></span>'
    );
}

function userThumbHtml(email, displayName) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    const name = String(displayName || '').trim();
    const ini = personInitials(name, em);
    if (!em || em.indexOf('@') === -1) {
        return (
            '<span class="ms365-user-photo-thumb ms365-user-photo-thumb--chip is-empty" aria-hidden="true">' +
            '<span class="ms365-user-photo-thumb__initials">' +
            escapeHtml(ini) +
            '</span></span>'
        );
    }
    return (
        '<span class="ms365-user-photo-thumb ms365-user-photo-thumb--chip" data-ms365-user-photo="' +
        escapeHtml(em) +
        '" data-ms365-user-photo-name="' +
        escapeHtml(name) +
        '" aria-hidden="true">' +
        '<img alt="" hidden />' +
        '<span class="ms365-user-photo-thumb__initials">' +
        escapeHtml(ini) +
        '</span></span>'
    );
}

function saveLayout(layout) {
    const settings = loadSettings();
    settings.schulOrganigramm = normalizeSchulOrganigramm(layout);
    // Nur Layout/Orte speichern – Stammdaten→Org-Sync würde die Landkarte unnötig anfassen
    try {
        window.__ms365SchulOrgSkipSyncOnce = true;
    } catch {
        /* ignore */
    }
    if (typeof window.ms365TenantSettingsSave === 'function') {
        window.ms365TenantSettingsSave(settings);
    }
    return settings.schulOrganigramm;
}

function enrichLayoutMeta(layout, settings) {
    const seeded = seedSchulOrganigrammFromStammdaten(settings);
    const metaByRef = new Map();
    seeded.nodes.forEach(function (n) {
        if (n.refKind && n.refKind !== 'none' && n.refKey && n.meta) {
            metaByRef.set(n.refKind + ':' + String(n.refKey).toUpperCase(), n.meta);
        }
    });
    const base = normalizeSchulOrganigramm(layout);
    return Object.assign({}, base, {
        nodes: (base.nodes || []).map(function (n) {
            const key =
                n.refKind && n.refKey ? n.refKind + ':' + String(n.refKey).toUpperCase() : '';
            const meta = key ? metaByRef.get(key) : null;
            return meta ? Object.assign({}, n, { meta: meta }) : n;
        })
    });
}

function loadViewMode() {
    try {
        const v = localStorage.getItem(VIEW_MODE_KEY);
        if (v === 'list' || v === 'map') return v;
    } catch {
        /* ignore */
    }
    return 'map';
}

function saveViewMode(mode) {
    try {
        localStorage.setItem(VIEW_MODE_KEY, mode === 'list' ? 'list' : 'map');
    } catch {
        /* ignore */
    }
}

function occupancyClass(node) {
    if (node.type === NODE.ROLE) {
        const n = node.meta && typeof node.meta.peopleCount === 'number' ? node.meta.peopleCount : 0;
        if (n === 0) return 'is-vacant';
        if (n === 1) return 'is-single';
        return 'is-team';
    }
    if (isArgeNode(node) || isClassNode(node)) {
        const m = node.meta || {};
        return m.headName || m.headEmail ? 'is-single' : 'is-vacant';
    }
    if (node.type === NODE.COUNCIL || node.type === NODE.SGA) {
        const m = node.meta || {};
        if (m.board || node.refKey === SCHUL_ORG_BOARD_REF) {
            const n = typeof m.peopleCount === 'number' ? m.peopleCount : 0;
            if (n === 0) return 'is-vacant';
            if (n === 1) return 'is-single';
            return 'is-team';
        }
        return m.email ? 'is-single' : 'is-vacant';
    }
    return '';
}

function scopeLabelDe(scope) {
    const s = String(scope || '').toLowerCase();
    if (s === 'teacher') return 'Lehrer';
    if (s === 'student') return 'Schüler';
    if (s === 'external') return 'Extern';
    return s;
}

function cardIconForType(type, node) {
    if (type === NODE.ROLE) return 'bi-person-vcard';
    if (type === NODE.ARGE || (node && isArgeNode(node))) return 'bi-journal-bookmark';
    if (type === NODE.CLASS || (node && isClassNode(node))) return 'bi-mortarboard';
    if (type === NODE.KURSTEAM) return 'bi-collection-play';
    if (type === NODE.COUNCIL) return 'bi-people';
    if (type === NODE.SGA) return 'bi-building';
    return 'bi-hexagon';
}

function occupancyTone(node) {
    const cls = occupancyClass(node);
    if (cls === 'is-single') return 'ok';
    if (cls === 'is-team') return 'team';
    if (cls === 'is-vacant') return 'vacant';
    return '';
}

function metaLine(node) {
    const m = node.meta || {};
    if (node.type === NODE.ROLE) {
        const n = typeof m.peopleCount === 'number' ? m.peopleCount : 0;
        return n === 0 ? 'unbesetzt' : n === 1 ? '1 Person' : n + ' Personen';
    }
    if (isArgeNode(node)) {
        if (!(m.headName || m.headEmail)) return 'Leitung offen';
        const label = m.headName || m.headEmail;
        return m.teacherLinked && m.teacherCode ? label + ' · ' + m.teacherCode : label;
    }
    if (isClassNode(node)) {
        const bits = [];
        if (m.year) bits.push(m.year);
        if (m.headName || m.headEmail) {
            const kv = m.headName || m.headEmail;
            bits.push(m.teacherLinked && m.teacherCode ? kv + ' · ' + m.teacherCode : kv);
        }
        return bits.join(' · ') || 'Klassenvorstand offen';
    }
    if (node.type === NODE.KURSTEAM) {
        const bits = [];
        if (m.fach) bits.push(m.fach);
        if (m.lehrerName || m.lehrerCode || m.lehrerEmail) {
            bits.push(m.lehrerName || m.lehrerCode || m.lehrerEmail);
        }
        if (Array.isArray(m.sharedClassCodes) && m.sharedClassCodes.length > 1) {
            bits.push(m.sharedClassCodes.join(' · '));
        }
        return bits.join(' · ') || 'Unterrichtsteam';
    }
    if (node.type === NODE.COUNCIL) {
        if (m.board || node.refKey === SCHUL_ORG_BOARD_REF) {
            const n = typeof m.peopleCount === 'number' ? m.peopleCount : 0;
            return n === 0 ? 'keine Personen' : n === 1 ? '1 Person' : n + ' Personen';
        }
        return m.klasse || m.email || 'ohne Klasse';
    }
    if (node.type === NODE.SGA) {
        if (m.board || node.refKey === SCHUL_ORG_BOARD_REF) {
            const bits = [];
            if (m.studentCount) bits.push(m.studentCount + ' Schüler');
            if (m.teacherCount) bits.push(m.teacherCount + ' Lehrer');
            if (m.externalCount) bits.push(m.externalCount + ' Eltern');
            return bits.join(' · ') || 'keine Personen';
        }
        const bits = [];
        if (m.scope) bits.push(scopeLabelDe(m.scope));
        if (m.email) bits.push(m.email);
        return bits.join(' · ') || 'SGA';
    }
    return '';
}

function hasAnyStammdaten(settings) {
    const s = settings || {};
    return (
        (s.adminRoles || []).length +
            (s.admin || []).length +
            (s.arges || []).length +
            (s.classes || []).length +
            (s.teachers || []).length +
            (s.studentCouncil || []).length +
            (s.sga || []).length >
        0
    );
}

/**
 * @param {HTMLElement} host
 * @param {{ embedded?: boolean }} [opts]
 */
export function mountSchulOrganigramm(host, opts) {
    if (!host) return null;
    const embedded = !!(opts && opts.embedded);
    /** @type {string} */
    let filterSectionId = (opts && opts.initialFilterSectionId) || '';
    /** @type {string} */
    let textFilter = '';
    /** @type {string} */
    let migrationNotice = '';
    /** @type {'map'|'list'} */
    let viewMode = loadViewMode();
    /** @type {{ name: string, email: string }|null} */
    let personFocus = null;
    /** @type {string} */
    let personFocusQuery = '';
    /** @type {{ setLayout: Function, getLayout: Function, destroy: Function, setPersonFocus?: Function }|null} */
    let mapApi = null;
    /** @type {string} */
    let dragNodeId = '';
    /** @type {'card'|'section'|'person'|''} */
    let dragKind = '';
    /** @type {object|null} */
    let dragPerson = null;
    const editor = createSchulOrganigrammEditor({
        onChanged: function () {
            render();
        }
    });

    function toast(msg, kind) {
        if (typeof window.ms365AppDialogAlert === 'function') {
            window.ms365AppDialogAlert(msg, { kind: kind || 'info', title: 'Schul-Organigramm' });
            return;
        }
        if (typeof window.ms365Toast === 'function') window.ms365Toast(msg);
    }

    async function confirmReseed() {
        if (typeof window.ms365AppDialogConfirm === 'function') {
            return window.ms365AppDialogConfirm(
                'Layout aus den aktuellen Stammdaten neu aufbauen? Die gespeicherte Anordnung im Organigramm wird ersetzt.',
                { title: 'Neu aufbauen', okText: 'Neu aufbauen', cancelText: 'Abbrechen' }
            );
        }
        return window.confirm('Organigramm aus Stammdaten neu aufbauen?');
    }

    async function confirmMigrateFromVerwaltung() {
        if (typeof window.ms365AppDialogConfirm === 'function') {
            return window.ms365AppDialogConfirm(
                'Layout aus dem alten Verwaltungs-Organigramm übernehmen? Die aktuelle Anordnung im Schul-Organigramm wird ersetzt; ARGEs/Klassen werden ergänzt. Das Legacy-Layout bleibt archiviert.',
                {
                    title: 'Aus Verwaltung übernehmen',
                    okText: 'Übernehmen',
                    cancelText: 'Abbrechen'
                }
            );
        }
        return window.confirm('Aus Verwaltungs-Organigramm übernehmen?');
    }

    function render() {
        // Live-Kartenpositionen vor Remount sichern (Speichern/Notify darf Orte nicht verwerfen)
        let liveMapLayout = null;
        if (mapApi) {
            try {
                if (typeof mapApi.flushLandkartePositions === 'function') {
                    liveMapLayout = mapApi.flushLandkartePositions();
                } else if (typeof mapApi.getLayout === 'function') {
                    liveMapLayout = mapApi.getLayout();
                }
            } catch {
                liveMapLayout = null;
            }
        }
        const settings = loadSettings();
        const resolved = resolveSchulOrganigrammLayout(
            liveMapLayout
                ? mergeSchulOrganigrammLivePositions(settings.schulOrganigramm, liveMapLayout)
                : settings.schulOrganigramm,
            settings
        );
        let layout = enrichLayoutMeta(resolved.layout, settings);
        if (liveMapLayout) {
            layout = enrichLayoutMeta(
                mergeSchulOrganigrammLivePositions(layout, liveMapLayout),
                settings
            );
        }
        if (resolved.dirty || liveMapLayout) {
            layout = enrichLayoutMeta(saveLayout(layout), settings);
            if (resolved.migratedFromVerwaltung) {
                const st = resolved.migrationStats || {};
                migrationNotice =
                    'Verwaltungs-Organigramm übernommen (' +
                    (st.roles || 0) +
                    ' Rollen' +
                    (st.sammel ? ', ' + st.sammel + ' Sammelgruppen' : '') +
                    '). Legacy bleibt unter verwaltungOrganigramm archiviert.';
            } else if (resolved.healed || (resolved.deduped && resolved.removedDuplicates > 0)) {
                migrationNotice =
                    'Organigramm bereinigt: ' +
                    (resolved.removedDuplicates || 0) +
                    ' doppelte/ungültige Einträge entfernt und aus den Stammdaten neu aufgebaut.';
            }
        }
        const stats = schulOrganigrammStats(layout);
        const roots = childrenOf(layout, null);
        const q = textFilter.trim().toLowerCase();
        const visibleRoots = roots.filter(function (r) {
            if (filterSectionId && r.id !== filterSectionId) return false;
            if (!q) return true;
            if (nodeMatchesSchulOrgFilter(r, q)) return true;
            return childrenOf(layout, r.id).some(function (k) {
                return nodeMatchesSchulOrgFilter(k, q, metaLine(k));
            });
        });
        const legacyAvailable = hasVerwaltungOrganigrammLayout(settings);

        host.replaceChildren();
        host.classList.add('schul-org');
        host.classList.toggle('schul-org--map', viewMode === 'map');
        host.classList.toggle('schul-org--list', viewMode === 'list');
        if (embedded) host.classList.add('schul-org--embedded');

        const chrome = document.createElement('div');
        chrome.className =
            'schul-org__chrome schul-org__toolbar--no-print' +
            (viewMode === 'map' ? ' schul-org__chrome--compact' : '');
        const searchHtml =
            '<div class="schul-org__search">' +
            '<i class="bi bi-search" aria-hidden="true"></i>' +
            '<label class="visually-hidden" for="schulOrgSearch">Suche</label>' +
            '<input type="search" id="schulOrgSearch" placeholder="Suche …" autocomplete="off">' +
            '</div>';
        chrome.innerHTML =
            '<div class="schul-org__chrome-top">' +
            '<div class="schul-org__brand">' +
            '<span class="schul-org__brand-icon" aria-hidden="true"><i class="bi bi-diagram-3"></i></span>' +
            '<div class="schul-org__brand-text">' +
            '<h2 class="schul-org__title">Schul-Organigramm</h2>' +
            (viewMode === 'map'
                ? ''
                : '<p class="schul-org__lead">Bereiche &amp; Rollen · Personen ziehen · Doppelklick bearbeiten</p>') +
            '</div></div>' +
            '<div class="schul-org__stat-row" id="schulOrgMeta" aria-label="Kurzstatistik"></div>' +
            '<div class="schul-org__top-end">' +
            '<div class="schul-org__actions" id="schulOrgActions"></div>' +
            (viewMode === 'map' ? searchHtml : '') +
            '</div></div>' +
            (viewMode === 'map'
                ? ''
                : '<div class="schul-org__chrome-bar">' +
                  '<div class="schul-org__filters" role="toolbar" aria-label="Bereich filtern"></div>' +
                  searchHtml +
                  '</div>' +
                  '<div class="schul-org__legend" aria-label="Legende">' +
                  '<span class="schul-org__legend-item"><span class="schul-org__swatch schul-org__swatch--vacant"></span>unbesetzt</span>' +
                  '<span class="schul-org__legend-item"><span class="schul-org__swatch schul-org__swatch--single"></span>1 Person</span>' +
                  '<span class="schul-org__legend-item"><span class="schul-org__swatch schul-org__swatch--team"></span>Team</span>' +
                  '<span class="schul-org__legend-item"><span class="schul-org__m365-mark-legend is-group" aria-hidden="true"><i class="bi bi-microsoft"></i></span>M365 verknüpft</span>' +
                  '<span class="schul-org__legend-item schul-org__legend-item--hint">Person-Chips ziehen zum Umbesetzen</span>' +
                  '</div>');
        host.appendChild(chrome);

        if (migrationNotice) {
            const banner = document.createElement('div');
            banner.className = 'schul-org__banner schul-org__banner--ok';
            banner.innerHTML =
                '<p><strong>Hinweis.</strong> ' +
                escapeHtml(migrationNotice) +
                '</p>' +
                '<p class="schul-org__banner-actions">' +
                '<button type="button" class="btn btn-sm" data-schul-org-dismiss-mig>Schließen</button></p>';
            banner.querySelector('[data-schul-org-dismiss-mig]').addEventListener('click', function () {
                migrationNotice = '';
                render();
            });
            host.appendChild(banner);
        }

        const metaEl = chrome.querySelector('#schulOrgMeta');
        if (metaEl) {
            if (viewMode === 'map') {
                metaEl.innerHTML =
                    '<span class="schul-org__stat"><strong>' +
                    stats.roles +
                    '</strong> Rollen</span>' +
                    '<span class="schul-org__stat"><strong>' +
                    stats.classes +
                    '</strong> Klassen</span>' +
                    '<span class="schul-org__stat schul-org__stat--total"><strong>' +
                    stats.content +
                    '</strong></span>';
            } else {
                metaEl.innerHTML =
                    '<span class="schul-org__stat"><strong>' +
                    stats.roles +
                    '</strong> Rollen</span>' +
                    '<span class="schul-org__stat"><strong>' +
                    stats.arges +
                    '</strong> ARGEs</span>' +
                    '<span class="schul-org__stat"><strong>' +
                    stats.classes +
                    '</strong> Klassen</span>' +
                    '<span class="schul-org__stat"><strong>' +
                    stats.other +
                    '</strong> weitere</span>' +
                    '<span class="schul-org__stat schul-org__stat--total"><strong>' +
                    stats.content +
                    '</strong> Einträge</span>';
            }
        }

        if (!hasAnyStammdaten(settings) && !layout.nodes.filter(function (n) {
            return n.type !== NODE.SECTION;
        }).length) {
            const empty = document.createElement('div');
            empty.className = 'schul-org__banner schul-org__banner--hint';
            empty.innerHTML =
                '<p><strong>Noch keine Stammdaten.</strong> Zuerst Listen pflegen (Rollen, ARGEs, Klassen …), dann hier neu aufbauen.</p>' +
                '<p class="schul-org__banner-actions">' +
                '<a class="btn btn-sm" href="' +
                (embedded ? 'tenant.html#verwaltung' : '../tenant.html#verwaltung') +
                '"><i class="bi bi-database" aria-hidden="true"></i> Stammdaten öffnen</a></p>';
            host.appendChild(empty);
        }

        const filterBar = chrome.querySelector('.schul-org__filters');
        if (filterBar) {
            const allBtn = document.createElement('button');
            allBtn.type = 'button';
            allBtn.className = 'schul-org__filter' + (!filterSectionId ? ' is-active' : '');
            allBtn.textContent = 'Alle';
            allBtn.addEventListener('click', function () {
                filterSectionId = '';
                render();
            });
            filterBar.appendChild(allBtn);
            roots.forEach(function (sec) {
                const btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'schul-org__filter' + (filterSectionId === sec.id ? ' is-active' : '');
                const nKids = childrenOf(layout, sec.id).length;
                const short =
                    sec.id === 'sec_schulleitung'
                        ? 'Schulleitung'
                        : sec.id === 'sec_verwaltung'
                          ? 'Verwaltung'
                          : sec.id === 'sec_arge'
                            ? 'ARGEs'
                            : sec.id === 'sec_klassen'
                              ? 'Klassen'
                              : sec.id === 'sec_vertretung'
                                ? 'Vertretung'
                                : sec.label;
                btn.innerHTML = escapeHtml(short) + ' <em>' + nKids + '</em>';
                btn.title = sec.label;
                btn.addEventListener('click', function () {
                    filterSectionId = sec.id;
                    render();
                });
                filterBar.appendChild(btn);
            });
        }

        const searchEl = /** @type {HTMLInputElement} */ (chrome.querySelector('#schulOrgSearch'));
        searchEl.value = textFilter;
        searchEl.addEventListener('input', function () {
            textFilter = searchEl.value || '';
            render();
            const again = host.querySelector('#schulOrgSearch');
            if (again instanceof HTMLInputElement) {
                again.focus();
                const len = again.value.length;
                again.setSelectionRange(len, len);
            }
        });

        const actions = chrome.querySelector('#schulOrgActions');
        function addAction(opts) {
            const el = document.createElement(opts.href ? 'a' : 'button');
            if (!opts.href) el.type = 'button';
            else el.href = opts.href;
            const iconOnly =
                viewMode === 'map' &&
                opts.iconOnly !== false &&
                !opts.primary &&
                !opts.success &&
                !opts.keepLabel;
            el.className =
                'schul-org__action' +
                (opts.primary ? ' schul-org__action--primary' : '') +
                (opts.success ? ' schul-org__action--success' : '') +
                (iconOnly ? ' schul-org__action--icon' : '');
            el.title = opts.title || opts.label;
            el.innerHTML =
                '<i class="bi ' +
                opts.icon +
                '" aria-hidden="true"></i>' +
                (iconOnly ? '' : '<span>' + escapeHtml(opts.label) + '</span>');
            if (opts.onClick) el.addEventListener('click', opts.onClick);
            actions.appendChild(el);
            return el;
        }

        function saveOrganigrammToStammdaten() {
            const cur = loadSettings();
            let layoutNow = null;
            if (mapApi) {
                try {
                    // Sichtbare Landkarten-Orte aus dem DOM ins Layout schreiben
                    if (typeof mapApi.flushLandkartePositions === 'function') {
                        layoutNow = mapApi.flushLandkartePositions();
                    } else if (typeof mapApi.getLayout === 'function') {
                        layoutNow = mapApi.getLayout();
                    }
                } catch {
                    layoutNow = null;
                }
            }
            if (!layoutNow) layoutNow = cur.schulOrganigramm || layout;
            layoutNow = normalizeSchulOrganigramm(layoutNow);
            // Snapshot der Root-Orte – nach Save/Sync wieder aufzwingen
            const positionAnchor = layoutNow;
            const applied = applySchulOrganigrammToStammdaten(cur, layoutNow);
            applied.settings.schulOrganigramm = mergeSchulOrganigrammLivePositions(
                applied.settings.schulOrganigramm,
                positionAnchor
            );
            // Stammdaten-Sync soll die gerade gespeicherten Orte nicht zerlegen
            try {
                window.__ms365SchulOrgSkipSyncOnce = true;
            } catch {
                /* ignore */
            }
            if (typeof window.ms365TenantSettingsSave === 'function') {
                window.ms365TenantSettingsSave(applied.settings);
            }
            // Nach Save nochmals Orte in den Kanon schreiben (falls Sync trotzdem lief)
            try {
                const api = window.ms365AppDataV2;
                if (api && typeof api.getContainer === 'function' && typeof api.setContainer === 'function') {
                    const c = api.getContainer();
                    if (c && c.core) {
                        c.core.schulOrganigramm = mergeSchulOrganigrammLivePositions(
                            c.core.schulOrganigramm,
                            positionAnchor
                        );
                        api.setContainer(c);
                    }
                }
            } catch {
                /* ignore */
            }
            try {
                const spo = window.ms365StammdatenSpoAutoSync;
                if (spo && typeof spo.schedulePush === 'function') {
                    spo.schedulePush(0);
                } else if (spo && typeof spo.runPush === 'function') {
                    void spo.runPush({ reason: 'schul-organigramm-save' });
                }
            } catch {
                /* ignore */
            }
            const savedLayout = mergeSchulOrganigrammLivePositions(
                applied.layout,
                positionAnchor
            );
            const ch = applied.changed || {};
            const bits = [];
            if (ch.roles) bits.push(ch.roles + ' Rollen');
            if (ch.arges) bits.push(ch.arges + ' ARGEs');
            if (ch.classes) bits.push(ch.classes + ' Klassen');
            if (ch.audiences) bits.push(ch.audiences + ' Bereiche');
            toast(
                bits.length
                    ? 'Organigramm gespeichert und in Stammdaten übernommen (' +
                          bits.join(', ') +
                          ').'
                    : 'Organigramm gespeichert (Stammdaten aktualisiert).',
                'success'
            );
            // Kein Voll-Remount: Positionen bleiben, nur Layout/Stats aktualisieren
            if (mapApi && typeof mapApi.setLayout === 'function') {
                try {
                    mapApi.setLayout(savedLayout);
                } catch {
                    render();
                    return;
                }
                const meta = host.querySelector('#schulOrgMeta');
                if (meta) {
                    const st = schulOrganigrammStats(savedLayout);
                    meta.innerHTML =
                        '<span class="schul-org__stat"><strong>' +
                        st.sections +
                        '</strong> Bereiche</span>' +
                        '<span class="schul-org__stat"><strong>' +
                        st.roles +
                        '</strong> Rollen</span>' +
                        '<span class="schul-org__stat"><strong>' +
                        st.arges +
                        '</strong> ARGEs</span>' +
                        '<span class="schul-org__stat"><strong>' +
                        st.classes +
                        '</strong> Klassen</span>' +
                        '<span class="schul-org__stat schul-org__stat--total"><strong>' +
                        st.total +
                        '</strong> Knoten</span>';
                }
            } else {
                render();
            }
        }
        addAction({
            icon: viewMode === 'map' ? 'bi-list-ul' : 'bi-map',
            label: viewMode === 'map' ? 'Liste' : 'Landkarte',
            keepLabel: true,
            title:
                viewMode === 'map'
                    ? 'Zur Listenansicht wechseln'
                    : 'Zur Schul-Landkarte (Zoom/Pan) wechseln',
            onClick: function () {
                if (mapApi) {
                    try {
                        mapApi.destroy();
                    } catch {
                        /* ignore */
                    }
                    mapApi = null;
                }
                viewMode = viewMode === 'map' ? 'list' : 'map';
                saveViewMode(viewMode);
                render();
            }
        });
        if (viewMode === 'map') {
            addAction({
                icon: 'bi-person-bounding-box',
                label: personFocus ? 'Person · an' : 'Person',
                keepLabel: true,
                title: personFocus
                    ? 'Personen-Ansicht beenden'
                    : 'Person wählen: alle Mitgliedschaften & Funktionen hervorheben',
                primary: !!personFocus,
                onClick: function () {
                    if (personFocus) {
                        personFocus = null;
                        personFocusQuery = '';
                        render();
                        return;
                    }
                    // Ansicht aktivieren – Picker erscheint darunter
                    personFocusQuery = '';
                    personFocus = { name: '', email: '' };
                    render();
                    window.setTimeout(function () {
                        const inp = host.querySelector('#schulOrgPersonFocusSearch');
                        if (inp) /** @type {HTMLInputElement} */ (inp).focus();
                    }, 30);
                }
            });
        }
        addAction({
            icon: 'bi-plus-lg',
            label: 'Bereich',
            title: 'Neue Verwaltungsgruppe / Bereich (z. B. Erweiterte Schulleitung)',
            primary: true,
            onClick: function () {
                void editor.createAudienceSection().then(function (created) {
                    if (created && created.sectionId) {
                        filterSectionId = '';
                        render();
                    }
                });
            }
        });
        addAction({
            icon: 'bi-person-badge',
            label: 'Funktion',
            title:
                'Neue Funktion im zuletzt genutzten Bereich anlegen (Rechtsklick → Funktion hinzufügen setzt den Ort)',
            primary: true,
            keepLabel: true,
            onClick: function () {
                const parentId =
                    (typeof editor.getLastRoleParent === 'function' &&
                        editor.getLastRoleParent()) ||
                    filterSectionId ||
                    'sec_verwaltung';
                void editor.createRoleUnderSection(parentId).then(function () {
                    filterSectionId = '';
                    render();
                });
            }
        });
        addAction({
            icon: 'bi-person-plus',
            label: 'Person',
            title: 'Person einem Bereich auf der Landkarte zuordnen',
            primary: true,
            onClick: function () {
                void editor.openAddCanvasPerson(null);
            }
        });
        addAction({
            icon: 'bi-printer',
            label: 'Drucken',
            title: 'Druckansicht',
            onClick: function () {
                document.documentElement.classList.add('schul-org-printing');
                window.setTimeout(function () {
                    try {
                        window.print();
                    } finally {
                        document.documentElement.classList.remove('schul-org-printing');
                    }
                }, 50);
            }
        });
        addAction({
            icon: 'bi-arrow-clockwise',
            label: 'Aktualisieren',
            onClick: function () {
                render();
                toast('Ansicht aktualisiert.', 'success');
            }
        });
        addAction({
            icon: 'bi-arrow-repeat',
            label: 'Neu aufbauen',
            title: 'Aus Stammdaten neu aufbauen',
            onClick: function () {
                void confirmReseed().then(function (ok) {
                    if (!ok) return;
                    const cur = loadSettings();
                    const next = rebuildSchulOrganigrammFromStammdaten({ version: 1, nodes: [] }, cur);
                    saveLayout(next);
                    filterSectionId = '';
                    textFilter = '';
                    migrationNotice = 'Organigramm frisch aus den Stammdaten aufgebaut.';
                    render();
                    toast('Organigramm aus den Stammdaten neu aufgebaut.', 'success');
                });
            }
        });
        if (legacyAvailable) {
            addAction({
                icon: 'bi-box-arrow-in-down',
                label: 'Aus Verwaltung',
                title: 'Legacy-Verwaltungs-Organigramm übernehmen',
                onClick: function () {
                    void confirmMigrateFromVerwaltung().then(function (ok) {
                        if (!ok) return;
                        const cur = loadSettings();
                        const migrated = migrateVerwaltungOrganigrammToSchul(cur);
                        if (!migrated) {
                            toast('Kein Verwaltungs-Organigramm gefunden.', 'warning');
                            return;
                        }
                        saveLayout(migrated.layout);
                        const st = migrated.stats || {};
                        migrationNotice =
                            'Erneut aus Verwaltung übernommen (' + (st.roles || 0) + ' Rollen).';
                        filterSectionId = '';
                        render();
                        toast('Verwaltungs-Organigramm übernommen.', 'success');
                    });
                }
            });
        }
        addAction({
            icon: 'bi-floppy',
            label: 'Speichern',
            title: 'Aktuelles Schul-Organigramm speichern und in die Stammdaten übertragen',
            success: true,
            keepLabel: true,
            onClick: function () {
                saveOrganigrammToStammdaten();
            }
        });
        addAction({
            icon: 'bi-cloud-arrow-up',
            label: 'Abgleich',
            title:
                'Vorschau & Abgleich mit Microsoft 365: Gruppen/Mitglieder (Graph), Postfächer (PowerShell/.cmd)',
            primary: true,
            keepLabel: true,
            onClick: function () {
                openSchulOrgProvisionDiffPanel({
                    loadSettings: loadSettings,
                    saveSettings: function (s) {
                        if (typeof window.ms365TenantSettingsSave === 'function') {
                            window.ms365TenantSettingsSave(s);
                        }
                    },
                    onChanged: function () {
                        render();
                    },
                    toast: toast
                });
            }
        });
        if (!embedded) {
            addAction({
                icon: 'bi-database',
                label: 'Stammdaten',
                href: '../tenant.html#verwaltung',
                primary: true
            });
        } else {
            addAction({
                icon: 'bi-arrows-fullscreen',
                label: 'Vollbild',
                href: 'tools/schul-organigramm.html',
                primary: true
            });
        }

        if (mapApi) {
            try {
                mapApi.destroy();
            } catch {
                /* ignore */
            }
            mapApi = null;
        }

        if (viewMode === 'map') {
            const focusActive = !!(
                personFocus &&
                (String(personFocus.email || '').trim() || String(personFocus.name || '').trim())
            );
            const focusBundle = focusActive
                ? collectSchulOrgPersonMemberships(settings, layout, personFocus)
                : null;

            if (personFocus) {
                const focusBar = document.createElement('div');
                focusBar.className = 'schul-org__person-focus';
                focusBar.setAttribute('aria-label', 'Personen-Ansicht');
                const suggestions = searchSchulOrgDirectoryPeople(
                    settings,
                    personFocusQuery || '',
                    12
                );
                focusBar.innerHTML =
                    '<div class="schul-org__person-focus-head">' +
                    '<div class="schul-org__person-focus-title">' +
                    '<i class="bi bi-person-bounding-box" aria-hidden="true"></i>' +
                    '<div>' +
                    '<strong>Personen-Ansicht</strong>' +
                    '<span class="muted">Mitgliedschaften &amp; Funktionen auf der Landkarte</span>' +
                    '</div></div>' +
                    '<div class="schul-org__person-focus-search">' +
                    '<label class="visually-hidden" for="schulOrgPersonFocusSearch">Person suchen</label>' +
                    '<input type="search" id="schulOrgPersonFocusSearch" placeholder="Name oder E-Mail …" autocomplete="off" value="' +
                    escapeHtml(personFocusQuery) +
                    '">' +
                    (focusActive
                        ? '<button type="button" class="btn btn-sm" data-person-focus-clear title="Auswahl leeren">Auswahl leeren</button>'
                        : '') +
                    '<button type="button" class="btn btn-sm" data-person-focus-close title="Ansicht schließen">Schließen</button>' +
                    '</div></div>' +
                    (!focusActive
                        ? '<div class="schul-org__person-focus-suggest" role="listbox" aria-label="Vorschläge">' +
                          (suggestions.length
                              ? suggestions
                                    .map(function (p) {
                                        return (
                                            '<button type="button" class="schul-org__person-focus-suggest-item" role="option" data-focus-email="' +
                                            escapeHtml(p.email || '') +
                                            '" data-focus-name="' +
                                            escapeHtml(p.name || '') +
                                            '">' +
                                            userThumbHtml(p.email, p.name) +
                                            '<span class="schul-org__person-focus-suggest-text">' +
                                            '<strong>' +
                                            escapeHtml(p.name || p.email || '?') +
                                            '</strong>' +
                                            (p.email
                                                ? '<span class="muted">' +
                                                  escapeHtml(p.email) +
                                                  '</span>'
                                                : '') +
                                            (p.sources && p.sources.length
                                                ? '<span class="schul-org__person-focus-sources">' +
                                                  escapeHtml(p.sources.join(' · ')) +
                                                  '</span>'
                                                : '') +
                                            '</span></button>'
                                        );
                                    })
                                    .join('')
                              : '<p class="muted schul-org__person-focus-empty">Keine Treffer – anderen Suchbegriff versuchen.</p>') +
                          '</div>'
                        : '<div class="schul-org__person-focus-result">' +
                          '<div class="schul-org__person-focus-who">' +
                          userThumbHtml(personFocus.email, personFocus.name) +
                          '<div><strong>' +
                          escapeHtml(personFocus.name || personFocus.email || 'Person') +
                          '</strong>' +
                          (personFocus.email
                              ? '<span class="muted">' + escapeHtml(personFocus.email) + '</span>'
                              : '') +
                          '<span class="schul-org__person-focus-count">' +
                          ((focusBundle && focusBundle.memberships.length) || 0) +
                          ((focusBundle && focusBundle.memberships.length) === 1
                              ? ' Treffer'
                              : ' Treffer') +
                          '</span></div></div>' +
                          '<ul class="schul-org__person-focus-list">' +
                          ((focusBundle && focusBundle.memberships.length
                              ? focusBundle.memberships
                                    .map(function (m) {
                                        return (
                                            '<li>' +
                                            '<button type="button" class="schul-org__person-focus-hit" data-focus-node="' +
                                            escapeHtml(m.nodeId) +
                                            '">' +
                                            '<span class="schul-org__person-focus-kind">' +
                                            escapeHtml(m.kindLabel) +
                                            '</span>' +
                                            '<span class="schul-org__person-focus-node">' +
                                            escapeHtml(m.nodeLabel) +
                                            '</span>' +
                                            (m.pathText
                                                ? '<span class="muted">' +
                                                  escapeHtml(m.pathText) +
                                                  '</span>'
                                                : '') +
                                            '</button></li>'
                                        );
                                    })
                                    .join('')
                              : '<li class="muted">Keine Mitgliedschaft auf der Landkarte gefunden.</li>') +
                          '</ul></div>'));
                host.appendChild(focusBar);

                function bindFocusSuggestClicks(root) {
                    (root || focusBar)
                        .querySelectorAll('[data-focus-email], [data-focus-name]')
                        .forEach(function (btn) {
                            btn.addEventListener('click', function () {
                                const em = btn.getAttribute('data-focus-email') || '';
                                const nm = btn.getAttribute('data-focus-name') || '';
                                if (!em && !nm) return;
                                personFocus = { email: em, name: nm };
                                personFocusQuery = nm || em;
                                render();
                            });
                        });
                }
                function renderFocusSuggestions() {
                    const box = focusBar.querySelector('.schul-org__person-focus-suggest');
                    if (!box) return;
                    const next = searchSchulOrgDirectoryPeople(
                        settings,
                        personFocusQuery || '',
                        12
                    );
                    box.innerHTML = next.length
                        ? next
                              .map(function (p) {
                                  return (
                                      '<button type="button" class="schul-org__person-focus-suggest-item" role="option" data-focus-email="' +
                                      escapeHtml(p.email || '') +
                                      '" data-focus-name="' +
                                      escapeHtml(p.name || '') +
                                      '">' +
                                      userThumbHtml(p.email, p.name) +
                                      '<span class="schul-org__person-focus-suggest-text">' +
                                      '<strong>' +
                                      escapeHtml(p.name || p.email || '?') +
                                      '</strong>' +
                                      (p.email
                                          ? '<span class="muted">' +
                                            escapeHtml(p.email) +
                                            '</span>'
                                          : '') +
                                      (p.sources && p.sources.length
                                          ? '<span class="schul-org__person-focus-sources">' +
                                            escapeHtml(p.sources.join(' · ')) +
                                            '</span>'
                                          : '') +
                                      '</span></button>'
                                  );
                              })
                              .join('')
                        : '<p class="muted schul-org__person-focus-empty">Keine Treffer – anderen Suchbegriff versuchen.</p>';
                    bindFocusSuggestClicks(box);
                    if (typeof hydrateOrgPhotos === 'function') {
                        try {
                            hydrateOrgPhotos(box);
                        } catch {
                            /* ignore */
                        }
                    }
                }
                const searchEl = /** @type {HTMLInputElement|null} */ (
                    focusBar.querySelector('#schulOrgPersonFocusSearch')
                );
                if (searchEl) {
                    searchEl.addEventListener('input', function () {
                        personFocusQuery = searchEl.value || '';
                        // Karte nicht remounten – nur Vorschläge aktualisieren
                        if (focusActive) return;
                        renderFocusSuggestions();
                    });
                    searchEl.addEventListener('keydown', function (ev) {
                        if (ev.key === 'Escape') {
                            personFocus = null;
                            personFocusQuery = '';
                            render();
                        }
                    });
                }
                bindFocusSuggestClicks(focusBar);
                const clearBtn = focusBar.querySelector('[data-person-focus-clear]');
                if (clearBtn) {
                    clearBtn.addEventListener('click', function () {
                        personFocus = { name: '', email: '' };
                        personFocusQuery = '';
                        render();
                        window.setTimeout(function () {
                            const inp = host.querySelector('#schulOrgPersonFocusSearch');
                            if (inp) /** @type {HTMLInputElement} */ (inp).focus();
                        }, 30);
                    });
                }
                const closeBtn = focusBar.querySelector('[data-person-focus-close]');
                if (closeBtn) {
                    closeBtn.addEventListener('click', function () {
                        personFocus = null;
                        personFocusQuery = '';
                        render();
                    });
                }
                focusBar.querySelectorAll('[data-focus-node]').forEach(function (btn) {
                    btn.addEventListener('click', function () {
                        const id = btn.getAttribute('data-focus-node') || '';
                        if (!id || !mapApi || typeof mapApi.enterFocus !== 'function') return;
                        mapApi.enterFocus(id);
                    });
                });
            }

            const mapHost = document.createElement('div');
            mapHost.className =
                'schul-org__map-host' + (personFocus ? ' schul-org__map-host--person-focus' : '');
            mapHost.setAttribute('aria-label', 'Schul-Landkarte');
            host.appendChild(mapHost);
            mapApi = mountSchulOrganigrammLandkarte({
                host: mapHost,
                layout: layout,
                personFocus: focusBundle
                    ? {
                          active: true,
                          email: focusBundle.person.email,
                          name: focusBundle.person.name,
                          hitIds: focusBundle.hitIds,
                          pathIds: focusBundle.pathIds
                      }
                    : null,
                onLayoutChange: function (next) {
                    saveLayout(next);
                },
                onMoveNode: function (nodeId, newParentId, liveLayout) {
                    const cur = loadSettings();
                    // Aktuelles Karten-Layout bevorzugen (nicht veraltete Settings)
                    const before = normalizeSchulOrganigramm(
                        liveLayout && Array.isArray(liveLayout.nodes)
                            ? liveLayout
                            : (mapApi && typeof mapApi.getLayout === 'function'
                                  ? mapApi.getLayout()
                                  : null) || cur.schulOrganigramm
                    );
                    const moved = moveSchulOrganigrammNode(before, nodeId, newParentId);
                    const nodeBefore = before.nodes.find(function (n) {
                        return n.id === nodeId;
                    });
                    const nodeAfter = moved.nodes.find(function (n) {
                        return n.id === nodeId;
                    });
                    if (
                        !nodeBefore ||
                        !nodeAfter ||
                        String(nodeBefore.parentId || '') === String(nodeAfter.parentId || '')
                    ) {
                        toast('Hier kann diese Kachel nicht abgelegt werden.', 'warning');
                        return null;
                    }
                    const applied = applySchulOrganigrammToStammdaten(cur, moved);
                    if (typeof window.ms365TenantSettingsSave === 'function') {
                        window.ms365TenantSettingsSave(applied.settings);
                    } else {
                        saveLayout(moved);
                    }
                    return applied.layout;
                },
                onMovePerson: function (person, targetNode) {
                    const result = editor.movePersonToNode(person, targetNode);
                    if (!result || !result.ok) {
                        toast((result && result.message) || 'Verschieben nicht möglich.', 'warning');
                        return result;
                    }
                    toast(result.message || 'Person zugeordnet.', 'success');
                    render();
                    return result;
                },
                peopleCountForArea: function (areaNode) {
                    // Nur direkte Bereichspersonen – Rollenbesetzungen zählen auf der Rolle
                    const settings = loadSettings();
                    const all = listPeopleForSchulOrgNode(settings, areaNode);
                    const org = normalizeSchulOrganigramm(
                        settings.schulOrganigramm || (mapApi && mapApi.getLayout && mapApi.getLayout())
                    );
                    /** @type {Set<string>} */
                    const onRoles = new Set();
                    childrenOf(org, areaNode && areaNode.id).forEach(function (ch) {
                        const ct = String((ch && ch.type) || '').toLowerCase();
                        if (ct !== NODE.ROLE && ct !== 'role') return;
                        listPeopleForSchulOrgNode(settings, ch).forEach(function (p) {
                            const em = String((p && p.email) || '')
                                .trim()
                                .toLowerCase();
                            if (em) onRoles.add(em);
                        });
                    });
                    return all.filter(function (p) {
                        const em = String((p && p.email) || '')
                            .trim()
                            .toLowerCase();
                        if (em && onRoles.has(em)) return false;
                        return !!(p && (p.name || p.email) && !p.vacant);
                    }).length;
                },
                peopleForNode: function (node) {
                    // Übersicht: begrenzte Liste; Fokus/Zoom lädt MS365-Mitglieder nach
                    if (nodeUsesGraphFocusPeople(node)) {
                        const local = listPeopleForSchulOrgNode(loadSettings(), node);
                        const localLead = (local || []).find(function (p) {
                            return String((p && p.kind) || '').toLowerCase() === 'head';
                        });
                        return ensureLeadSeat(local, node, localLead || null);
                    }
                    return previewPeopleForSchulOrgNode(loadSettings(), node, 500).all;
                },
                fetchFocusPeople: fetchFocusPeopleForNode,
                userThumbHtml: userThumbHtml,
                groupThumbHtml: groupThumbHtml,
                resolveGroupId: function (node) {
                    return resolveGroupIdForNode(node, loadSettings());
                },
                hydratePhotos: hydrateOrgPhotos,
                onEditNode: function (node) {
                    if (!node) return;
                    if (isSgaBoardArea(node) || isSgaScopeArea(node) || node.type === NODE.SGA) {
                        editor.openSgaEditor(SCHUL_ORG_BOARD_REF);
                        return;
                    }
                    if (isArgeAreaNode(node) && node.refKey) editor.openArgeEditor(node.refKey);
                    else if (isClassAreaNode(node) && node.refKey) editor.openClassEditor(node.refKey);
                    else if (isAreaNode(node)) editor.openAreaEditor(node);
                    else if (node.type === NODE.ROLE && node.refKey) editor.openRoleEditor(node.refKey);
                    else if (isArgeNode(node) && node.refKey) editor.openArgeEditor(node.refKey);
                    else if (isClassNode(node) && node.refKey) editor.openClassEditor(node.refKey);
                    else if (node.type === NODE.KURSTEAM) {
                        const parentClass = node.parentId
                            ? layout.nodes.find(function (n) {
                                  return n.id === node.parentId;
                              })
                            : null;
                        if (parentClass && parentClass.type === NODE.CLASS && parentClass.refKey) {
                            editor.openClassEditor(parentClass.refKey);
                        } else if (node.meta && node.meta.placementClass) {
                            editor.openClassEditor(node.meta.placementClass);
                        }
                    } else if (node.type === NODE.COUNCIL) editor.openCouncilEditor(node.refKey || SCHUL_ORG_BOARD_REF);
                    else if (node.type === NODE.PERSON) editor.openAreaEditor(node.parentId || node);
                },
                onCreateArea: function (place) {
                    if (place && place.parentId && typeof editor.rememberRoleParent === 'function') {
                        editor.rememberRoleParent(place.parentId);
                    }
                    return editor.createAreaAt({
                        parentId: place && place.parentId != null ? place.parentId : null,
                        x: place && place.x,
                        y: place && place.y
                    });
                },
                onAddPerson: function (areaNode) {
                    if (areaNode && areaNode.id && typeof editor.rememberRoleParent === 'function') {
                        editor.rememberRoleParent(areaNode.id);
                    }
                    if (!areaNode) {
                        return editor.openAddCanvasPerson(null);
                    }
                    return editor.openAddCanvasPerson(areaNode);
                },
                onCreateRole: function (areaNode) {
                    if (!areaNode || !areaNode.id) return;
                    if (typeof editor.rememberRoleParent === 'function') {
                        editor.rememberRoleParent(areaNode.id);
                    }
                    return editor.createRoleUnderSection(areaNode.id);
                },
                onEditPerson: function (person, contextNode) {
                    editor.openEditCanvasPerson(person, contextNode);
                }
            });
        } else {
            const canvas = document.createElement('div');
            canvas.className = 'schul-org__canvas';
            canvas.setAttribute('aria-label', 'Schul-Organigramm');

            if (!visibleRoots.length) {
                const empty = document.createElement('p');
                empty.className = 'muted';
                empty.textContent = q
                    ? 'Keine Treffer für die aktuelle Suche/Filterung.'
                    : 'Noch keine Struktur – zuerst Listen in den Stammdaten pflegen oder „Aus Stammdaten neu aufbauen“.';
                canvas.appendChild(empty);
            } else {
                visibleRoots.forEach(function (sec) {
                    canvas.appendChild(renderSection(sec, layout));
                });
            }
            host.appendChild(canvas);
            hydrateOrgPhotos(host);
        }
    }

    function persistLayout(next) {
        saveLayout(normalizeSchulOrganigramm(next));
        render();
    }

    function renderSection(sec, layout) {
        const el = document.createElement('section');
        el.className = 'schul-org__section schul-org__section--' + (sec.tone || 'custom');
        el.dataset.nodeId = sec.id;
        el.dataset.dropSection = '1';
        const headRow = document.createElement('div');
        headRow.className = 'schul-org__section-head';
        headRow.draggable = !textFilter.trim();
        headRow.title = 'Bereich ziehen zum Umsortieren';
        headRow.addEventListener('dragstart', function (ev) {
            dragNodeId = sec.id;
            dragKind = 'section';
            el.classList.add('is-dragging');
            try {
                ev.dataTransfer.setData('text/plain', 'section:' + sec.id);
                ev.dataTransfer.effectAllowed = 'move';
            } catch {
                /* ignore */
            }
        });
        headRow.addEventListener('dragend', function () {
            dragNodeId = '';
            dragKind = '';
            el.classList.remove('is-dragging');
            host.querySelectorAll('.is-drop-target').forEach(function (n) {
                n.classList.remove('is-drop-target');
            });
        });
        el.addEventListener('dragover', function (ev) {
            if (dragKind !== 'section' || dragNodeId === sec.id) return;
            ev.preventDefault();
            el.classList.add('is-drop-target');
        });
        el.addEventListener('dragleave', function () {
            el.classList.remove('is-drop-target');
        });
        el.addEventListener('drop', function (ev) {
            el.classList.remove('is-drop-target');
            if (dragKind !== 'section' || !dragNodeId || dragNodeId === sec.id) return;
            ev.preventDefault();
            const settings = loadSettings();
            const next = reorderSchulOrganigrammSections(
                settings.schulOrganigramm,
                dragNodeId,
                sec.id
            );
            dragNodeId = '';
            dragKind = '';
            persistLayout(next);
        });

        const h = document.createElement('h3');
        h.className = 'schul-org__section-title';
        const allKids = childrenOf(layout, sec.id).filter(function (k) {
            return k.type !== NODE.PERSON;
        });
        const kids = textFilter.trim()
            ? allKids.filter(function (k) {
                  if (nodeMatchesSchulOrgFilter(k, textFilter, metaLine(k))) return true;
                  return childrenOf(layout, k.id).some(function (ch) {
                      return nodeMatchesSchulOrgFilter(ch, textFilter, metaLine(ch));
                  });
              })
            : allKids;
        const leafKids = kids.filter(function (k) {
            return !isAreaNode(k);
        });
        const nestedAreas = kids.filter(isAreaNode);
        const sectionGroupId = resolveGroupIdForNode(sec, loadSettings());
        const sectionM365 = nodeM365LinkFlags(sec, sectionGroupId);
        const sectionIcon = sectionM365.hasGroup
            ? groupThumbHtml(sectionM365.groupId || sectionGroupId, sec.label)
            : '<i class="bi bi-collection" aria-hidden="true"></i>';
        const m365Bits = [];
        if (sectionM365.hasGroup) {
            m365Bits.push(
                sec.m365 && sec.m365.groupId
                    ? schulOrgGroupKindLabel(sec.m365.groupKind)
                    : 'Gruppe'
            );
        }
        if (sectionM365.hasMailbox) m365Bits.push('Postfach');
        const sammelKey = sammelgruppeKeyForArea(sec);
        const sammelPeopleTotal = sammelKey
            ? previewPeopleForSchulOrgNode(loadSettings(), sec, 0).total
            : 0;
        const nestedClassAreas =
            sammelKey === 'schueler'
                ? allKids.filter(function (k) {
                      return isAreaNode(k) || k.type === NODE.CLASS;
                  }).length
                : 0;
        const countLabel =
            sammelKey === 'schueler'
                ? String(nestedClassAreas || (sec.meta && sec.meta.classCount) || 0)
                : sammelKey
                  ? String(sammelPeopleTotal)
                  : String(leafKids.length) +
                    (leafKids.length !==
                    allKids.filter(function (k) {
                        return !isAreaNode(k);
                    }).length
                        ? '/' +
                          allKids.filter(function (k) {
                              return !isAreaNode(k);
                          }).length
                        : '');
        h.innerHTML =
            sectionIcon +
            ' ' +
            escapeHtml(sec.label) +
            ' <span class="muted schul-org__section-count">(' +
            escapeHtml(countLabel) +
            ')</span>' +
            (m365Bits.length
                ? ' <span class="schul-org__m365-badge" title="M365-Anbindung"><i class="bi bi-microsoft" aria-hidden="true"></i> ' +
                  escapeHtml(m365Bits.join(' · ')) +
                  '</span>'
                : '');
        if (sectionM365.linked) {
            headRow.classList.add('is-m365-linked');
        }
        headRow.appendChild(h);
        const areaActions = document.createElement('div');
        areaActions.className = 'schul-org__section-add-group';
        const btnArea = document.createElement('button');
        btnArea.type = 'button';
        btnArea.className = 'btn btn-sm schul-org__section-add';
        btnArea.innerHTML = '<i class="bi bi-pencil" aria-hidden="true"></i> Bearbeiten';
        btnArea.title = 'Bereich: Name, Personen, M365-Gruppe / Postfach';
        btnArea.addEventListener('click', function () {
            editor.openAreaEditor(sec);
        });
        areaActions.appendChild(btnArea);
        if (isRoleCapableSection(sec.id, sec)) {
            const btnAdd = document.createElement('button');
            btnAdd.type = 'button';
            btnAdd.className = 'btn btn-sm btn-success schul-org__section-add';
            btnAdd.innerHTML = '<i class="bi bi-person-vcard" aria-hidden="true"></i> Rolle';
            btnAdd.title = 'Neue Rolle in diesem Bereich anlegen';
            btnAdd.addEventListener('click', function () {
                if (typeof editor.rememberRoleParent === 'function') {
                    editor.rememberRoleParent(sec.id);
                }
                void editor.createRoleUnderSection(sec.id);
            });
            areaActions.appendChild(btnAdd);
            const audId = audienceIdFromSectionId(sec.id) || (sec.refKind === 'audience' ? sec.refKey : '');
            if (audId && String(audId).indexOf('va-') === 0) {
                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'btn btn-sm schul-org__section-add schul-org__section-add--danger';
                btnDel.innerHTML = '<i class="bi bi-trash" aria-hidden="true"></i>';
                btnDel.title = 'Bereich / Verwaltungsgruppe entfernen';
                btnDel.addEventListener('click', function () {
                    void editor.removeAudienceSection(audId).then(function (ok) {
                        if (ok) {
                            if (filterSectionId === sec.id) filterSectionId = '';
                            render();
                        }
                    });
                });
                areaActions.appendChild(btnDel);
            } else if (canDeleteSchulOrganigrammArea(sec)) {
                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'btn btn-sm schul-org__section-add schul-org__section-add--danger';
                btnDel.innerHTML = '<i class="bi bi-trash" aria-hidden="true"></i>';
                btnDel.title = 'Bereich von der Landkarte entfernen';
                btnDel.addEventListener('click', function () {
                    void editor.removeArea(sec).then(function (ok) {
                        if (ok) {
                            if (filterSectionId === sec.id) filterSectionId = '';
                            render();
                        }
                    });
                });
                areaActions.appendChild(btnDel);
            }
        }
        headRow.appendChild(areaActions);
        if (sec.id === 'sec_arge') {
            const btnAdd = document.createElement('button');
            btnAdd.type = 'button';
            btnAdd.className = 'btn btn-sm btn-success schul-org__section-add';
            btnAdd.innerHTML = '<i class="bi bi-journal-bookmark" aria-hidden="true"></i> ARGE';
            btnAdd.title = 'Neue ARGE / Fachschaft anlegen';
            btnAdd.addEventListener('click', function () {
                void editor.createArgeUnderSection(sec.id);
            });
            headRow.appendChild(btnAdd);
        }
        if (sec.id === 'sec_klassen' || isJahrgangArea(sec)) {
            const btnAdd = document.createElement('button');
            btnAdd.type = 'button';
            btnAdd.className = 'btn btn-sm btn-success schul-org__section-add';
            btnAdd.innerHTML = '<i class="bi bi-mortarboard" aria-hidden="true"></i> Klasse';
            btnAdd.title = 'Neue Klasse anlegen';
            btnAdd.addEventListener('click', function () {
                void editor.createClassUnderSection(sec.id);
            });
            headRow.appendChild(btnAdd);
        }
        if (sec.id === 'sec_vertretung') {
            const wrap = document.createElement('div');
            wrap.className = 'schul-org__section-add-group';
            const btnCouncil = document.createElement('button');
            btnCouncil.type = 'button';
            btnCouncil.className = 'btn btn-sm btn-success schul-org__section-add';
            btnCouncil.innerHTML = '<i class="bi bi-people" aria-hidden="true"></i> Vertretung';
            btnCouncil.title = 'Schülervertretung bearbeiten';
            btnCouncil.addEventListener('click', function () {
                editor.openCouncilEditor(SCHUL_ORG_BOARD_REF);
            });
            const btnSga = document.createElement('button');
            btnSga.type = 'button';
            btnSga.className = 'btn btn-sm btn-success schul-org__section-add';
            btnSga.innerHTML = '<i class="bi bi-building" aria-hidden="true"></i> SGA';
            btnSga.title = 'SGA bearbeiten';
            btnSga.addEventListener('click', function () {
                editor.openSgaEditor(SCHUL_ORG_BOARD_REF);
            });
            wrap.appendChild(btnCouncil);
            wrap.appendChild(btnSga);
            headRow.appendChild(wrap);
        }
        el.appendChild(headRow);

        const areaPeople = previewPeopleForSchulOrgNode(
            loadSettings(),
            sec,
            sammelKey === 'lehrer' ? 8 : SCHUL_ORG_PEOPLE_PREVIEW_LIMIT
        );
        const peopleStrip = document.createElement('div');
        peopleStrip.className = 'schul-org__area-people';
        peopleStrip.dataset.dropAreaPeople = sec.id;
        {
            const emptyHint = sammelKey
                ? sammelKey === 'lehrer'
                    ? 'Noch keine Lehrkräfte in den Stammdaten'
                    : 'Klassen liegen in den Jahrgangs-Bereichen'
                : 'Keine direkten Personen – per Drag oder „Bereich“ hinzufügen';
            peopleStrip.innerHTML =
                '<div class="schul-org__area-people-chips">' +
                (areaPeople.shown.length
                    ? areaPeople.shown
                          .map(function (p) {
                              const movable = p.movable !== false && !sammelKey;
                              return (
                                  '<span class="schul-org__person' +
                                  (movable ? ' schul-org__person--movable' : '') +
                                  (p.scope === 'teacher' || p.kind === 'teacher'
                                      ? ' schul-org__person--teacher'
                                      : '') +
                                  (p.scope === 'student' || p.kind === 'student'
                                      ? ' schul-org__person--student'
                                      : '') +
                                  '"' +
                                  (movable
                                      ? ' draggable="true" data-person-key="' +
                                        escapeHtml(p.key) +
                                        '"'
                                      : '') +
                                  ' title="' +
                                  escapeHtml(p.email || p.name) +
                                  '">' +
                                  userThumbHtml(p.email, p.name) +
                                  '<span class="schul-org__person-name">' +
                                  escapeHtml(p.name || p.email) +
                                  '</span></span>'
                              );
                          })
                          .join('')
                    : '<span class="muted schul-org__area-people-empty">' +
                      escapeHtml(emptyHint) +
                      '</span>') +
                (areaPeople.more
                    ? '<button type="button" class="schul-org__person schul-org__person--more" data-area-people-more="1">+' +
                      areaPeople.more +
                      '</button>'
                    : '') +
                '</div>';
            peopleStrip.querySelectorAll('.schul-org__person--movable').forEach(function (chip) {
                chip.addEventListener('dragstart', function (ev) {
                    const key = chip.getAttribute('data-person-key') || '';
                    const person = areaPeople.all.find(function (p) {
                        return p.key === key;
                    });
                    if (!person) {
                        ev.preventDefault();
                        return;
                    }
                    dragKind = 'person';
                    dragPerson = {
                        kind: person.kind,
                        key: person.key,
                        name: person.name,
                        email: person.email,
                        label: person.label,
                        sourceNodeId: sec.id,
                        sourceRefKind: 'area',
                        sourceRefKey: sec.id
                    };
                    chip.classList.add('is-dragging');
                    try {
                        ev.dataTransfer.setData('text/plain', 'person:' + person.key);
                        ev.dataTransfer.effectAllowed = 'move';
                    } catch {
                        /* ignore */
                    }
                });
                chip.addEventListener('dragend', function () {
                    dragKind = '';
                    dragPerson = null;
                    chip.classList.remove('is-dragging');
                });
            });
            const moreBtn = peopleStrip.querySelector('[data-area-people-more]');
            if (moreBtn) {
                moreBtn.addEventListener('click', function () {
                    editor.openAreaEditor(sec);
                });
            }
        }
        peopleStrip.addEventListener('dragover', function (ev) {
            if (dragKind !== 'person') return;
            ev.preventDefault();
            peopleStrip.classList.add('is-person-drop');
        });
        peopleStrip.addEventListener('dragleave', function () {
            peopleStrip.classList.remove('is-person-drop');
        });
        peopleStrip.addEventListener('drop', function (ev) {
            peopleStrip.classList.remove('is-person-drop');
            if (dragKind !== 'person' || !dragPerson) return;
            ev.preventDefault();
            const result = editor.movePersonToNode(dragPerson, sec);
            dragKind = '';
            dragPerson = null;
            if (!result.ok) {
                toast(result.message || 'Verschieben nicht möglich.', 'warning');
                return;
            }
            toast(result.message || 'Person zugeordnet.', 'success');
            render();
        });
        el.appendChild(peopleStrip);

        const grid = document.createElement('div');
        grid.className = 'schul-org__grid';
        grid.dataset.sectionId = sec.id;
        grid.addEventListener('dragover', function (ev) {
            if (dragKind !== 'card' && dragKind !== 'person') return;
            ev.preventDefault();
            grid.classList.add('is-drop-target');
        });
        grid.addEventListener('dragleave', function () {
            grid.classList.remove('is-drop-target');
        });
        grid.addEventListener('drop', function (ev) {
            grid.classList.remove('is-drop-target');
            if (dragKind === 'person' && dragPerson) {
                ev.preventDefault();
                const result = editor.movePersonToNode(dragPerson, sec);
                dragKind = '';
                dragPerson = null;
                if (!result.ok) {
                    toast(result.message || 'Verschieben nicht möglich.', 'warning');
                    return;
                }
                toast(result.message || 'Person zugeordnet.', 'success');
                render();
                return;
            }
            if (dragKind !== 'card' || !dragNodeId) return;
            ev.preventDefault();
            const settings = loadSettings();
            const next = moveSchulOrganigrammNode(
                settings.schulOrganigramm,
                dragNodeId,
                sec.id,
                ''
            );
            dragNodeId = '';
            dragKind = '';
            persistLayout(next);
        });
        if (!leafKids.length) {
            const p = document.createElement('p');
            p.className = 'muted schul-org__empty';
            p.textContent = textFilter.trim()
                ? 'Keine Treffer in diesem Bereich.'
                : nestedAreas.length
                  ? 'Unterbereiche unten – Rollen/Karten hier oder Personen oben.'
                  : 'Noch keine Einträge in diesem Bereich.';
            grid.appendChild(p);
        } else {
            leafKids.forEach(function (k) {
                // Kursteams hängen unter Klassen – in der Liste darunter zeigen
                if (k.type === NODE.KURSTEAM) return;
                grid.appendChild(renderCard(k, sec.id));
                if (k.type === NODE.CLASS) {
                    let kursKids = childrenOf(layout, k.id).filter(function (ch) {
                        return ch.type === NODE.KURSTEAM;
                    });
                    if (textFilter.trim()) {
                        const classMatch = nodeMatchesSchulOrgFilter(k, textFilter, metaLine(k));
                        if (!classMatch) {
                            kursKids = kursKids.filter(function (ch) {
                                return nodeMatchesSchulOrgFilter(ch, textFilter, metaLine(ch));
                            });
                        }
                    }
                    kursKids.forEach(function (kt) {
                        const nested = renderCard(kt, k.id);
                        nested.classList.add('schul-org__card--nested-kursteam');
                        grid.appendChild(nested);
                    });
                }
            });
        }
        el.appendChild(grid);
        nestedAreas.forEach(function (nested) {
            el.appendChild(renderSection(nested, layout));
        });
        return el;
    }

    function renderCard(node, sectionId) {
        const card = document.createElement('article');
        const occ = occupancyClass(node);
        card.className =
            'schul-org__card schul-org__card--' + (node.type || 'custom') + (occ ? ' ' + occ : '');
        card.dataset.nodeId = node.id;
        card.dataset.refKind = node.refKind || '';
        card.dataset.refKey = node.refKey || '';
        const settingsNow = loadSettings();
        const isBoard =
            (node.type === NODE.COUNCIL || node.type === NODE.SGA) &&
            (node.refKey === SCHUL_ORG_BOARD_REF ||
                node.id === 'council_board' ||
                node.id === 'sga_board' ||
                !!(node.meta && node.meta.board));
        const peoplePreview = previewPeopleForSchulOrgNode(
            settingsNow,
            node,
            isBoard ? 12 : SCHUL_ORG_PEOPLE_PREVIEW_LIMIT
        );
        const acceptsPeople =
            (node.type === NODE.ROLE && node.refKind === 'adminRole') ||
            (isArgeNode(node) && node.refKind === 'arge') ||
            (isClassNode(node) && node.refKind === 'class');

        card.draggable = !textFilter.trim();
        card.addEventListener('dragstart', function (ev) {
            const t = /** @type {HTMLElement} */ (ev.target);
            if (t && (t.closest('button') || t.closest('.schul-org__person'))) {
                ev.preventDefault();
                return;
            }
            dragNodeId = node.id;
            dragKind = 'card';
            dragPerson = null;
            card.classList.add('is-dragging');
            try {
                ev.dataTransfer.setData('text/plain', 'card:' + node.id);
                ev.dataTransfer.effectAllowed = 'move';
            } catch {
                /* ignore */
            }
        });
        card.addEventListener('dragend', function () {
            dragNodeId = '';
            dragKind = '';
            dragPerson = null;
            card.classList.remove('is-dragging');
            host.querySelectorAll('.is-drop-target, .is-person-drop').forEach(function (n) {
                n.classList.remove('is-drop-target', 'is-person-drop');
            });
        });
        card.addEventListener('dragover', function (ev) {
            if (dragKind === 'person' && acceptsPeople) {
                ev.preventDefault();
                ev.stopPropagation();
                card.classList.add('is-person-drop');
                return;
            }
            if (dragKind !== 'card' || dragNodeId === node.id) return;
            ev.preventDefault();
            ev.stopPropagation();
            card.classList.add('is-drop-target');
        });
        card.addEventListener('dragleave', function () {
            card.classList.remove('is-drop-target', 'is-person-drop');
        });
        card.addEventListener('drop', function (ev) {
            card.classList.remove('is-drop-target', 'is-person-drop');
            if (dragKind === 'person' && dragPerson && acceptsPeople) {
                ev.preventDefault();
                ev.stopPropagation();
                const result = editor.movePersonToNode(dragPerson, node);
                dragPerson = null;
                dragKind = '';
                if (result && result.ok) {
                    render();
                } else {
                    toast((result && result.message) || 'Verschieben nicht möglich.', 'warning');
                }
                return;
            }
            if (dragKind !== 'card' || !dragNodeId || dragNodeId === node.id) return;
            ev.preventDefault();
            ev.stopPropagation();
            const settings = loadSettings();
            const next = moveSchulOrganigrammNode(
                settings.schulOrganigramm,
                dragNodeId,
                sectionId || node.parentId,
                node.id
            );
            dragNodeId = '';
            dragKind = '';
            persistLayout(next);
        });
        const meta = metaLine(node);
        const editableRole = node.type === NODE.ROLE && node.refKind === 'adminRole' && node.refKey;
        const editableArge = isArgeNode(node) && node.refKind === 'arge' && node.refKey;
        const editableClass = isClassNode(node) && node.refKind === 'class' && node.refKey;
        const editableCouncil =
            node.type === NODE.COUNCIL && node.refKind === 'studentCouncil' && node.refKey;
        const editableSga = node.type === NODE.SGA && node.refKind === 'sga' && node.refKey;
        const tone = occupancyTone(node);
        const descBits = [];
        if (node.refKey && node.refKey !== SCHUL_ORG_BOARD_REF) descBits.push(String(node.refKey));
        if (node.type === NODE.CLASS && node.meta && node.meta.year) descBits.push(String(node.meta.year));
        else if (isBoard && meta) descBits.push(meta);
        else if (
            meta &&
            node.type !== NODE.ROLE &&
            !isArgeNode(node) &&
            !isClassNode(node) &&
            !isBoard
        ) {
            descBits.push(meta);
        }
        const desc =
            descBits.filter(Boolean).join(' · ') ||
            (isBoard ? meta || '' : 'Stammdaten-Eintrag');
        let statusLabel = '';
        if (peoplePreview.total > 0) {
            statusLabel =
                peoplePreview.total === 1 ? '1 Person' : peoplePreview.total + ' Personen';
        } else if (
            node.type === NODE.ROLE ||
            isArgeNode(node) ||
            isClassNode(node)
        ) {
            statusLabel = meta || 'unbesetzt';
        } else if (node.type === NODE.KURSTEAM) {
            statusLabel = meta || 'Unterrichtsteam';
        } else if (node.type === NODE.COUNCIL || node.type === NODE.SGA) {
            statusLabel = meta || '–';
        }

        function personChipHtml(p) {
            const userLinked = personHasM365UserLink(p);
            const cls =
                'schul-org__person' +
                (p.kind === 'head' ? ' schul-org__person--head' : '') +
                (p.kind === 'student' || p.scope === 'student' ? ' schul-org__person--student' : '') +
                (p.scope === 'teacher' ? ' schul-org__person--teacher' : '') +
                (p.scope === 'external' ? ' schul-org__person--external' : '') +
                (p.movable ? ' schul-org__person--movable' : '') +
                (userLinked ? ' is-m365-user' : '');
            return (
                '<span class="' +
                cls +
                '" draggable="' +
                (p.movable && !textFilter.trim() ? 'true' : 'false') +
                '" data-person-key="' +
                escapeHtml(p.key) +
                '"' +
                (userLinked ? ' data-m365-user="1"' : '') +
                ' title="' +
                escapeHtml(
                    (p.name || p.label) +
                        (p.email ? ' · ' + p.email : '') +
                        (userLinked ? ' · M365-Konto' : p.email ? ' · nur lokal' : '') +
                        (p.klasse ? ' · ' + p.klasse : '') +
                        (p.movable ? ' · ziehen zum Verschieben' : '')
                ) +
                '">' +
                userThumbHtml(p.email, p.name || p.label) +
                '<span class="schul-org__person-name">' +
                escapeHtml(p.label) +
                '</span>' +
                (userLinked
                    ? '<span class="schul-org__m365-user" title="M365-Konto" aria-label="Mit Microsoft 365 verknüpft"><i class="bi bi-microsoft" aria-hidden="true"></i></span>'
                    : '') +
                '</span>'
            );
        }

        let peopleHtml = '';
        if (isBoard && node.type === NODE.SGA) {
            const groups = [
                {
                    key: 'student',
                    label: 'Schüler',
                    items: peoplePreview.all.filter(function (p) {
                        return p.scope === 'student';
                    })
                },
                {
                    key: 'teacher',
                    label: 'Lehrer',
                    items: peoplePreview.all.filter(function (p) {
                        return p.scope === 'teacher';
                    })
                },
                {
                    key: 'external',
                    label: 'Eltern',
                    items: peoplePreview.all.filter(function (p) {
                        return p.scope === 'external';
                    })
                }
            ];
            peopleHtml =
                '<div class="schul-org__people-groups" data-people-host="1">' +
                groups
                    .map(function (g) {
                        const shown = g.items.slice(0, 4);
                        const more = Math.max(0, g.items.length - shown.length);
                        return (
                            '<div class="schul-org__people-group">' +
                            '<div class="schul-org__people-group-label">' +
                            escapeHtml(g.label) +
                            ' <em>' +
                            g.items.length +
                            '</em></div>' +
                            '<div class="schul-org__people">' +
                            (shown.length
                                ? shown.map(personChipHtml).join('')
                                : '<span class="schul-org__people-empty">–</span>') +
                            (more
                                ? '<button type="button" class="schul-org__person schul-org__person--more" data-people-more="1">+' +
                                  more +
                                  '</button>'
                                : '') +
                            '</div></div>'
                        );
                    })
                    .join('') +
                '</div>';
            if (card.className.indexOf('schul-org__card--board') === -1) {
                card.className += ' schul-org__card--board schul-org__card--sga-board';
            }
        } else {
            peopleHtml =
                '<div class="schul-org__people" data-people-host="1">' +
                (peoplePreview.shown.length
                    ? peoplePreview.shown.map(personChipHtml).join('')
                    : acceptsPeople || isBoard
                      ? '<span class="schul-org__people-empty">Keine Personen</span>'
                      : '') +
                (peoplePreview.more
                    ? '<button type="button" class="schul-org__person schul-org__person--more" data-people-more="1" title="Alle Mitglieder im Bearbeiten-Dialog">' +
                      '+' +
                      peoplePreview.more +
                      '</button>'
                    : '') +
                '</div>';
            if (isBoard) card.className += ' schul-org__card--board';
        }

        const cardGroupId = resolveGroupIdForNode(node, settingsNow);
        const cardM365 = nodeM365LinkFlags(node, cardGroupId);
        const cardIcon = cardGroupId
            ? groupThumbHtml(cardGroupId, node.label)
            : '<i class="bi ' + cardIconForType(node.type, node) + '" aria-hidden="true"></i>';
        if (cardM365.linked) {
            card.classList.add('is-m365-linked');
            card.dataset.m365Linked = '1';
        }

        card.innerHTML =
            '<div class="schul-org__card-body">' +
            '<div class="schul-org__card-head">' +
            cardIcon +
            '<div class="schul-org__card-titles">' +
            '<div class="schul-org__card-title">' +
            escapeHtml(node.label) +
            (cardM365.linked
                ? ' <span class="schul-org__m365-badge" title="' +
                  escapeHtml(
                      [
                          cardM365.hasGroup ? 'M365-Gruppe' : '',
                          cardM365.hasMailbox ? 'Postfach' : ''
                      ]
                          .filter(Boolean)
                          .join(' · ') || 'M365 verknüpft'
                  ) +
                  '"><i class="bi bi-microsoft" aria-hidden="true"></i> M365</span>'
                : '') +
            '</div>' +
            '<div class="schul-org__card-desc">' +
            escapeHtml(desc) +
            '</div>' +
            '</div></div>' +
            peopleHtml +
            '<div class="schul-org__card-foot">' +
            (statusLabel
                ? '<span class="schul-org__badge"' +
                  (tone ? ' data-tone="' + tone + '"' : '') +
                  '>' +
                  escapeHtml(statusLabel) +
                  '</span>'
                : '<span></span>') +
            '<button type="button" class="schul-org__card-action" hidden>Bearbeiten</button>' +
            '</div></div>';

        function attachEdit(openFn, title) {
            card.classList.add('schul-org__card--editable');
            card.title =
                'Doppelklick oder „Bearbeiten“: ' +
                title +
                ' · Personen ziehen · Karte sortieren';
            const btn = /** @type {HTMLButtonElement|null} */ (
                card.querySelector('.schul-org__card-action')
            );
            if (btn) {
                btn.hidden = false;
                btn.setAttribute('aria-label', title);
                btn.title = title;
                btn.addEventListener('click', function (ev) {
                    ev.stopPropagation();
                    openFn(node.refKey);
                });
            }
            card.addEventListener('dblclick', function (ev) {
                if (
                    ev.target &&
                    /** @type {HTMLElement} */ (ev.target).closest('.schul-org__person')
                ) {
                    return;
                }
                openFn(node.refKey);
            });
            const moreBtn = card.querySelector('[data-people-more]');
            if (moreBtn) {
                moreBtn.addEventListener('click', function (ev) {
                    ev.stopPropagation();
                    openFn(node.refKey);
                });
            }
        }

        card.querySelectorAll('.schul-org__person--movable').forEach(function (chip) {
            chip.addEventListener('dragstart', function (ev) {
                ev.stopPropagation();
                const key = chip.getAttribute('data-person-key') || '';
                const person = peoplePreview.all.find(function (p) {
                    return p.key === key;
                });
                if (!person || !person.movable) {
                    ev.preventDefault();
                    return;
                }
                dragKind = 'person';
                dragNodeId = '';
                dragPerson = {
                    kind: person.kind,
                    key: person.key,
                    name: person.name,
                    email: person.email,
                    label: person.label,
                    sourceNodeId: node.id,
                    sourceRefKind: node.refKind || '',
                    sourceRefKey: node.refKey || ''
                };
                chip.classList.add('is-dragging');
                card.classList.add('is-person-source');
                try {
                    ev.dataTransfer.setData('text/plain', 'person:' + person.key);
                    ev.dataTransfer.effectAllowed = 'move';
                } catch {
                    /* ignore */
                }
            });
            chip.addEventListener('dragend', function () {
                chip.classList.remove('is-dragging');
                card.classList.remove('is-person-source');
                dragKind = '';
                dragPerson = null;
                host.querySelectorAll('.is-drop-target, .is-person-drop').forEach(function (n) {
                    n.classList.remove('is-drop-target', 'is-person-drop');
                });
            });
        });

        if (editableRole) attachEdit(editor.openRoleEditor, 'Rolle bearbeiten');
        else if (editableArge) attachEdit(editor.openArgeEditor, 'ARGE / Leitung bearbeiten');
        else if (editableClass) attachEdit(editor.openClassEditor, 'Klasse / KV bearbeiten');
        else if (editableCouncil) attachEdit(editor.openCouncilEditor, 'Schülervertretung bearbeiten');
        else if (editableSga) attachEdit(editor.openSgaEditor, 'SGA bearbeiten');
        else card.title = 'Ziehen zum Sortieren';
        return card;
    }

    function flushLayoutToStammdaten() {
        const cur = loadSettings();
        let layoutNow = null;
        if (mapApi) {
            try {
                if (typeof mapApi.flushLandkartePositions === 'function') {
                    layoutNow = mapApi.flushLandkartePositions();
                } else if (typeof mapApi.getLayout === 'function') {
                    layoutNow = mapApi.getLayout();
                }
            } catch {
                layoutNow = null;
            }
        }
        if (!layoutNow) layoutNow = cur.schulOrganigramm;
        if (!layoutNow) return null;
        return saveLayout(normalizeSchulOrganigramm(layoutNow));
    }

    try {
        window.ms365FlushSchulOrganigrammBeforeBackup = flushLayoutToStammdaten;
    } catch {
        /* ignore */
    }

    render();
    return {
        refresh: render,
        getLayout: function () {
            const settings = loadSettings();
            return enrichLayoutMeta(
                resolveSchulOrganigrammLayout(settings.schulOrganigramm, settings).layout,
                settings
            );
        },
        flushBeforeBackup: flushLayoutToStammdaten
    };
}

if (typeof window !== 'undefined') {
    window.ms365MountSchulOrganigramm = mountSchulOrganigramm;
}

export default { mountSchulOrganigramm };
