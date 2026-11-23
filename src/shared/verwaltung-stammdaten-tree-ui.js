/**
 * Stammdaten → Verwaltung: Baum wie Landkarte.
 * Verschieben: Funktion/Bereich/Person → Zielbereich (direkt speichern, Sync einmal aussetzen).
 */
import { buildVerwaltungStammdatenTree } from './verwaltung-stammdaten-tree.js';
import {
    createSchulOrganigrammEditor,
    moveSchulOrgRole,
    moveSchulOrgPerson
} from '../tools/schul-organigramm/schul-organigramm-edit.js';
import {
    moveSchulOrganigrammArea,
    normalizeSchulOrganigramm,
    applySchulOrganigrammToStammdaten,
    rebuildSchulOrganigrammFromStammdaten,
    isAreaNode
} from '../tools/schul-organigramm/schul-organigramm-logic.js';

/** Kern-Vorlagen, wenn Verwaltung noch leer ist (UX-Maßnahme 5). */
const VERWALTUNG_ROLE_TEMPLATES = [
    { name: 'Direktion', code: 'DIREKTION', tier: 'schulleitung' },
    { name: 'Sekretariat', code: 'SEKRETARIAT', tier: 'verwaltung' },
    { name: 'Administration', code: 'ADMINISTRATION', tier: 'verwaltung' }
];
import {
    ensureAdminAudienceOnSettings,
    normalizeVerwaltungAudienceGroups
} from './administration-audience-groups.js';
import { buildAudienceGroupM365Panel } from './tenant-admin-audience-m365.js';

const DRAG_THRESHOLD_PX = 6;

/**
 * @param {{
 *   host: HTMLElement,
 *   getSettings: () => object,
 *   saveSettings?: (settings: object) => void,
 *   onOpenOrganigramm?: () => void,
 *   onStructureChanged?: () => void,
 *   scheduleAutoSave?: () => void,
 *   getSchoolName?: () => string,
 *   setSummary?: (msg: string, kind?: string) => void,
 *   onAudienceGroupsChanged?: (groups: object[]) => void
 * }} opts
 */
export function mountVerwaltungStammdatenTree(opts) {
    const host = opts && opts.host;
    if (!host) return { refresh: function () {} };

    /** @type {{ kind: string, id: string, code?: string, audienceId?: string, label?: string }|null} */
    let selected = null;
    /** @type {Map<string, { node: object, parentId: string, li: HTMLElement }>} */
    let nodeIndex = new Map();
    /** @type {object|null} */
    let currentTree = null;
    let suppressClick = false;

    /** @type {any} */
    let ptr = null;

    const editor = createSchulOrganigrammEditor({
        loadSettings: loadSettings,
        saveSettings: persistSettings,
        onChanged: function () {
            afterPersist();
        }
    });

    function loadSettings() {
        if (typeof opts.getSettings === 'function') return opts.getSettings() || {};
        if (typeof window !== 'undefined' && typeof window.ms365TenantSettingsLoad === 'function') {
            return window.ms365TenantSettingsLoad() || {};
        }
        return {};
    }

    function persistSettings(settings) {
        if (typeof window !== 'undefined') {
            window.__ms365SchulOrgSkipSyncOnce = true;
        }
        if (typeof opts.saveSettings === 'function') {
            opts.saveSettings(settings);
            return;
        }
        if (typeof window !== 'undefined' && typeof window.ms365TenantSettingsSave === 'function') {
            window.ms365TenantSettingsSave(settings);
        }
    }

    function afterPersist() {
        if (typeof opts.onStructureChanged === 'function') opts.onStructureChanged();
        // Autosave triggert setCore → Sync. Skip erneut setzen, sonst kann Sync
        // Custom-Bereiche/Rollen-Parent zurücksetzen (genau der IT-Support→IT-Team-Bug).
        if (typeof opts.scheduleAutoSave === 'function') {
            if (typeof window !== 'undefined') {
                window.__ms365SchulOrgSkipSyncOnce = true;
            }
            opts.scheduleAutoSave();
        }
        render();
    }

    function applyVerwaltungRoleTemplates() {
        const settings = loadSettings();
        const defaults =
            typeof window !== 'undefined' &&
            typeof window.ms365TenantSettingsDefaultAdminRoles === 'function'
                ? window.ms365TenantSettingsDefaultAdminRoles()
                : VERWALTUNG_ROLE_TEMPLATES.slice();
        const wanted = VERWALTUNG_ROLE_TEMPLATES.map(function (t) {
            return String(t.name).toLowerCase();
        });
        const fromDefaults = (Array.isArray(defaults) ? defaults : []).filter(function (d) {
            return wanted.indexOf(String((d && d.name) || '').toLowerCase()) !== -1;
        });
        const seedRoles = fromDefaults.length ? fromDefaults : VERWALTUNG_ROLE_TEMPLATES;
        const roles = Array.isArray(settings.adminRoles) ? settings.adminRoles.slice() : [];
        const seen = new Set(
            roles.map(function (r) {
                return String((r && r.name) || '')
                    .trim()
                    .toLowerCase();
            })
        );
        seedRoles.forEach(function (d) {
            const name = String((d && d.name) || '').trim();
            const key = name.toLowerCase();
            if (!name || seen.has(key)) return;
            seen.add(key);
            roles.push({
                code: String((d && d.code) || name)
                    .toUpperCase()
                    .replace(/\s+/g, '')
                    .slice(0, 24),
                name: name,
                tier: String((d && d.tier) || 'verwaltung')
            });
        });
        const next = Object.assign({}, settings, { adminRoles: roles });
        next.schulOrganigramm = rebuildSchulOrganigrammFromStammdaten(
            next.schulOrganigramm,
            next
        );
        const applied = applySchulOrganigrammToStammdaten(next, next.schulOrganigramm);
        persistSettings((applied && applied.settings) || next);
        notifyUser('Vorlagen Direktion, Sekretariat und Administration übernommen.', 'ok');
        afterPersist();
    }

    function notifyUser(msg, kind) {
        if (typeof window !== 'undefined' && typeof window.ms365AppDialogAlert === 'function') {
            window.ms365AppDialogAlert(msg, {
                kind: kind || 'info',
                title: 'Verwaltung'
            });
            return;
        }
        if (typeof opts.setSummary === 'function') {
            opts.setSummary(msg, kind || 'info');
            return;
        }
        if (typeof window !== 'undefined' && typeof window.ms365Toast === 'function') {
            window.ms365Toast(msg, kind || 'info');
        }
    }

    function escapeHtml(s) {
        return String(s ?? '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function kindIcon(kind) {
        if (kind === 'area') return 'bi-folder2-open';
        if (kind === 'role') return 'bi-person-vcard';
        return 'bi-person';
    }

    function kindLabel(kind) {
        if (kind === 'area') return 'Bereich';
        if (kind === 'role') return 'Funktion';
        return 'Person';
    }

    function findNodeById(roots, id) {
        const want = String(id || '');
        for (let i = 0; i < (roots || []).length; i++) {
            const n = roots[i];
            if (String(n.id) === want) return n;
            const hit = findNodeById(n.children, want);
            if (hit) return hit;
        }
        return null;
    }

    function containsNodeId(node, targetId) {
        if (!node || !targetId) return false;
        if (String(node.id) === String(targetId)) return true;
        return (node.children || []).some(function (ch) {
            return containsNodeId(ch, targetId);
        });
    }

    function layoutNodeById(settings, id) {
        const nodes =
            settings &&
            settings.schulOrganigramm &&
            Array.isArray(settings.schulOrganigramm.nodes)
                ? settings.schulOrganigramm.nodes
                : [];
        const want = String(id || '');
        return (
            nodes.find(function (n) {
                return n && String(n.id) === want;
            }) || null
        );
    }

    /**
     * Was soll strukturell verschoben werden?
     * Person unter Funktion → ganze Funktion (sonst bleibt die Karte „IT-Support“ stehen).
     * @param {object} node
     */
    function resolveMovePayload(node) {
        if (!node) return null;
        if (node.kind === 'role') {
            return { mode: 'role', id: String(node.id), label: node.label || 'Funktion', node: node };
        }
        if (node.kind === 'area') {
            return { mode: 'area', id: String(node.id), label: node.label || 'Bereich', node: node };
        }
        if (node.kind === 'person' && node.source === 'role' && node.sourceRoleId) {
            const role = nodeIndex.get(String(node.sourceRoleId));
            const roleNode = role ? role.node : findNodeById(
                buildVerwaltungStammdatenTree(loadSettings()).roots,
                node.sourceRoleId
            );
            return {
                mode: 'role',
                id: String(node.sourceRoleId),
                label: (roleNode && roleNode.label) || 'Funktion',
                node: roleNode || node
            };
        }
        if (node.kind === 'person') {
            return { mode: 'person', id: String(node.id), label: node.label || 'Person', node: node };
        }
        return null;
    }

    /**
     * @param {object} payload from resolveMovePayload
     * @param {string} targetAreaId
     */
    function moveIntoArea(payload, targetAreaId) {
        const areaId = String(targetAreaId || '').trim();
        if (!payload || !areaId) {
            notifyUser('Bitte einen Zielbereich wählen.', 'warning');
            return false;
        }
        const cur = loadSettings();
        const target = layoutNodeById(cur, areaId);
        if (!target || !isAreaNode(target)) {
            notifyUser('Zielbereich nicht gefunden.', 'warning');
            return false;
        }

        if (payload.mode === 'role') {
            const result =
                typeof editor.moveRoleToArea === 'function'
                    ? editor.moveRoleToArea(payload.id, areaId, '')
                    : moveSchulOrgRole(cur, payload.id, areaId, '');
            if (!result || !result.ok || !result.settings) {
                notifyUser(
                    (result && result.message) || 'Funktion kann hier nicht abgelegt werden.',
                    'warning'
                );
                return false;
            }
            // editor.moveRoleToArea speichert bereits; Fallback-Pfad speichern
            if (typeof editor.moveRoleToArea !== 'function') {
                persistSettings(result.settings);
            }
            const check = layoutNodeById(loadSettings(), payload.id);
            if (!check || String(check.parentId || '') !== areaId) {
                notifyUser(
                    'Verschieben wurde nicht übernommen – bitte Seite neu laden und erneut versuchen.',
                    'warning'
                );
                afterPersist();
                return false;
            }
            notifyUser(
                '„' + (payload.label || 'Funktion') + '“ nach „' + (target.label || areaId) + '“ verschoben.',
                'success'
            );
            afterPersist();
            return true;
        }

        if (payload.mode === 'area') {
            if (String(payload.id) === areaId) {
                notifyUser('Bereich kann nicht in sich selbst.', 'warning');
                return false;
            }
            if (payload.node && containsNodeId(payload.node, areaId)) {
                notifyUser('Bereich kann nicht in einen Unterbereich von sich selbst.', 'warning');
                return false;
            }
            if (typeof editor.moveAreaToParent === 'function') {
                const result = editor.moveAreaToParent(payload.id, areaId, '');
                if (!result || !result.ok) {
                    notifyUser(
                        (result && result.message) || 'Bereich kann hier nicht abgelegt werden.',
                        'warning'
                    );
                    return false;
                }
            } else {
                const layout = normalizeSchulOrganigramm(cur.schulOrganigramm);
                const moved = moveSchulOrganigrammArea(layout, payload.id, areaId, '');
                const after = moved.nodes.find(function (n) {
                    return n.id === payload.id;
                });
                if (!after || String(after.parentId || '') !== areaId) {
                    notifyUser('Bereich kann hier nicht abgelegt werden.', 'warning');
                    return false;
                }
                cur.schulOrganigramm = moved;
                const applied = applySchulOrganigrammToStammdaten(cur, moved);
                persistSettings(applied.settings);
            }
            notifyUser(
                '„' + (payload.label || 'Bereich') + '“ nach „' + (target.label || areaId) + '“ verschoben.',
                'success'
            );
            afterPersist();
            return true;
        }

        if (payload.mode === 'person' && payload.node) {
            const src = payload.node;
            const person = {
                kind: src.source === 'area' ? 'areaperson' : 'admin',
                email: src.email || '',
                name: src.label || '',
                label: src.label || '',
                sourceRefKind:
                    src.source === 'role' ? 'adminRole' : src.source === 'area' ? 'area' : '',
                sourceRefKey:
                    src.source === 'role' ? src.sourceRoleCode || '' : src.sourceAreaId || '',
                sourceNodeId:
                    src.source === 'area' ? src.sourceAreaId || '' : src.sourceRoleId || ''
            };
            const result = moveSchulOrgPerson(cur, person, target);
            if (!result || !result.ok || !result.settings) {
                notifyUser(
                    (result && result.message) || 'Person kann hier nicht abgelegt werden.',
                    'warning'
                );
                return false;
            }
            persistSettings(result.settings);
            notifyUser(
                '„' + (payload.label || 'Person') + '“ nach „' + (target.label || areaId) + '“ verschoben.',
                'success'
            );
            afterPersist();
            return true;
        }
        return false;
    }

    function listAreaOptions(roots, excludeId) {
        /** @type {Array<{ id: string, label: string }>} */
        const out = [];
        const skip = String(excludeId || '');
        function walk(nodes, prefix) {
            (nodes || []).forEach(function (n) {
                if (!n || n.kind !== 'area') return;
                if (String(n.id) === skip) return;
                if (skip) {
                    const excl = findNodeById(roots, skip);
                    if (excl && containsNodeId(excl, n.id)) return;
                }
                const label = (prefix ? prefix + ' › ' : '') + (n.label || n.id);
                out.push({ id: String(n.id), label: label });
                walk(
                    (n.children || []).filter(function (c) {
                        return c.kind === 'area';
                    }),
                    label
                );
            });
        }
        walk(roots, '');
        return out;
    }

    function clearDropHighlights() {
        host.querySelectorAll('.is-drop-target, .is-dragging').forEach(function (el) {
            el.classList.remove('is-drop-target', 'is-dragging');
        });
    }

    function endPointerDrag() {
        if (ptr && ptr.ghost && ptr.ghost.parentNode) ptr.ghost.parentNode.removeChild(ptr.ghost);
        clearDropHighlights();
        host.classList.remove('is-tree-dragging');
        document.removeEventListener('pointermove', onPtrMove, true);
        document.removeEventListener('pointerup', onPtrUp, true);
        document.removeEventListener('pointercancel', onPtrUp, true);
        ptr = null;
    }

    function hitAreaFromPoint(clientX, clientY) {
        const el = document.elementFromPoint(clientX, clientY);
        if (!el || !host.contains(el)) return null;
        const li = el.closest('.ts-vw-tree__item[data-kind="area"]');
        if (!li || !host.contains(li)) return null;
        const id = li.getAttribute('data-id') || '';
        return nodeIndex.get(id) || null;
    }

    function onPtrMove(ev) {
        if (!ptr || ev.pointerId !== ptr.pointerId) return;
        const dist = Math.hypot(ev.clientX - ptr.startX, ev.clientY - ptr.startY);
        if (!ptr.moved && dist < DRAG_THRESHOLD_PX) return;
        if (!ptr.moved) {
            ptr.moved = true;
            suppressClick = true;
            host.classList.add('is-tree-dragging');
            const src = nodeIndex.get(ptr.payload.id);
            if (src) src.li.classList.add('is-dragging');
            const ghost = document.createElement('div');
            ghost.className = 'ts-vw-tree__drag-ghost';
            ghost.textContent = '→ ' + (ptr.payload.label || ptr.payload.id);
            document.body.appendChild(ghost);
            ptr.ghost = ghost;
        }
        if (ptr.ghost) {
            ptr.ghost.style.left = ev.clientX + 12 + 'px';
            ptr.ghost.style.top = ev.clientY + 12 + 'px';
        }
        clearDropHighlights();
        const src = nodeIndex.get(ptr.payload.id);
        if (src) src.li.classList.add('is-dragging');
        const hit = hitAreaFromPoint(ev.clientX, ev.clientY);
        if (hit && String(hit.node.id) !== String(ptr.payload.id)) {
            hit.li.classList.add('is-drop-target');
        }
    }

    function onPtrUp(ev) {
        if (!ptr || ev.pointerId !== ptr.pointerId) return;
        const payload = ptr.payload;
        const moved = ptr.moved;
        const hit = moved ? hitAreaFromPoint(ev.clientX, ev.clientY) : null;
        endPointerDrag();
        if (moved && hit) {
            moveIntoArea(payload, hit.node.id);
        } else if (moved) {
            notifyUser('Auf einen Bereich (Ordner) ablegen – z. B. IT-Team.', 'warning');
        }
        window.setTimeout(function () {
            suppressClick = false;
        }, 0);
    }

    function bindPointerDrag(row, node) {
        row.classList.add('is-draggable');
        row.addEventListener('pointerdown', function (ev) {
            if (ev.button !== 0) return;
            if (ev.target && /** @type {Element} */ (ev.target).closest('.ts-vw-tree__toggle')) {
                return;
            }
            if (ev.target && /** @type {Element} */ (ev.target).closest('[data-move-here]')) {
                return;
            }
            const payload = resolveMovePayload(node);
            if (!payload) return;
            if (ptr) endPointerDrag();
            ptr = {
                payload: payload,
                startX: ev.clientX,
                startY: ev.clientY,
                moved: false,
                ghost: null,
                pointerId: ev.pointerId
            };
            document.addEventListener('pointermove', onPtrMove, true);
            document.addEventListener('pointerup', onPtrUp, true);
            document.addEventListener('pointercancel', onPtrUp, true);
        });
    }

    function selectNode(node) {
        if (suppressClick) return;
        if (!node) {
            selected = null;
            render();
            return;
        }
        selected = {
            kind: node.kind,
            id: node.id,
            code: node.code || '',
            audienceId: node.audienceId || '',
            label: node.label || ''
        };
        render();
    }

    function renderTreeNode(node, parentId) {
        const li = document.createElement('li');
        li.className = 'ts-vw-tree__item ts-vw-tree__item--' + (node.kind || 'area');
        li.dataset.kind = node.kind || '';
        li.dataset.id = node.id || '';
        if (parentId) li.dataset.parentId = parentId;
        if (selected && String(selected.id) === String(node.id)) li.classList.add('is-selected');

        nodeIndex.set(String(node.id), {
            node: node,
            parentId: parentId ? String(parentId) : '',
            li: li
        });

        const row = document.createElement('div');
        row.className = 'ts-vw-tree__row';
        row.style.setProperty('--ts-vw-depth', String(node.depth || 0));
        row.tabIndex = 0;
        row.setAttribute('role', 'treeitem');
        row.title = 'Ziehen auf einen Bereich oder rechts „Verschieben nach…“';

        const hasKids = Array.isArray(node.children) && node.children.length > 0;
        if (hasKids) {
            const tog = document.createElement('button');
            tog.type = 'button';
            tog.className = 'ts-vw-tree__toggle';
            tog.setAttribute('aria-expanded', 'true');
            tog.title = 'Ein-/ausklappen';
            tog.innerHTML = '<i class="bi bi-chevron-down" aria-hidden="true"></i>';
            tog.addEventListener('click', function (ev) {
                ev.stopPropagation();
                const open = li.classList.toggle('is-collapsed');
                tog.setAttribute('aria-expanded', open ? 'false' : 'true');
                const icon = tog.querySelector('i');
                if (icon) icon.className = open ? 'bi bi-chevron-right' : 'bi bi-chevron-down';
            });
            row.appendChild(tog);
        } else {
            const spacer = document.createElement('span');
            spacer.className = 'ts-vw-tree__toggle-spacer';
            row.appendChild(spacer);
        }

        const icon = document.createElement('i');
        icon.className = 'bi ' + kindIcon(node.kind);
        icon.setAttribute('aria-hidden', 'true');
        row.appendChild(icon);

        const main = document.createElement('span');
        main.className = 'ts-vw-tree__label';
        main.textContent = node.label || node.id || '–';
        row.appendChild(main);

        const meta = document.createElement('span');
        meta.className = 'ts-vw-tree__meta muted';
        if (node.kind === 'role' && node.code) meta.textContent = node.code;
        else if (node.kind === 'person' && node.email) meta.textContent = node.email;
        else if (node.kind === 'area') {
            const nRoles = (node.children || []).filter(function (c) {
                return c.kind === 'role';
            }).length;
            const nAreas = (node.children || []).filter(function (c) {
                return c.kind === 'area';
            }).length;
            const bits = [];
            if (nAreas) bits.push(nAreas + (nAreas === 1 ? ' Unterbereich' : ' Unterbereiche'));
            if (nRoles) bits.push(nRoles + (nRoles === 1 ? ' Funktion' : ' Funktionen'));
            meta.textContent = bits.join(' · ') || kindLabel(node.kind);
        } else meta.textContent = kindLabel(node.kind);
        row.appendChild(meta);

        // Wenn etwas Ausgewähltes hierher kann: Button am Bereich
        if (node.kind === 'area' && selected && currentTree) {
            const selNode =
                findNodeById(currentTree.roots, selected.id) ||
                (nodeIndex.get(String(selected.id)) &&
                    nodeIndex.get(String(selected.id)).node) ||
                selected;
            const payload = resolveMovePayload(selNode);
            let currentParent = '';
            if (payload) {
                const pe = nodeIndex.get(payload.id);
                if (pe) currentParent = pe.parentId || '';
                else {
                    // Parent aus Baum ableiten
                    function parentOf(roots, id, parent) {
                        for (let i = 0; i < (roots || []).length; i++) {
                            const n = roots[i];
                            if (String(n.id) === String(id)) return parent || '';
                            const hit = parentOf(n.children, id, n.id);
                            if (hit !== undefined && hit !== null) return hit;
                        }
                        return null;
                    }
                    const p = parentOf(currentTree.roots, payload.id, '');
                    if (p != null) currentParent = String(p);
                }
            }
            if (
                payload &&
                String(node.id) !== String(payload.id) &&
                String(node.id) !== String(currentParent)
            ) {
                const btnHere = document.createElement('button');
                btnHere.type = 'button';
                btnHere.className = 'btn btn-sm btn-success ts-vw-tree__move-here';
                btnHere.setAttribute('data-move-here', '1');
                btnHere.textContent = 'Hierher';
                btnHere.title =
                    '„' + (payload.label || '') + '“ nach „' + (node.label || '') + '“ verschieben';
                btnHere.addEventListener('click', function (ev) {
                    ev.stopPropagation();
                    moveIntoArea(payload, node.id);
                });
                row.appendChild(btnHere);
            }
        }

        row.addEventListener('click', function (ev) {
            if (ev.target && /** @type {Element} */ (ev.target).closest('[data-move-here]')) return;
            selectNode(node);
        });
        row.addEventListener('keydown', function (ev) {
            if (ev.key === 'Enter' || ev.key === ' ') {
                ev.preventDefault();
                selectNode(node);
            }
        });

        if (node.kind === 'area' || node.kind === 'role' || node.kind === 'person') {
            bindPointerDrag(row, node);
        }

        li.appendChild(row);

        if (hasKids) {
            const ul = document.createElement('ul');
            ul.className = 'ts-vw-tree__children';
            ul.setAttribute('role', 'group');
            node.children.forEach(function (ch) {
                ul.appendChild(renderTreeNode(ch, node.id));
            });
            li.appendChild(ul);
        }
        return li;
    }

    function renderDetail(settings, tree) {
        const detail = document.createElement('div');
        detail.className = 'ts-vw-tree__detail';

        if (!selected) {
            detail.classList.add('is-quiet-help');
            detail.innerHTML =
                '<details class="ts-vw-tree__howto">' +
                '<summary>Einträge verschieben</summary>' +
                '<ol class="ts-vw-tree__howto-list">' +
                '<li>Funktion wählen (z.&nbsp;B. IT-Support)</li>' +
                '<li>Am Zielbereich auf <strong>Hierher</strong> klicken – oder dorthin ziehen</li>' +
                '</ol>' +
                '</details>';
            return detail;
        }

        const live = findNodeById(tree.roots, selected.id) || selected;
        const payload = resolveMovePayload(live);

        const head = document.createElement('header');
        head.className = 'ts-vw-tree__detail-head';
        head.innerHTML =
            '<span class="ts-vw-tree__detail-icon"><i class="bi ' +
            kindIcon(live.kind) +
            '" aria-hidden="true"></i></span>' +
            '<div><h4 class="ts-vw-tree__detail-title"></h4>' +
            '<p class="muted ts-vw-tree__detail-sub"></p></div>';
        const titleEl = head.querySelector('.ts-vw-tree__detail-title');
        const subEl = head.querySelector('.ts-vw-tree__detail-sub');
        if (titleEl) titleEl.textContent = live.label || live.id || '–';
        if (subEl) {
            if (live.kind === 'person' && payload && payload.mode === 'role') {
                subEl.textContent =
                    'Person · Verschieben nimmt die Funktion „' + payload.label + '“ mit';
            } else if (live.kind === 'area') {
                subEl.textContent = live.audienceId
                    ? 'Verwaltungsbereich · Microsoft-365-Zielgruppe'
                    : 'Bereich auf der Landkarte';
            } else if (live.kind === 'role') {
                subEl.textContent = 'Funktion' + (live.code ? ' · ' + live.code : '');
            } else {
                subEl.textContent = live.email || 'Person';
            }
        }
        detail.appendChild(head);

        const actions = document.createElement('div');
        actions.className = 'ts-vw-tree__detail-actions';

        if (live.kind === 'area') {
            const btnEdit = document.createElement('button');
            btnEdit.type = 'button';
            btnEdit.className = 'btn btn-primary';
            btnEdit.innerHTML =
                '<i class="bi bi-pencil-square" aria-hidden="true"></i> Bereich bearbeiten';
            btnEdit.addEventListener('click', function () {
                editor.openAreaEditor(live.id);
            });
            actions.appendChild(btnEdit);
            const btnRole = document.createElement('button');
            btnRole.type = 'button';
            btnRole.className = 'btn btn-success';
            btnRole.innerHTML =
                '<i class="bi bi-person-badge" aria-hidden="true"></i> Funktion hier anlegen';
            btnRole.addEventListener('click', function () {
                void Promise.resolve(editor.createRoleUnderSection(live.id)).then(function () {
                    render();
                });
            });
            actions.appendChild(btnRole);
        } else if (live.kind === 'role') {
            const btnEdit = document.createElement('button');
            btnEdit.type = 'button';
            btnEdit.className = 'btn btn-primary';
            btnEdit.innerHTML =
                '<i class="bi bi-pencil-square" aria-hidden="true"></i> Funktion bearbeiten';
            btnEdit.addEventListener('click', function () {
                const key = live.code || String(live.id || '').replace(/^role_/i, '');
                editor.openRoleEditor(key);
            });
            actions.appendChild(btnEdit);
        }

        if (typeof opts.onOpenOrganigramm === 'function') {
            const btnMap = document.createElement('button');
            btnMap.type = 'button';
            btnMap.className = 'btn btn-sm';
            btnMap.innerHTML =
                '<i class="bi bi-diagram-3" aria-hidden="true"></i> Auf der Landkarte';
            btnMap.addEventListener('click', function () {
                opts.onOpenOrganigramm();
            });
            actions.appendChild(btnMap);
        }
        detail.appendChild(actions);

        // Klare Ziel-Buttons
        if (payload) {
            const excludeId = payload.mode === 'area' ? payload.id : '';
            const areas = listAreaOptions(tree.roots, excludeId);
            const currentParent =
                payload.mode === 'role' || payload.mode === 'area'
                    ? (nodeIndex.get(payload.id) && nodeIndex.get(payload.id).parentId) || ''
                    : '';

            const moveBox = document.createElement('div');
            moveBox.className = 'ts-vw-tree__move';
            const title = document.createElement('div');
            title.className = 'ts-vw-tree__move-label';
            title.innerHTML =
                '<i class="bi bi-arrows-move" aria-hidden="true"></i> „' +
                escapeHtml(payload.label) +
                '“ verschieben nach';
            moveBox.appendChild(title);

            const grid = document.createElement('div');
            grid.className = 'ts-vw-tree__move-targets';
            let shown = 0;
            areas.forEach(function (a) {
                if (String(a.id) === String(currentParent)) return;
                shown += 1;
                const btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'btn btn-sm btn-success';
                btn.innerHTML =
                    '<i class="bi bi-folder-symlink" aria-hidden="true"></i> ' +
                    escapeHtml(a.label);
                btn.addEventListener('click', function () {
                    moveIntoArea(payload, a.id);
                });
                grid.appendChild(btn);
            });
            if (!shown) {
                const p = document.createElement('p');
                p.className = 'muted';
                p.textContent = 'Kein anderer Bereich als Ziel verfügbar.';
                moveBox.appendChild(p);
            } else {
                moveBox.appendChild(grid);
            }
            const hint = document.createElement('p');
            hint.className = 'muted ts-vw-tree__move-hint';
            hint.textContent =
                'Oder links auf dem Zielordner „Hierher“ klicken bzw. dorthin ziehen.';
            moveBox.appendChild(hint);
            detail.appendChild(moveBox);
        }

        if (live.kind === 'area' && Array.isArray(live.children) && live.children.length) {
            const list = document.createElement('ul');
            list.className = 'ts-vw-tree__detail-kids';
            live.children.forEach(function (ch) {
                const li = document.createElement('li');
                li.innerHTML =
                    '<i class="bi ' +
                    kindIcon(ch.kind) +
                    '" aria-hidden="true"></i> ' +
                    escapeHtml(ch.label || ch.id);
                li.addEventListener('click', function () {
                    selectNode(ch);
                });
                list.appendChild(li);
            });
            detail.appendChild(list);
        }

        if (live.kind === 'area' && live.audienceId) {
            const st = ensureAdminAudienceOnSettings(settings || {});
            const group = (st.groups || []).find(function (g) {
                return g.id === live.audienceId;
            });
            if (group) {
                const m365Wrap = document.createElement('div');
                m365Wrap.className = 'ts-vw-tree__detail-m365';
                const m365Title = document.createElement('h5');
                m365Title.className = 'ts-vw-tree__detail-m365-title';
                m365Title.textContent = 'Microsoft 365-Zielgruppe';
                m365Wrap.appendChild(m365Title);
                const panel = buildAudienceGroupM365Panel({
                    group: group,
                    memberships: st.memberships || [],
                    schoolName:
                        (typeof opts.getSchoolName === 'function' ? opts.getSchoolName() : '') || '',
                    setSummary: opts.setSummary,
                    onGroupMetaChange: function (groupId, meta) {
                        const next = normalizeVerwaltungAudienceGroups(st.groups).map(function (g) {
                            if (g.id !== groupId) return g;
                            return Object.assign({}, g, {
                                mailNick:
                                    meta.mailNick != null ? String(meta.mailNick).trim() : g.mailNick,
                                newDisplayName:
                                    meta.newDisplayName != null
                                        ? String(meta.newDisplayName).trim()
                                        : g.newDisplayName
                            });
                        });
                        if (typeof opts.onAudienceGroupsChanged === 'function') {
                            opts.onAudienceGroupsChanged(next);
                        }
                        if (typeof opts.scheduleAutoSave === 'function') opts.scheduleAutoSave();
                        render();
                    },
                    onMatchChanged: function () {
                        render();
                    }
                });
                m365Wrap.appendChild(panel.inline);
                if (panel.aux && panel.aux.childNodes.length) m365Wrap.appendChild(panel.aux);
                detail.appendChild(m365Wrap);
            }
        }

        return detail;
    }

    function render() {
        endPointerDrag();
        const settings = loadSettings();
        const tree = buildVerwaltungStammdatenTree(settings);
        currentTree = tree;
        nodeIndex = new Map();

        if (selected) {
            const still = findNodeById(tree.roots, selected.id);
            if (!still) selected = null;
            else {
                selected = {
                    kind: still.kind,
                    id: still.id,
                    code: still.code || '',
                    audienceId: still.audienceId || '',
                    label: still.label || ''
                };
            }
        }

        host.replaceChildren();
        host.classList.add('ts-vw-tree', 'ts-vw-tree--editor');

        const head = document.createElement('div');
        head.className = 'ts-vw-tree__head';
        head.innerHTML =
            '<h3 class="ts-vw-tree__title">Struktur</h3>' +
            '<p class="muted ts-vw-tree__lead">' +
            'Eintrag wählen und am Zielordner <strong>Hierher</strong> ablegen, oder ziehen.' +
            '</p>';
        host.appendChild(head);

        const toolbar = document.createElement('div');
        toolbar.className = 'ts-vw-tree__toolbar';
        const btnArea = document.createElement('button');
        btnArea.type = 'button';
        btnArea.className = 'btn fb-btn ts-act-primary';
        btnArea.innerHTML =
            '<i class="bi bi-plus-lg" aria-hidden="true"></i> Verwaltungsgruppe';
        btnArea.addEventListener('click', function () {
            void Promise.resolve(editor.createAudienceSection()).then(function () {
                afterPersist();
            });
        });
        toolbar.appendChild(btnArea);
        const btnFn = document.createElement('button');
        btnFn.type = 'button';
        btnFn.className = 'btn fb-btn fb-btn--ghost';
        btnFn.innerHTML = '<i class="bi bi-plus-lg" aria-hidden="true"></i> Funktion';
        btnFn.addEventListener('click', function () {
            const parentId =
                selected && selected.kind === 'area'
                    ? selected.id
                    : (typeof editor.getLastRoleParent === 'function' &&
                          editor.getLastRoleParent()) ||
                      '';
            void Promise.resolve(editor.createRoleUnderSection(parentId || null)).then(function () {
                render();
            });
        });
        toolbar.appendChild(btnFn);
        if (typeof opts.onOpenOrganigramm === 'function') {
            const btnMap = document.createElement('button');
            btnMap.type = 'button';
            btnMap.className = 'btn fb-btn fb-btn--ghost ts-vw-map';
            btnMap.innerHTML =
                '<i class="bi bi-diagram-3" aria-hidden="true"></i> Landkarte';
            btnMap.addEventListener('click', function () {
                opts.onOpenOrganigramm();
            });
            toolbar.appendChild(btnMap);
        }
        host.appendChild(toolbar);

        const stats = document.createElement('p');
        stats.className = 'ts-vw-tree__stats muted';
        if (tree.source === 'empty') {
            stats.textContent = 'Noch keine Verwaltungsbereiche – zuerst eine Verwaltungsgruppe anlegen.';
        } else {
            stats.textContent =
                tree.counts.areas +
                ' Bereiche · ' +
                tree.counts.roles +
                ' Funktionen · ' +
                tree.counts.people +
                ' Personen';
        }
        host.appendChild(stats);

        const split = document.createElement('div');
        split.className = 'ts-vw-tree__split';
        const master = document.createElement('div');
        master.className = 'ts-vw-tree__master';
        if (!tree.roots.length) {
            const lead = document.createElement('p');
            lead.className = 'muted';
            lead.textContent =
                'Noch leer – Standardrollen als Vorlage. Übernehmen legt Schulleitung/Verwaltung an; Personen können danach besetzt werden.';
            master.appendChild(lead);
            const list = document.createElement('ul');
            list.className = 'ts-vw-tree__templates';
            list.setAttribute('aria-label', 'Vorlagen-Rollen');
            VERWALTUNG_ROLE_TEMPLATES.forEach(function (tpl) {
                const li = document.createElement('li');
                li.className = 'ts-vw-tree__template';
                li.innerHTML =
                    '<span class="ts-vw-tree__template-label">' +
                    escapeHtml(tpl.name) +
                    '</span>' +
                    '<span class="ts-vw-tree__template-meta">unbesetzt · Vorlage</span>';
                list.appendChild(li);
            });
            master.appendChild(list);
            const actions = document.createElement('div');
            actions.className = 'ts-vw-tree__templates-actions';
            const btnSeed = document.createElement('button');
            btnSeed.type = 'button';
            btnSeed.className = 'btn fb-btn';
            btnSeed.innerHTML =
                '<i class="bi bi-magic" aria-hidden="true"></i> Vorlagen übernehmen';
            btnSeed.addEventListener('click', function () {
                applyVerwaltungRoleTemplates();
            });
            const btnAreaEmpty = document.createElement('button');
            btnAreaEmpty.type = 'button';
            btnAreaEmpty.className = 'btn fb-btn fb-btn--ghost';
            btnAreaEmpty.innerHTML =
                '<i class="bi bi-plus-circle" aria-hidden="true"></i> Eigene Verwaltungsgruppe';
            btnAreaEmpty.addEventListener('click', function () {
                void Promise.resolve(editor.createAudienceSection()).then(function () {
                    afterPersist();
                });
            });
            actions.append(btnSeed, btnAreaEmpty);
            master.appendChild(actions);
        } else {
            const ul = document.createElement('ul');
            ul.className = 'ts-vw-tree__list';
            ul.setAttribute('role', 'tree');
            tree.roots.forEach(function (root) {
                ul.appendChild(renderTreeNode(root, ''));
            });
            master.appendChild(ul);
        }
        split.appendChild(master);
        split.appendChild(renderDetail(settings, tree));
        host.appendChild(split);
    }

    render();
    return {
        refresh: render,
        openAreaEditor: function (id) {
            editor.openAreaEditor(id);
        },
        openRoleEditor: function (key) {
            editor.openRoleEditor(key);
        }
    };
}

export default { mountVerwaltungStammdatenTree: mountVerwaltungStammdatenTree };
