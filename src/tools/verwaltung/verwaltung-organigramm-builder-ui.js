/**
 * Drag-and-Drop Organigramm-Builder für Verwaltung.
 */
import {
    ORG_NODE_GROUP,
    ORG_NODE_ROLE,
    ORG_NODE_SAMMEL,
    normalizeOrganigrammLayout,
    seedOrganigrammFromRoles,
    syncOrganigrammWithRoles,
    childrenOf,
    movePersonToRole,
    moveNode,
    addOrgNode,
    removeOrgNode,
    newOrgNodeId,
    updateOrgNodeLabel
} from './verwaltung-organigramm-logic.js';

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function normStr(v) {
    return String(v ?? '').trim();
}

function normEmail(v) {
    return String(v ?? '')
        .trim()
        .toLowerCase();
}

/**
 * @param {object} deps
 */
export function createOrganigrammBuilder(deps) {
    const d = deps || {};

    function getLayout() {
        const raw = typeof d.getLayout === 'function' ? d.getLayout() : null;
        const roles = typeof d.getRoles === 'function' ? d.getRoles() : [];
        if (!raw || !Array.isArray(raw.nodes) || !raw.nodes.length) {
            return seedOrganigrammFromRoles(roles);
        }
        return syncOrganigrammWithRoles(raw, roles);
    }

    function saveLayout(layout) {
        if (typeof d.setLayout === 'function') d.setLayout(normalizeOrganigrammLayout(layout));
    }

    function roleByCode(code) {
        const c = String(code || '').toUpperCase();
        const roles = typeof d.getRoles === 'function' ? d.getRoles() : [];
        return (
            roles.find(function (r) {
                return String((r && r.code) || '').toUpperCase() === c;
            }) || null
        );
    }

    function peopleForRoleCode(code) {
        const role = roleByCode(code);
        if (!role || typeof d.getPeopleForRole !== 'function') return [];
        return d.getPeopleForRole(role) || [];
    }

    function toast(msg) {
        if (typeof d.toast === 'function') d.toast(msg);
    }

    function passesFilter(text) {
        const q = normStr(typeof d.getFilter === 'function' ? d.getFilter() : '').toLowerCase();
        if (!q) return true;
        return String(text || '')
            .toLowerCase()
            .includes(q);
    }

    async function onAddGroup(parentId) {
        const name =
            typeof d.prompt === 'function'
                ? await d.prompt('Name der Gruppe (z. B. Fachleitungen GS, Erweiterte Schulleitung)', '', {
                      title: 'Gruppe anlegen',
                      inputLabel: 'Gruppenname'
                  })
                : null;
        const label = normStr(name);
        if (!label) return;
        const layout = getLayout();
        const next = addOrgNode(layout, {
            id: newOrgNodeId('grp'),
            type: ORG_NODE_GROUP,
            label: label,
            parentId: parentId || null,
            tone: 'custom'
        });
        saveLayout(next);
        toast('Gruppe angelegt.');
        render();
    }

    async function onAddRole(parentId) {
        const name =
            typeof d.prompt === 'function'
                ? await d.prompt('Bezeichnung der Rolle (z. B. Schulleiter, Sekretariat)', '', {
                      title: 'Rolle im Organigramm',
                      inputLabel: 'Rolle'
                  })
                : null;
        const label = normStr(name);
        if (!label) return;
        if (typeof d.createRole !== 'function') {
            toast('Rolle anlegen nicht verfügbar.');
            return;
        }
        const created = await d.createRole(label);
        if (!created || !created.code) return;
        // getLayout() sync’t fehlende Rollen automatisch – Knoten ggf. nur verschieben.
        let layout = getLayout();
        const code = String(created.code || '').toUpperCase();
        let node = layout.nodes.find(function (n) {
            return n.type === ORG_NODE_ROLE && String(n.roleCode || '').toUpperCase() === code;
        });
        if (!node) {
            layout = addOrgNode(layout, {
                id: newOrgNodeId('role'),
                type: ORG_NODE_ROLE,
                label: created.name || label,
                parentId: parentId || null,
                roleCode: created.code,
                tone: 'staff'
            });
        } else if (parentId) {
            layout = moveNode(layout, node.id, parentId);
        }
        saveLayout(layout);
        toast('Rolle angelegt.');
        render();
    }

    async function onAddSammel(parentId) {
        let picked = null;
        if (typeof d.pickM365Group === 'function') {
            picked = await d.pickM365Group({
                title: 'Sammelgruppe / Microsoft-365-Gruppe wählen'
            });
        }
        const layout = getLayout();
        if (picked && picked.id) {
            const next = addOrgNode(layout, {
                id: newOrgNodeId('sam'),
                type: ORG_NODE_SAMMEL,
                label: normStr(picked.displayName || picked.name || 'Sammelgruppe'),
                parentId: parentId || null,
                m365GroupId: String(picked.id),
                m365GroupLabel: normStr(picked.displayName || picked.name),
                m365GroupMail: normEmail(picked.mail || picked.email),
                tone: 'sammel'
            });
            saveLayout(next);
            toast('Sammelgruppe im Organigramm ergänzt.');
            render();
            return;
        }
        const name =
            typeof d.prompt === 'function'
                ? await d.prompt('Bezeichnung der Sammelgruppe (ohne M365-Bezug)', 'Sammelgruppe', {
                      title: 'Sammelgruppe',
                      inputLabel: 'Name'
                  })
                : null;
        const label = normStr(name);
        if (!label) return;
        const next = addOrgNode(layout, {
            id: newOrgNodeId('sam'),
            type: ORG_NODE_SAMMEL,
            label: label,
            parentId: parentId || null,
            tone: 'sammel'
        });
        saveLayout(next);
        toast('Sammelgruppe angelegt.');
        render();
    }

    async function onAddPersonToRole(roleCode) {
        const role = roleByCode(roleCode);
        if (!role) {
            toast('Rolle nicht gefunden.');
            return;
        }
        if (typeof d.addPeopleFromM365 !== 'function') {
            toast('M365-Personenauswahl nicht verfügbar.');
            return;
        }
        await d.addPeopleFromM365(role);
        render();
    }

    function onDropPerson(emailOrKey, toRoleCode) {
        const role = roleByCode(toRoleCode);
        if (!role) return;
        const rows = typeof d.getRows === 'function' ? d.getRows() : [];
        const result = movePersonToRole(rows, emailOrKey, role.name || role.code, role.code);
        if (!result.moved) {
            toast('Person konnte nicht verschoben werden.');
            return;
        }
        if (typeof d.setRows === 'function') d.setRows(result.rows);
        toast('Zuordnung geändert: ' + (result.person.name || result.person.email) + ' → ' + (role.name || role.code));
        render();
    }

    function onDropNode(nodeId, newParentId) {
        const layout = getLayout();
        const next = moveNode(layout, nodeId, newParentId);
        saveLayout(next);
        render();
    }

    function onDeleteNode(nodeId) {
        const layout = getLayout();
        const node = layout.nodes.find(function (n) {
            return n.id === nodeId;
        });
        if (!node) return;
        const label = node.label || node.id;
        Promise.resolve(
            typeof d.confirm === 'function'
                ? d.confirm('„' + label + '“ aus dem Organigramm entfernen? (Rollen/Personen in den Stammdaten bleiben erhalten.)', {
                      title: 'Knoten entfernen',
                      okText: 'Entfernen',
                      cancelText: 'Abbrechen'
                  })
                : true
        ).then(function (ok) {
            if (!ok) return;
            saveLayout(removeOrgNode(layout, nodeId, { cascade: true }));
            toast('Aus Organigramm entfernt.');
            render();
        });
    }

    function bindDropZone(el, onDropPayload) {
        el.addEventListener('dragover', function (ev) {
            ev.preventDefault();
            el.classList.add('is-drop-target');
        });
        el.addEventListener('dragleave', function () {
            el.classList.remove('is-drop-target');
        });
        el.addEventListener('drop', function (ev) {
            ev.preventDefault();
            el.classList.remove('is-drop-target');
            let payload = null;
            try {
                payload = JSON.parse(ev.dataTransfer.getData('application/json') || '{}');
            } catch {
                payload = null;
            }
            if (payload) onDropPayload(payload);
        });
    }

    /**
     * Einheitliche Icon-Buttons im Organigramm (Beschriftung nur als Tooltip/aria-label).
     * @param {{ icon: string, label: string, title?: string, variant?: 'primary'|'success'|'danger'|'ghost', onClick: function }} opts
     */
    function makeActionBtn(opts) {
        const btn = document.createElement('button');
        btn.type = 'button';
        const variant = opts.variant || 'ghost';
        const caption = opts.title || opts.label;
        btn.className =
            'btn btn-sm vw-org-builder__btn vw-org-builder__btn--icon vw-org-builder__btn--' +
            variant +
            (variant === 'primary' ? ' btn-primary' : '') +
            (variant === 'success' ? ' btn-success' : '') +
            (variant === 'danger' ? ' btn-danger' : '') +
            (variant === 'ghost' ? ' btn-ghost' : '');
        btn.title = caption;
        btn.setAttribute('aria-label', caption);
        btn.innerHTML = '<i class="bi ' + escapeHtml(opts.icon) + '" aria-hidden="true"></i>';
        btn.addEventListener('click', function (ev) {
            ev.stopPropagation();
            opts.onClick(ev);
        });
        return btn;
    }

    function openPersonEditor(globalIdx) {
        if (typeof d.editPerson === 'function') {
            d.editPerson(globalIdx);
            return;
        }
        toast('Person bearbeiten nicht verfügbar.');
    }

    function openGroupEditor(node) {
        if (typeof d.editGroup === 'function') {
            d.editGroup(node);
            return;
        }
        // Fallback: einfacher Prompt
        void (async function () {
            const name =
                typeof d.prompt === 'function'
                    ? await d.prompt('Bezeichnung', node.label || '', {
                          title: 'Gruppe bearbeiten',
                          inputLabel: 'Name'
                      })
                    : null;
            const label = normStr(name);
            if (!label) return;
            saveLayout(updateOrgNodeLabel(getLayout(), node.id, label));
            toast('Gruppe umbenannt.');
            render();
        })();
    }

    function renderPersonChip(person, globalIdx) {
        const chip = document.createElement('div');
        chip.className = 'vw-org-builder__person';
        chip.draggable = true;
        chip.title =
            (person.email ? person.email + ' – ' : '') +
            'Doppelklick: bearbeiten · Ziehen: Rolle ändern';
        const name = escapeHtml(person.name || person.email || '–');
        const mail =
            person.email && person.name
                ? '<span class="vw-org-builder__person-mail">' + escapeHtml(person.email) + '</span>'
                : '';
        chip.innerHTML =
            '<i class="bi bi-grip-vertical vw-org-builder__person-grip" aria-hidden="true"></i>' +
            '<span class="vw-org-builder__person-text"><span class="vw-org-builder__person-name">' +
            name +
            '</span>' +
            mail +
            '</span>';
        const key =
            person.email && String(person.email).indexOf('@') !== -1
                ? String(person.email).toLowerCase()
                : 'idx:' + String(globalIdx);
        chip.addEventListener('dragstart', function (ev) {
            ev.dataTransfer.setData(
                'application/json',
                JSON.stringify({ kind: 'person', key: key })
            );
            ev.dataTransfer.effectAllowed = 'move';
            chip.classList.add('is-dragging');
        });
        chip.addEventListener('dragend', function () {
            chip.classList.remove('is-dragging');
        });
        chip.addEventListener('dblclick', function (ev) {
            ev.preventDefault();
            ev.stopPropagation();
            openPersonEditor(globalIdx);
        });
        return chip;
    }

    function renderRoleCard(node, layout) {
        const role = roleByCode(node.roleCode);
        const people = role ? peopleForRoleCode(node.roleCode) : [];
        const filled = people.filter(function (p) {
            return normStr(p && p.name) || (normEmail(p && p.email) && String(p.email).indexOf('@') !== -1);
        });
        const n = filled.length;
        const seatClass = n === 0 ? 'is-vacant' : n === 1 ? 'is-single' : 'is-team';
        const seatLabel =
            n === 0 ? 'unbesetzt' : n === 1 ? 'Einzelrolle · 1 Person' : 'Team-Rolle · ' + n + ' Personen';
        const rows = typeof d.getRows === 'function' ? d.getRows() : [];
        const card = document.createElement('article');
        card.className = 'vw-org-builder__role ' + seatClass;
        card.dataset.nodeId = node.id;
        card.dataset.seat = n === 0 ? 'vacant' : n === 1 ? 'single' : 'team';
        card.title = seatLabel + ' – Doppelklick: Details · Karte ziehen: Gruppe wechseln';
        card.draggable = true;
        card.addEventListener('dragstart', function (ev) {
            if (ev.target.closest('.vw-org-builder__person, .vw-org-builder__btn, button')) {
                ev.preventDefault();
                return;
            }
            ev.dataTransfer.setData('application/json', JSON.stringify({ kind: 'node', id: node.id }));
            ev.dataTransfer.effectAllowed = 'move';
        });
        card.addEventListener('dblclick', function (ev) {
            if (ev.target.closest('.vw-org-builder__person, .vw-org-builder__btn, button, input')) {
                return;
            }
            ev.preventDefault();
            if (typeof d.openRole === 'function') d.openRole(node.roleCode);
        });

        const head = document.createElement('div');
        head.className = 'vw-org-builder__role-head';
        head.innerHTML =
            '<div class="vw-org-builder__role-title-row">' +
            '<strong>' +
            escapeHtml(node.label || (role && role.name) || node.roleCode) +
            '</strong>' +
            '<span class="vw-org-builder__seat-badge vw-org-builder__seat-badge--' +
            (n === 0 ? 'vacant' : n === 1 ? 'single' : 'team') +
            '" title="' +
            escapeHtml(seatLabel) +
            '">' +
            (n === 0 ? 'frei' : n === 1 ? '1×' : n + '×') +
            '</span></div>' +
            '<span class="muted vw-org-builder__role-code">' +
            escapeHtml(node.roleCode || '') +
            '</span>';
        card.appendChild(head);

        const peopleBox = document.createElement('div');
        peopleBox.className = 'vw-org-builder__people';
        peopleBox.setAttribute('aria-label', 'Personen – hierher ziehen zum Zuordnen');
        if (!people.length) {
            const vac = document.createElement('div');
            vac.className = 'vw-org-builder__vacant';
            vac.innerHTML =
                '<i class="bi bi-person-dash" aria-hidden="true"></i><span>Nicht besetzt – Person hierher ziehen</span>';
            peopleBox.appendChild(vac);
        } else {
            people.forEach(function (p) {
                const globalIdx = rows.indexOf(p);
                peopleBox.appendChild(renderPersonChip(p, globalIdx));
            });
        }
        bindDropZone(peopleBox, function (payload) {
            if (payload.kind === 'person') onDropPerson(payload.key, node.roleCode);
        });
        card.appendChild(peopleBox);

        const actions = document.createElement('div');
        actions.className = 'vw-org-builder__role-actions';
        actions.append(
            makeActionBtn({
                icon: 'bi-microsoft',
                label: 'Aus M365',
                title: 'Person aus Microsoft 365 dieser Rolle zuordnen',
                variant: 'primary',
                onClick: function () {
                    void onAddPersonToRole(node.roleCode);
                }
            }),
            makeActionBtn({
                icon: 'bi-pencil',
                label: 'Bearbeiten',
                title: 'Rolle bearbeiten (Pop-up)',
                variant: 'ghost',
                onClick: function () {
                    if (typeof d.openRole === 'function') d.openRole(node.roleCode);
                }
            }),
            makeActionBtn({
                icon: 'bi-trash',
                label: 'Entfernen',
                title: 'Rolle nur aus dem Organigramm entfernen (Stammdaten bleiben)',
                variant: 'danger',
                onClick: function () {
                    onDeleteNode(node.id);
                }
            })
        );
        card.appendChild(actions);

        // Filter
        const blob =
            (node.label || '') +
            ' ' +
            (node.roleCode || '') +
            ' ' +
            people
                .map(function (p) {
                    return (p.name || '') + ' ' + (p.email || '');
                })
                .join(' ');
        if (!passesFilter(blob)) card.hidden = true;
        return card;
    }

    function renderGroupSection(node, layout, depth) {
        const sec = document.createElement('section');
        sec.className =
            'vw-org-builder__group vw-org-builder__group--' +
            (node.tone || (node.type === ORG_NODE_SAMMEL ? 'sammel' : 'custom'));
        sec.dataset.nodeId = node.id;
        if (depth > 0) sec.classList.add('vw-org-builder__group--nested');

        const head = document.createElement('div');
        head.className = 'vw-org-builder__group-head';
        const titles = document.createElement('div');
        titles.className = 'vw-org-builder__group-titles';
        const icon =
            node.type === ORG_NODE_SAMMEL
                ? 'bi-people-fill'
                : node.tone === 'lead'
                  ? 'bi-person-badge'
                  : 'bi-collection';
        titles.innerHTML =
            '<h3><i class="bi ' +
            icon +
            '" aria-hidden="true"></i> ' +
            escapeHtml(node.label) +
            '</h3>';
        if (node.type === ORG_NODE_SAMMEL && (node.m365GroupMail || node.m365GroupLabel)) {
            const meta = document.createElement('p');
            meta.className = 'muted vw-org-builder__group-meta';
            meta.textContent = node.m365GroupLabel || node.m365GroupMail;
            titles.appendChild(meta);
        }
        titles.title = 'Doppelklick: Bereich bearbeiten';
        titles.addEventListener('dblclick', function (ev) {
            ev.preventDefault();
            openGroupEditor(node);
        });
        head.appendChild(titles);
        const headActions = document.createElement('div');
        headActions.className = 'vw-org-builder__group-actions';
        headActions.append(
            makeActionBtn({
                icon: 'bi-pencil',
                label: 'Bearbeiten',
                title: 'Bereich bearbeiten (Pop-up)',
                variant: 'ghost',
                onClick: function () {
                    openGroupEditor(node);
                }
            }),
            makeActionBtn({
                icon: 'bi-person-vcard',
                label: 'Rolle',
                title: 'Neue Rolle in diesem Bereich anlegen',
                variant: 'success',
                onClick: function () {
                    void onAddRole(node.id);
                }
            }),
            makeActionBtn({
                icon: 'bi-folder-plus',
                label: 'Gruppe',
                title: 'Untergruppe anlegen (z. B. Fachleitungen)',
                variant: 'ghost',
                onClick: function () {
                    void onAddGroup(node.id);
                }
            }),
            makeActionBtn({
                icon: 'bi-microsoft',
                label: 'Sammelgruppe',
                title: 'Microsoft-365-Sammelgruppe hier einhängen',
                variant: 'primary',
                onClick: function () {
                    void onAddSammel(node.id);
                }
            })
        );
        if (node.parentId || node.tone === 'custom' || node.type === ORG_NODE_SAMMEL) {
            headActions.appendChild(
                makeActionBtn({
                    icon: 'bi-trash',
                    label: 'Entfernen',
                    title: 'Diesen Bereich aus dem Organigramm entfernen',
                    variant: 'danger',
                    onClick: function () {
                        onDeleteNode(node.id);
                    }
                })
            );
        }
        head.appendChild(headActions);
        sec.appendChild(head);

        const body = document.createElement('div');
        body.className = 'vw-org-builder__group-body';
        bindDropZone(body, function (payload) {
            if (payload.kind === 'node') onDropNode(payload.id, node.id);
        });

        const kids = childrenOf(layout, node.id);
        const roleKids = kids.filter(function (k) {
            return k.type === ORG_NODE_ROLE;
        });
        const groupKids = kids.filter(function (k) {
            return k.type === ORG_NODE_GROUP || k.type === ORG_NODE_SAMMEL;
        });

        if (roleKids.length) {
            const grid = document.createElement('div');
            grid.className = 'vw-org-builder__roles';
            roleKids.forEach(function (rk) {
                grid.appendChild(renderRoleCard(rk, layout));
            });
            body.appendChild(grid);
        }
        groupKids.forEach(function (gk) {
            body.appendChild(renderGroupSection(gk, layout, depth + 1));
        });
        if (!kids.length) {
            const empty = document.createElement('p');
            empty.className = 'muted vw-org-builder__empty-hint';
            empty.textContent = 'Leer – Rolle/Gruppe hierher ziehen oder oben hinzufügen.';
            body.appendChild(empty);
        }
        sec.appendChild(body);
        return sec;
    }

    function renderToolbar(host) {
        const bar = document.createElement('div');
        bar.className = 'vw-org-builder__toolbar';
        const hint = document.createElement('p');
        hint.className = 'vw-org-builder__toolbar-hint muted';
        hint.textContent =
            'Türkis = 1 Person · Violett = Team · Grau = unbesetzt · Doppelklick oder Stift = Details im Pop-up.';
        bar.appendChild(hint);
        const actions = document.createElement('div');
        actions.className = 'vw-org-builder__toolbar-actions';
        actions.append(
            makeActionBtn({
                icon: 'bi-folder-plus',
                label: 'Gruppe',
                title: 'Neuen Bereich oben anlegen',
                variant: 'ghost',
                onClick: function () {
                    void onAddGroup(null);
                }
            }),
            makeActionBtn({
                icon: 'bi-person-vcard',
                label: 'Rolle',
                title: 'Neue Rolle oben anlegen',
                variant: 'success',
                onClick: function () {
                    void onAddRole(null);
                }
            }),
            makeActionBtn({
                icon: 'bi-microsoft',
                label: 'Sammelgruppe',
                title: 'Microsoft-365-Sammelgruppe einhängen',
                variant: 'primary',
                onClick: function () {
                    void onAddSammel(null);
                }
            }),
            makeActionBtn({
                icon: 'bi-arrow-repeat',
                label: 'Neu aufbauen',
                title: 'Organigramm aus aktuellen Stammdaten-Rollen neu erzeugen',
                variant: 'ghost',
                onClick: function () {
                    Promise.resolve(
                        typeof d.confirm === 'function'
                            ? d.confirm(
                                  'Organigramm aus aktuellen Rollen neu aufbauen? Eigene Gruppen-Struktur geht verloren.',
                                  {
                                      title: 'Neu aufbauen',
                                      okText: 'Neu aufbauen',
                                      cancelText: 'Abbrechen',
                                      danger: true
                                  }
                              )
                            : true
                    ).then(function (ok) {
                        if (!ok) return;
                        const roles = typeof d.getRoles === 'function' ? d.getRoles() : [];
                        saveLayout(seedOrganigrammFromRoles(roles));
                        toast('Organigramm neu aufgebaut.');
                        render();
                    });
                }
            })
        );
        bar.appendChild(actions);
        host.appendChild(bar);
    }

    let hostEl = null;

    function render() {
        if (!hostEl) return;
        const layout = getLayout();
        // Persist sync'd layout once so new roles appear next time
        if (typeof d.setLayout === 'function') {
            const prev = typeof d.getLayout === 'function' ? d.getLayout() : null;
            if (!prev || !prev.nodes || !prev.nodes.length) saveLayout(layout);
            else saveLayout(layout);
        }
        hostEl.replaceChildren();
        renderToolbar(hostEl);
        const canvas = document.createElement('div');
        canvas.className = 'vw-org-builder__canvas';
        canvas.setAttribute('aria-label', 'Organigramm-Arbeitsfläche');
        bindDropZone(canvas, function (payload) {
            if (payload.kind === 'node') onDropNode(payload.id, null);
        });
        const roots = childrenOf(layout, null);
        if (!roots.length) {
            const p = document.createElement('p');
            p.className = 'muted';
            p.textContent = 'Noch leer – „Gruppe“ oder „Aus Rollen neu aufbauen“ wählen.';
            canvas.appendChild(p);
        } else {
            roots.forEach(function (root) {
                if (root.type === ORG_NODE_ROLE) {
                    const wrap = document.createElement('div');
                    wrap.className = 'vw-org-builder__roles';
                    wrap.appendChild(renderRoleCard(root, layout));
                    canvas.appendChild(wrap);
                } else {
                    canvas.appendChild(renderGroupSection(root, layout, 0));
                }
            });
        }
        hostEl.appendChild(canvas);
        if (typeof d.onMeta === 'function') {
            const roleNodes = layout.nodes.filter(function (n) {
                return n.type === ORG_NODE_ROLE;
            });
            let people = 0;
            roleNodes.forEach(function (n) {
                people += peopleForRoleCode(n.roleCode).length;
            });
            d.onMeta(
                roleNodes.length +
                    ' Rollen · ' +
                    people +
                    ' Personen · Personen ziehen = Zuordnung ändern · „Entfernen“ = nur aus Organigramm'
            );
        }
    }

    return {
        mount: function (host) {
            hostEl = host;
            render();
        },
        refresh: render,
        getLayout: getLayout
    };
}

if (typeof window !== 'undefined') {
    window.ms365VerwaltungOrganigrammBuilder = { createOrganigrammBuilder: createOrganigrammBuilder };
}

export default { createOrganigrammBuilder };
