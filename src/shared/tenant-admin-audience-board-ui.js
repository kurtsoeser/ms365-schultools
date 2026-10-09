/**
 * Stammdaten Tab Verwaltung: Zielgruppen als Überschriften mit Mitgliedern (Mehrfach-Zuordnung pro E-Mail).
 */
import {
    addCustomVerwaltungAudienceGroup,
    audienceGroupIdsForEmail,
    BUILTIN_AUDIENCE_SCHULLEITUNG,
    BUILTIN_AUDIENCE_VERWALTUNG,
    ensureAdminAudienceOnSettings,
    findAdminDisplayRowByEmail,
    normalizeAdminAudienceMemberships,
    normalizeVerwaltungAudienceGroups,
    removeCustomVerwaltungAudienceGroup,
    setAudienceGroupsForEmail
} from './administration-audience-groups.js';
import { buildAudienceGroupM365Panel } from './tenant-admin-audience-m365.js';
import {
    appendAdminDirectoryActionButtons,
    paintDirectoryMatchCell
} from './tenant-admin-directory-actions-ui.js';
import { mountAudienceMemberM365Search } from './tenant-admin-audience-member-search.js';

/**
 * @param {{
 *   host: HTMLElement,
 *   getSettings: () => object,
 *   onChange: (patch: { verwaltungAudienceGroups: object[], adminAudienceMemberships: object[] }) => void,
 *   scheduleAutoSave: () => void,
 *   normStr: (v: unknown) => string,
 *   getSchoolName?: () => string,
 *   setSummary?: (msg: string, kind?: string) => void,
 *   onVerifyDirectory?: (email: string) => Promise<void>,
 *   onCreateEntraUser?: (email: string, personName: string) => Promise<void>,
 *   onDirectoryChanged?: () => void,
 *   getAdminDisplayRows?: () => Array<{ code?: string, name?: string, personName?: string, email?: string }>,
 *   onAdminDisplayFieldEdit?: (
 *     email: string,
 *     field: 'name' | 'code',
 *     value: string,
 *     ctx: { audienceGroupId?: string, personName?: string },
 *     meta?: { cancelled?: boolean }
 *   ) => void,
 * }} opts
 */
function audienceGroupIcon(groupId) {
    const id = String(groupId || '');
    if (id === BUILTIN_AUDIENCE_SCHULLEITUNG) return 'bi-person-badge';
    if (id === BUILTIN_AUDIENCE_VERWALTUNG) return 'bi-building';
    return 'bi-people';
}

function memberGridCell(className, text, title) {
    const el = document.createElement('span');
    el.className = 'ts-audience-board__member-cell ' + (className || '');
    const t = text != null ? String(text).trim() : '';
    el.textContent = t || '–';
    if (!t) el.classList.add('muted');
    if (title) el.title = title;
    else if (t) el.title = t;
    return el;
}

function bindEditableMemberCell(cell, initialValue, label, normStr, onCommit) {
    cell.classList.add('ts-audience-board__cell-editable');
    const hint = (label || 'Feld') + ' – Doppelklick zum Bearbeiten';
    cell.title = hint;
    cell.addEventListener('dblclick', function (ev) {
        ev.stopPropagation();
        if (cell.querySelector('input.cell-editor')) return;
        const prev = String(initialValue ?? '').trim();
        const input = document.createElement('input');
        input.className = 'cell-editor';
        input.type = 'text';
        input.value = prev;
        cell.replaceChildren(input);
        input.focus();
        input.select();
        let finished = false;
        const finish = function (cancelled) {
            if (finished) return;
            finished = true;
            onCommit(cancelled ? prev : normStr(input.value), { cancelled: !!cancelled });
        };
        input.addEventListener('keydown', function (e) {
            if (e.key === 'Enter') {
                e.preventDefault();
                finish(false);
            } else if (e.key === 'Escape') {
                e.preventDefault();
                finish(true);
            }
        });
        input.addEventListener('blur', function () {
            finish(false);
        });
    });
}

export function mountTenantAdminAudienceBoard(opts) {
    const host = opts.host;
    if (!host) return { refresh: function () {} };

    function stateFromSettings() {
        return ensureAdminAudienceOnSettings(opts.getSettings() || {});
    }

    function emit(groups, memberships) {
        opts.onChange({
            verwaltungAudienceGroups: normalizeVerwaltungAudienceGroups(groups),
            adminAudienceMemberships: normalizeAdminAudienceMemberships(memberships)
        });
        opts.scheduleAutoSave();
        render();
    }

    function render() {
        const st = stateFromSettings();
        const groups = st.groups;
        const memberships = st.memberships;
        host.replaceChildren();

        const head = document.createElement('div');
        head.className = 'ts-audience-board__head';
        head.innerHTML =
            '<h3 class="ts-audience-board__title">Microsoft&nbsp;365-Zielgruppen</h3>' +
            '<p class="muted ts-audience-board__lead">Personen können in <strong>mehreren</strong> Gruppen stehen. ' +
            'Pro Gruppe: Personen zuordnen <em>und</em> die Microsoft‑365‑Gruppe prüfen, anlegen oder Mitglieder synchronisieren.</p>';
        host.appendChild(head);

        const actions = document.createElement('div');
        actions.className = 'ts-audience-board__actions';
        const btnAdd = document.createElement('button');
        btnAdd.type = 'button';
        btnAdd.className = 'btn btn-success btn-sm';
        btnAdd.innerHTML = '<i class="bi bi-plus-circle"></i>Verwaltungsgruppe';
        btnAdd.addEventListener('click', function () {
            const label = window.prompt('Name der neuen Verwaltungsgruppe (z. B. Schupersonal):', '');
            if (label === null) return;
            const added = addCustomVerwaltungAudienceGroup(label, groups);
            emit(added.groups, memberships);
        });
        actions.appendChild(btnAdd);
        host.appendChild(actions);

        groups.forEach(function (group) {
            const section = document.createElement('section');
            section.className = 'ts-audience-board__group';
            section.dataset.audienceGroupId = group.id;

            const banner = document.createElement('header');
            banner.className = 'ts-audience-board__group-banner';
            const icon = document.createElement('i');
            icon.className = 'bi ' + audienceGroupIcon(group.id);
            icon.setAttribute('aria-hidden', 'true');
            const heading = document.createElement('h3');
            heading.className = 'ts-audience-board__group-heading';
            heading.textContent = group.label;
            banner.append(icon, heading);
            if (group.builtin) {
                const badge = document.createElement('span');
                badge.className = 'ts-audience-board__badge ts-audience-board__badge--banner';
                badge.textContent = 'Standard';
                banner.appendChild(badge);
            }
            section.appendChild(banner);

            const m365Row = document.createElement('div');
            m365Row.className = 'ts-audience-board__group-m365-row';

            const m365 = buildAudienceGroupM365Panel({
                group: group,
                memberships: memberships,
                schoolName: (typeof opts.getSchoolName === 'function' ? opts.getSchoolName() : '') || '',
                setSummary: opts.setSummary,
                onGroupMetaChange: function (groupId, meta) {
                    const next = groups.map(function (g) {
                        if (g.id !== groupId) return g;
                        return Object.assign({}, g, {
                            mailNick: meta.mailNick != null ? String(meta.mailNick).trim() : g.mailNick,
                            newDisplayName:
                                meta.newDisplayName != null ? String(meta.newDisplayName).trim() : g.newDisplayName
                        });
                    });
                    emit(next, memberships);
                },
                onMatchChanged: function () {
                    render();
                }
            });
            m365Row.appendChild(m365.inline);

            if (!group.builtin) {
                const btnRm = document.createElement('button');
                btnRm.type = 'button';
                btnRm.className = 'ts-audience-board__dir-btn ts-audience-board__dir-btn--danger';
                btnRm.title = 'Gruppe löschen (Mitgliedschaften entfallen)';
                btnRm.setAttribute('aria-label', btnRm.title);
                btnRm.innerHTML = '<i class="bi bi-trash" aria-hidden="true"></i>';
                btnRm.addEventListener('click', function () {
                    if (!window.confirm('Gruppe „' + group.label + '“ und alle Zuordnungen löschen?')) return;
                    const next = removeCustomVerwaltungAudienceGroup(group.id, groups, memberships);
                    emit(next.groups, next.memberships);
                });
                m365Row.appendChild(btnRm);
            }
            section.appendChild(m365Row);
            if (m365.aux && m365.aux.childNodes.length) section.appendChild(m365.aux);

            const body = document.createElement('div');
            body.className = 'ts-audience-board__group-body';

            const members = memberships.filter(function (m) {
                return m.groupId === group.id;
            });
            if (members.length) {
                const membersHead = document.createElement('div');
                membersHead.className = 'ts-audience-board__members-head';
                membersHead.innerHTML =
                    '<span>Bezeichnung</span><span>Kürzel</span><span>Person</span><span>E-Mail</span>' +
                    '<span>Microsoft&nbsp;365</span><span>Aktion</span>';
                body.appendChild(membersHead);
            }
            const list = document.createElement('ul');
            list.className = 'ts-audience-board__members';
            if (!members.length) {
                const li = document.createElement('li');
                li.className = 'muted ts-audience-board__members-empty';
                li.textContent = 'Noch keine Personen – unten E-Mail hinzufügen.';
                list.appendChild(li);
            } else {
                members.forEach(function (m) {
                    const li = document.createElement('li');
                    li.className = 'ts-audience-board__member-row';

                    const adminRows =
                        typeof opts.getAdminDisplayRows === 'function' ? opts.getAdminDisplayRows() : [];
                    const adminRow = findAdminDisplayRowByEmail(adminRows, m.email);
                    const roleLabel = adminRow ? opts.normStr(adminRow.name) : '';
                    const roleCode = adminRow ? opts.normStr(adminRow.code) : '';
                    const personLabel = adminRow
                        ? opts.normStr(adminRow.personName)
                        : opts.normStr(m.name);

                    const roleCell = memberGridCell('ts-audience-board__cell-role', roleLabel);
                    const codeCell = memberGridCell('ts-audience-board__cell-code', roleCode);
                    if (roleCode) {
                        codeCell.textContent = '';
                        codeCell.classList.remove('muted');
                        const codeEl = document.createElement('code');
                        codeEl.textContent = roleCode;
                        codeCell.appendChild(codeEl);
                    }
                    const editCtx = {
                        audienceGroupId: group.id,
                        personName: personLabel
                    };
                    if (typeof opts.onAdminDisplayFieldEdit === 'function') {
                        bindEditableMemberCell(roleCell, roleLabel, 'Bezeichnung', opts.normStr, function (val, meta) {
                            opts.onAdminDisplayFieldEdit(m.email, 'name', val, editCtx, meta);
                        });
                        bindEditableMemberCell(codeCell, roleCode, 'Kürzel', opts.normStr, function (val, meta) {
                            opts.onAdminDisplayFieldEdit(m.email, 'code', val, editCtx, meta);
                        });
                    }
                    li.appendChild(roleCell);
                    li.appendChild(codeCell);
                    li.appendChild(memberGridCell('ts-audience-board__cell-person', personLabel));
                    li.appendChild(memberGridCell('ts-audience-board__cell-email', m.email));

                    const ms = document.createElement('span');
                    paintDirectoryMatchCell(ms, m.email);
                    li.appendChild(ms);

                    const actionsWrap = document.createElement('span');
                    actionsWrap.className = 'ts-audience-board__member-actions-wrap';
                    li.appendChild(actionsWrap);

                    const personName = personLabel || opts.normStr(m.email).split('@')[0] || '';
                    appendAdminDirectoryActionButtons(actionsWrap, {
                        email: m.email,
                        personName: personName,
                        removeTitle: 'Aus dieser Zielgruppe entfernen',
                        onVerify:
                            typeof opts.onVerifyDirectory === 'function'
                                ? async function (email) {
                                      await opts.onVerifyDirectory(email);
                                      if (typeof opts.onDirectoryChanged === 'function') opts.onDirectoryChanged();
                                  }
                                : undefined,
                        onCreateUser:
                            personName && typeof opts.onCreateEntraUser === 'function'
                                ? async function (email, name) {
                                      await opts.onCreateEntraUser(email, name);
                                      if (typeof opts.onDirectoryChanged === 'function') opts.onDirectoryChanged();
                                  }
                                : undefined,
                        onRemove: function () {
                            const nextM = memberships.filter(function (x) {
                                return !(x.groupId === group.id && x.email === m.email);
                            });
                            emit(groups, nextM);
                        }
                    });
                    list.appendChild(li);
                });
            }
            body.appendChild(list);

            const addRow = document.createElement('div');
            addRow.className = 'ts-audience-board__add';

            function addMemberToGroup(emailRaw, nameRaw) {
                const em = opts.normStr(emailRaw).toLowerCase();
                if (!em || em.indexOf('@') === -1) {
                    window.alert('Bitte eine gültige E-Mail eingeben.');
                    return false;
                }
                const nm = opts.normStr(nameRaw);
                const ids = audienceGroupIdsForEmail(em, memberships, groups);
                if (ids.indexOf(group.id) === -1) ids.push(group.id);
                const nextM = setAudienceGroupsForEmail(em, ids, nm || undefined, memberships);
                emit(groups, nextM);
                return true;
            }

            mountAudienceMemberM365Search(addRow, {
                setSummary: opts.setSummary,
                onSelect: function (user) {
                    if (addMemberToGroup(user.email, user.displayName)) {
                        if (typeof opts.setSummary === 'function') {
                            opts.setSummary(
                                (user.displayName ? user.displayName + ' ' : '') + '(' + user.email + ') zur Gruppe hinzugefügt.',
                                'ok'
                            );
                        }
                    }
                }
            });

            const manualRow = document.createElement('div');
            manualRow.className = 'ts-audience-board__add-manual';
            const emailIn = document.createElement('input');
            emailIn.type = 'email';
            emailIn.placeholder = 'E-Mail (manuell)';
            emailIn.autocomplete = 'off';
            emailIn.className = 'ts-audience-board__input';
            const nameIn = document.createElement('input');
            nameIn.type = 'text';
            nameIn.placeholder = 'Name (optional)';
            nameIn.autocomplete = 'off';
            nameIn.className = 'ts-audience-board__input';
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn btn-sm';
            btn.textContent = 'Hinzufügen';
            btn.addEventListener('click', function () {
                if (addMemberToGroup(emailIn.value, nameIn.value)) {
                    emailIn.value = '';
                    nameIn.value = '';
                }
            });
            manualRow.append(emailIn, nameIn, btn);
            addRow.appendChild(manualRow);
            body.appendChild(addRow);
            section.appendChild(body);
            host.appendChild(section);
        });
    }

    render();
    return { refresh: render };
}
