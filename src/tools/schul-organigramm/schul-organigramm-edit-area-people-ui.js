/**
 * Area editor people tab + save wiring.
 */
import {
    normalizeSchulOrganigramm,
    setSchulOrganigrammAreaM365,
    SCHUL_ORG_M365_OBJECT_KINDS,
    normalizeSchulOrgM365,
    normalizeSchulOrgM365Binding,
    emptySchulOrgM365Bindings,
    schulOrgM365ObjectKindLabel,
    audienceIdForArea,
    listSchulOrgClassFunctions,
    suggestSchulOrgMailNickname
} from './schul-organigramm-logic.js';
import { listPeopleForSchulOrgNode } from './schul-organigramm-edit-people.js';
import { addSchulOrgPersonToArea, removeSchulOrgPersonFromArea } from './schul-organigramm-edit-ops.js';
import { writeSchoolData } from '../../shared/stammdaten-write.js';
import {
    normalizeVerwaltungAudienceGroups
} from '../../shared/administration-audience-groups.js';
import { patchAudienceGroupGraphId } from '../../shared/tenant-admin-audience-m365.js';

/**
 * @param {Record<string, any>} ae
 */
export function wireAreaEditorPeople(ae) {

    const overlay = ae.overlay;
    const area = ae.area;
    const settings = ae.settings;
    const sammelKey = ae.sammelKey;
    const isKlassenchatArea = ae.isKlassenchatArea;
    const isArgeArea = ae.isArgeArea;
    const isClassArea = ae.isClassArea;
    const loadSettings = ae.loadSettings;
    const saveSettings = ae.saveSettings;
    const toast = ae.toast;
    const confirmDlg = ae.confirmDlg;
    const notify = ae.notify;
    const closeOverlay = ae.closeOverlay;
    const escapeHtml = ae.escapeHtml;
    const normStr = ae.normStr;
    const normEmail = ae.normEmail;
    const normCode = ae.normCode;
    const toggleInlineEntraUserPicker = ae.toggleInlineEntraUserPicker;
    const toggleInlineEntraGroupPicker = ae.toggleInlineEntraGroupPicker;
    const stripEphemeralSchoolSettings = ae.stripEphemeralSchoolSettings;
    const plainJsonClone = ae.plainJsonClone;
    const settingsForPersist = ae.settingsForPersist;
    const suggestLocalPersonEmail = ae.suggestLocalPersonEmail;
    const persistOrganigrammLayout = ae.persistOrganigrammLayout;
    const removeArea = ae.removeArea;
    const inlineEupHost = ae.inlineEupHost;
    const inlineEgpHost = ae.inlineEgpHost;
    const inlineMailHost = ae.inlineMailHost;
    const inlineOwnerHost = ae.inlineOwnerHost;
    const inlineEupState = ae.inlineEupState;
    const inlineEgpState = ae.inlineEgpState;
    const inlineMailState = ae.inlineMailState;
    const inlineOwnerState = ae.inlineOwnerState;
    const inlinePickers = ae.inlinePickers;
    const groupModeEl = ae.groupModeEl;
    const groupKindEl = ae.groupKindEl;
    const mailboxModeEl = ae.mailboxModeEl;
    const createNameEl = ae.createNameEl;
    const createNickEl = ae.createNickEl;
    const mailCreateNameEl = ae.mailCreateNameEl;
    const mailCreateAliasEl = ae.mailCreateAliasEl;
    const mailCreateSmtpEl = ae.mailCreateSmtpEl;
    const mailScriptEl = ae.mailScriptEl;
    const mailScriptCopyBtn = ae.mailScriptCopyBtn;
    const mailScriptDlBtn = ae.mailScriptDlBtn;
    const mailDelegatesEl = ae.mailDelegatesEl;
    const groupPeopleEl = ae.groupPeopleEl;
    const groupOwnerEl = ae.groupOwnerEl;


        function renderPeopleList() {
            const list = overlay.querySelector('[data-area-people-list]');
            if (!list) return;
            const cur = loadSettings();
            const people = listPeopleForSchulOrgNode(cur, area);
            if (!people.length) {
                list.innerHTML =
                    '<li class="muted">' +
                    (sammelKey === 'lehrer'
                        ? 'Noch keine Lehrkräfte in den Stammdaten.'
                        : sammelKey === 'schueler'
                          ? 'Keine direkten Personen – Klassen liegen in den Jahrgangs-Bereichen.'
                          : isKlassenchatArea
                            ? 'Noch keine Lehrkräfte aus Belegung/KV – Unterrichtsbelegung prüfen oder Personen manuell hinzufügen.'
                            : 'Noch keine Personen – „Person“ oder Microsoft 365 nutzen.') +
                    '</li>';
                return;
            }
            list.innerHTML = people
                .slice(0, sammelKey ? 80 : people.length)
                .map(function (p) {
                    const em = normEmail(p.email);
                    const localMail = /@local\.schul-org$/i.test(em);
                    const seat = normStr(p.seatLabel);
                    const canRemove =
                        (!!em && sammelKey === 'lehrer') ||
                        (!sammelKey && p.movable !== false && !!em);
                    return (
                        '<li class="schul-org-area-people__item">' +
                        '<span>' +
                        (seat
                            ? '<span class="schul-org-area-people__seat">' +
                              escapeHtml(seat) +
                              '</span><br/>'
                            : '') +
                        '<strong>' +
                        escapeHtml(p.name || p.email) +
                        '</strong>' +
                        (em && !localMail
                            ? '<br/><span class="muted">' + escapeHtml(em) + '</span>'
                            : localMail
                              ? '<br/><span class="muted">ohne E-Mail</span>'
                              : '') +
                        '</span>' +
                        (canRemove
                            ? '<button type="button" class="btn btn-sm alt" data-area-remove-email="' +
                              escapeHtml(p.email) +
                              '">Entfernen</button>'
                            : '') +
                        '</li>'
                    );
                })
                .join('');
            if (sammelKey && people.length > 80) {
                list.innerHTML +=
                    '<li class="muted">… und ' + (people.length - 80) + ' weitere in den Stammdaten</li>';
            }
            list.querySelectorAll('[data-area-remove-email]').forEach(function (btn) {
                btn.addEventListener('click', function () {
                    const em = btn.getAttribute('data-area-remove-email') || '';
                    const st = loadSettings();
                    const res = removeSchulOrgPersonFromArea(st, area, em);
                    if (!res.ok) {
                        toast(res.message || 'Entfernen fehlgeschlagen.', 'warning');
                        return;
                    }
                    saveSettings(res.settings);
                    renderPeopleList();
                    ae.refreshGroupPeopleUi();
                    ae.refreshMailDelegatesHint();
                    notify();
                });
            });
        }

        try {
            ae.syncM365ModeUi();
            renderPeopleList();
            ae.refreshGroupPeopleUi();
            ae.refreshMailDelegatesHint();
        } catch (initErr) {
            console.warn('[schul-org] area editor init', initErr);
        }

        overlay.querySelectorAll('[data-schul-org-close]').forEach(function (btn) {
            btn.addEventListener('click', closeOverlay);
        });
        overlay.addEventListener('click', function (ev) {
            if (ev.target === overlay) closeOverlay();
        });

        // Speichern früh verdrahten – spätere Init-Fehler dürfen den Button nicht totlegen
        const saveAreaBtn = overlay.querySelector('[data-schul-org-save]');
        if (saveAreaBtn) {
            saveAreaBtn.addEventListener('click', function () {
                /** @type {string[]} */
                let skippedCreate = [];
                /** @type {string} */
                let linkedGid = '';
                try {
                    const cur = stripEphemeralSchoolSettings(loadSettings());
                    const labelEl = /** @type {HTMLInputElement} */ (
                        overlay.querySelector('#schulOrgAreaLabel')
                    );
                    const nextLabel = normStr(labelEl && labelEl.value) || area.label;
                    let nextLayout = normalizeSchulOrganigramm(cur.schulOrganigramm);
                    const node = nextLayout.nodes.find(function (n) {
                        return n.id === area.id;
                    });
                    if (node) node.label = nextLabel;
                    try {
                        ae.readBindingsFromCards({ applyUi: false });
                    } catch (bindErr) {
                        console.warn('[schul-org] readBindingsFromCards', bindErr);
                    }
                    // Unfertige „Neu anlegen“-Slots nicht blockieren – zurück auf „nicht nutzen“
                    skippedCreate = [];
                    const areaBindings =
                        area.m365 && area.m365.bindings && typeof area.m365.bindings === 'object'
                            ? area.m365.bindings
                            : {};
                    SCHUL_ORG_M365_OBJECT_KINDS.forEach(function (kind) {
                        if (kind === 'groupChat') return;
                        const b = areaBindings[kind];
                        if (!b || b.mode !== 'create') return;
                        if (kind === 'mailbox') {
                            if (!b.upn && !b.id) {
                                b.mode = 'none';
                                skippedCreate.push(schulOrgM365ObjectKindLabel(kind));
                            }
                            return;
                        }
                        if (!b.id) {
                            b.mode = 'none';
                            skippedCreate.push(schulOrgM365ObjectKindLabel(kind));
                        }
                    });
                    SCHUL_ORG_M365_OBJECT_KINDS.forEach(function (kind) {
                        const b = areaBindings[kind];
                        if (b && b.mode === 'create' && b.id) b.mode = 'match';
                    });
                    const cleanBindings = emptySchulOrgM365Bindings();
                    SCHUL_ORG_M365_OBJECT_KINDS.forEach(function (kind) {
                        const b = areaBindings[kind] || {};
                        cleanBindings[kind] = normalizeSchulOrgM365Binding(
                            {
                                mode: b.mode,
                                id: b.id,
                                label: b.label,
                                mail: b.mail,
                                upn: b.upn,
                                webUrl: b.webUrl
                            },
                            kind
                        );
                    });
                    const nextM365 = normalizeSchulOrgM365({ bindings: cleanBindings });
                    if (nextM365) area.m365 = nextM365;
                    nextLayout = setSchulOrganigrammAreaM365(nextLayout, area.id, {
                        bindings: cleanBindings
                    });
                    const nodeAfter = nextLayout.nodes.find(function (n) {
                        return n.id === area.id;
                    });
                    if (nodeAfter) nodeAfter.label = nextLabel;
                    area.label = nextLabel;
                    const aid = audienceIdForArea(area);
                    if (aid && Array.isArray(cur.verwaltungAudienceGroups)) {
                        cur.verwaltungAudienceGroups = normalizeVerwaltungAudienceGroups(
                            cur.verwaltungAudienceGroups
                        ).map(function (g) {
                            if (g.id === aid) return Object.assign({}, g, { label: nextLabel });
                            return g;
                        });
                    }
                    if (aid) {
                        try {
                            if (area.m365 && area.m365.groupMode === 'match' && normStr(area.m365.groupId)) {
                                patchAudienceGroupGraphId(aid, {
                                    id: area.m365.groupId,
                                    displayName: area.m365.groupLabel || nextLabel,
                                    mailNickname: '',
                                    mode: 'matched'
                                });
                            } else if (area.m365 && area.m365.groupMode === 'none') {
                                patchAudienceGroupGraphId(aid, null);
                            }
                        } catch {
                            /* Setup-Spiegel optional */
                        }
                    }
                    if (isArgeArea && area.refKey) {
                        const arges = Array.isArray(cur.arges) ? cur.arges.slice() : [];
                        const ai = arges.findIndex(function (a) {
                            return (
                                normStr(a && a.code).toUpperCase() ===
                                normStr(area.refKey).toUpperCase()
                            );
                        });
                        if (ai >= 0) {
                            arges[ai] = Object.assign({}, arges[ai], { name: nextLabel });
                            cur.arges = arges;
                        }
                    } else if (isClassArea && area.refKey) {
                        const classes = Array.isArray(cur.classes) ? cur.classes.slice() : [];
                        const ci = classes.findIndex(function (c) {
                            return (
                                normStr(c && c.code).toUpperCase() ===
                                normStr(area.refKey).toUpperCase()
                            );
                        });
                        if (ci >= 0) {
                            classes[ci] = Object.assign({}, classes[ci], { name: nextLabel });
                            cur.classes = classes;
                        }
                    }
                    // Layout zuerst JSON-sicher machen (tiefe Nesting/Zyklen → kein Stack-Overflow)
                    const cleanLayout = plainJsonClone(
                        normalizeSchulOrganigramm(nextLayout),
                        normalizeSchulOrganigramm(nextLayout)
                    );
                    cur.schulOrganigramm = cleanLayout;
                    const listsTouched =
                        !!(aid && nextLabel) ||
                        !!(isArgeArea && area.refKey) ||
                        !!(isClassArea && area.refKey);
                    if (listsTouched) {
                        // Listen + Landkarte gemeinsam, aber nur bereinigtes Settings-Objekt
                        const cleanCur = settingsForPersist(cur);
                        cleanCur.schulOrganigramm = cleanLayout;
                        try {
                            window.__ms365SchulOrgSkipSyncOnce = true;
                        } catch {
                            /* ignore */
                        }
                        if (typeof writeSchoolData === 'function') {
                            writeSchoolData(cleanCur, {
                                reason: 'schul-org-area-save',
                                skipMaterialize: true
                            });
                        } else {
                            saveSettings(cleanCur);
                        }
                    } else {
                        // Nur Landkarte (M365/Label am Knoten) – ohne vollen Stammdaten-Save
                        persistOrganigrammLayout(cleanLayout, 'schul-org-area-save');
                    }
                    linkedGid =
                        area.m365 &&
                        (area.m365.groupMode === 'match' || area.m365.groupMode === 'create')
                            ? normStr(area.m365.groupId)
                            : '';
                    closeOverlay();
                } catch (saveErr) {
                    console.error('[schul-org] area save', saveErr);
                    toast(
                        'Speichern fehlgeschlagen: ' +
                            (saveErr && saveErr.message ? saveErr.message : String(saveErr)),
                        'error'
                    );
                    return;
                }
                try {
                    notify();
                } catch (notifyErr) {
                    console.warn('[schul-org] area save notify', notifyErr);
                }
                toast(
                    skippedCreate.length
                        ? 'Bereich gespeichert. Offene Neuanlagen zurückgesetzt: ' +
                              skippedCreate.join(', ')
                        : 'Bereich gespeichert.',
                    'success'
                );
                if (isArgeArea && linkedGid) {
                    void ae.syncArgeLeadFromLinkedGroup(linkedGid);
                }
                if (isClassArea && linkedGid) {
                    void ae.syncClassKvFromLinkedGroup(linkedGid);
                }
            });
        }

        const manualForm = /** @type {HTMLElement|null} */ (
            overlay.querySelector('[data-area-manual-form]')
        );
        const manualNameEl = /** @type {HTMLInputElement|null} */ (
            overlay.querySelector('#schulOrgAreaPersonName')
        );
        const manualEmailEl = /** @type {HTMLInputElement|null} */ (
            overlay.querySelector('#schulOrgAreaPersonEmail')
        );
        const roleHeadEl = /** @type {HTMLSelectElement|null} */ (
            overlay.querySelector('[data-area-person-role]')
        );
        const roleManualEl = /** @type {HTMLSelectElement|null} */ (
            overlay.querySelector('[data-area-person-role-manual]')
        );

        function selectedAreaPersonRole() {
            if (roleManualEl && manualForm && !manualForm.hidden) {
                return normStr(roleManualEl.value) || '';
            }
            if (roleHeadEl) return normStr(roleHeadEl.value) || '';
            if (roleManualEl) return normStr(roleManualEl.value) || '';
            if (isClassArea) return 'student';
            if (isArgeArea) return 'argeMember';
            return '';
        }

        function syncAreaPersonRoleSelects(from) {
            const val = normStr(from && from.value);
            if (!val) return;
            if (roleHeadEl && from !== roleHeadEl) roleHeadEl.value = val;
            if (roleManualEl && from !== roleManualEl) roleManualEl.value = val;
        }

        if (roleHeadEl) {
            roleHeadEl.addEventListener('change', function () {
                syncAreaPersonRoleSelects(roleHeadEl);
            });
        }
        if (roleManualEl) {
            roleManualEl.addEventListener('change', function () {
                syncAreaPersonRoleSelects(roleManualEl);
            });
        }

        function submitManualPerson() {
            if (!manualNameEl || !manualEmailEl) return;
            const name = normStr(manualNameEl.value);
            let email = normEmail(manualEmailEl.value);
            if (!name && !email) {
                toast('Bitte mindestens einen Namen eingeben.', 'warning');
                manualNameEl.focus();
                return;
            }
            if (!email) {
                email = suggestLocalPersonEmail(name);
                manualEmailEl.value = email;
            }
            const st = loadSettings();
            const res = addSchulOrgPersonToArea(st, area, {
                name: name || email,
                email: email,
                role: selectedAreaPersonRole()
            });
            if (!res.ok) {
                toast(res.message || 'Hinzufügen fehlgeschlagen.', 'warning');
                return;
            }
            saveSettings(res.settings);
            manualNameEl.value = '';
            manualEmailEl.value = '';
            if (manualForm) manualForm.hidden = true;
            renderPeopleList();
            ae.refreshGroupPeopleUi();
            ae.refreshMailDelegatesHint();
            notify();
            toast(res.message || 'Person hinzugefügt.', 'success');
        }

        const addManualBtn = overlay.querySelector('[data-area-add-person-manual]');
        if (addManualBtn && manualForm) {
            addManualBtn.addEventListener('click', function () {
                if (inlineEupState.api) {
                    inlineEupState.api.destroy();
                    inlineEupState.api = null;
                }
                if (inlineEupHost) inlineEupHost.hidden = true;
                manualForm.hidden = !manualForm.hidden;
                if (!manualForm.hidden) {
                    syncAreaPersonRoleSelects(roleHeadEl || roleManualEl);
                    if (manualNameEl) manualNameEl.focus();
                }
            });
        }
        const manualSubmitBtn = overlay.querySelector('[data-area-manual-submit]');
        if (manualSubmitBtn) {
            manualSubmitBtn.addEventListener('click', submitManualPerson);
        }
        if (manualEmailEl) {
            manualEmailEl.addEventListener('keydown', function (ev) {
                if (ev.key === 'Enter') {
                    ev.preventDefault();
                    submitManualPerson();
                }
            });
        }
        if (manualNameEl) {
            manualNameEl.addEventListener('keydown', function (ev) {
                if (ev.key === 'Enter') {
                    ev.preventDefault();
                    if (manualEmailEl) manualEmailEl.focus();
                }
            });
        }

        const addPersonBtn = overlay.querySelector('[data-area-add-person]');
        if (addPersonBtn) {
            addPersonBtn.addEventListener('click', function () {
                if (manualForm) manualForm.hidden = true;
                const role = selectedAreaPersonRole();
                let roleHint = '';
                if (role === 'kv') roleHint = ' als Klassenvorstand';
                else if (role === 'student') roleHint = ' als Schüler:in';
                else if (role === 'argeHead') roleHint = ' als ARGE-Leitung';
                else if (role === 'argeMember') roleHint = ' als Mitglied';
                else if (role && String(role).indexOf('fn:') === 0) {
                    const fnId = String(role).slice(3);
                    const def = listSchulOrgClassFunctions().find(function (d) {
                        return d.id === fnId;
                    });
                    roleHint = ' als ' + ((def && def.label) || 'Klassenfunktion');
                }
                void toggleInlineEntraUserPicker(
                    inlineEupHost,
                    inlineEupState,
                    {
                        hint:
                            'Person für „' +
                            area.label +
                            '“' +
                            roleHint +
                            ' in Microsoft 365 suchen und antippen.',
                        onPick: function (picked) {
                            const st = loadSettings();
                            const res = addSchulOrgPersonToArea(st, area, {
                                email: picked.mail || picked.userPrincipalName || '',
                                name: picked.displayName || '',
                                role: selectedAreaPersonRole()
                            });
                            if (!res.ok) {
                                toast(res.message || 'Hinzufügen fehlgeschlagen.', 'warning');
                                return;
                            }
                            saveSettings(res.settings);
                            renderPeopleList();
                            ae.refreshGroupPeopleUi();
                            ae.refreshMailDelegatesHint();
                            notify();
                            toast(res.message || 'Person hinzugefügt.', 'success');
                        }
                    },
                    inlinePickers
                ).catch(function (e) {
                    toast(
                        'Entra-Auswahl fehlgeschlagen: ' + (e && e.message ? e.message : e),
                        'warning'
                    );
                });
            });
        }

    ae.renderPeopleList = renderPeopleList;
    ae.selectedAreaPersonRole = selectedAreaPersonRole;
    ae.syncAreaPersonRoleSelects = syncAreaPersonRoleSelects;
    ae.submitManualPerson = submitManualPerson;
}
