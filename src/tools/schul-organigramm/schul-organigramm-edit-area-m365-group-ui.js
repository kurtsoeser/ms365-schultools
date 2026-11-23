/**
 * Area M365 group create/match listeners + graph sync.
 */
import {
    normalizeSchulOrgGroupKind,
    schulOrgGroupKindIsEntraGroup,
    suggestSchulOrgMailNickname,
    findSchulOrgNodesUsingM365Group,
    findSchulOrgNodesUsingM365Mailbox,
    inferSchulOrgGroupKind,
    normalizeSchulOrganigramm,
    isSgaBoardArea,
    SCHUL_ORG_BOARD_REF,
    schulOrgM365ConflictLabel,
    SCHUL_ORG_M365_OBJECT_KINDS,
    normalizeSchulOrgM365,
    normalizeSchulOrgM365Binding,
    emptySchulOrgM365Bindings,
    schulOrgM365ObjectKindLabel,
    schulOrgM365ObjectKindIcon,
    setSchulOrganigrammAreaM365,
    isArgeAreaNode,
    isClassAreaNode,
    listArgeCatalogGroupRefs
} from './schul-organigramm-logic.js';
import { applyArgeHeadFromGraphOwners, applyClassHeadFromGraphOwners } from './schul-organigramm-edit-people.js';
import { findClass, findArge } from './schul-organigramm-edit-find.js';
import { resolveClassGraphGroupId } from '../../shared/membership-hygiene.js';
import { getActiveOverlay } from './schul-organigramm-edit-overlay.js';

/**
 * @param {Record<string, any>} ae
 */
export function wireAreaEditorM365Group(ae) {

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

    const writeBinding = ae.writeBinding;
    const syncM365ModeUi = ae.syncM365ModeUi;
    const openArgeEditor = ae.openArgeEditor;
    const openClassEditor = ae.openClassEditor;
    const openSgaEditor = ae.openSgaEditor;
    const refreshM365Sums = ae.refreshM365Sums;
    const refreshGroupPeopleUi = ae.refreshGroupPeopleUi;
    const areaM365People = ae.areaM365People;

        if (groupModeEl) {
            groupModeEl.addEventListener('change', function () {
                area.m365.groupMode = groupModeEl.value || 'none';
                if (area.m365.groupMode === 'none') {
                    area.m365.groupId = null;
                    area.m365.groupLabel = null;
                    area.m365.groupMail = null;
                }
                if (groupModeEl.value === 'create' && createNickEl && !normStr(createNickEl.value)) {
                    createNickEl.value = suggestSchulOrgMailNickname(
                        (createNameEl && createNameEl.value) || area.label
                    );
                }
                syncM365ModeUi();
            });
        }
        if (groupKindEl) {
            groupKindEl.addEventListener('change', function () {
                area.m365.groupKind = normalizeSchulOrgGroupKind(groupKindEl.value);
                if (!schulOrgGroupKindIsEntraGroup(area.m365.groupKind)) {
                    area.m365.groupMode = 'none';
                    area.m365.groupId = null;
                    area.m365.groupLabel = null;
                    area.m365.groupMail = null;
                    if (groupModeEl) groupModeEl.value = 'none';
                }
                syncM365ModeUi();
            });
        }

        async function acquireAreaGraphToken() {
            let token = '';
            if (typeof window.ms365AuthAcquireTokenSilent === 'function') {
                try {
                    token = (await window.ms365AuthAcquireTokenSilent()) || '';
                } catch {
                    token = '';
                }
            }
            const gug = window.ms365GraphUnifiedGroups;
            if (!token && gug && typeof gug.getGraphToken === 'function') {
                token = await gug.getGraphToken();
            }
            if (!token && typeof window.ms365AuthAcquireToken === 'function') {
                token = await window.ms365AuthAcquireToken();
            }
            return token || '';
        }

        /**
         * ARGE: Gruppen-Besitzer → Stammdaten-Leitung (+ erscheint unter ARGE-Leiter).
         * @param {string} groupId
         */
        async function syncArgeLeadFromLinkedGroup(groupId) {
            if (!isArgeArea || !area.refKey || !groupId) return;
            const gug = window.ms365GraphUnifiedGroups;
            if (!gug || typeof gug.fetchGroupOwners !== 'function') return;
            try {
                const token = await acquireAreaGraphToken();
                if (!token) return;
                const owners = await gug.fetchGroupOwners(token, groupId).catch(function () {
                    return [];
                });
                const cur = loadSettings();
                const applied = applyArgeHeadFromGraphOwners(cur, area.refKey, owners);
                if (applied.changed) {
                    saveSettings(applied.settings);
                    toast(
                        'ARGE-Leitung aus Gruppen-Besitzer: ' +
                            (applied.head && (applied.head.name || applied.head.email)),
                        'success'
                    );
                    if (getActiveOverlay() === overlay && typeof ae.renderPeopleList === 'function') ae.renderPeopleList();
                    notify();
                }
            } catch (e) {
                console.warn('[schul-org] ARGE-Leitung aus Besitzern', e);
            }
        }

        /**
         * Klasse: Gruppen-Besitzer → Stammdaten-Klassenvorstand (Landkarte / Jahrgang).
         * @param {string} groupId
         */
        async function syncClassKvFromLinkedGroup(groupId) {
            if (!isClassArea || !area.refKey || !groupId) return;
            const gug = window.ms365GraphUnifiedGroups;
            if (!gug || typeof gug.fetchGroupOwners !== 'function') return;
            try {
                const token = await acquireAreaGraphToken();
                if (!token) return;
                const owners = await gug.fetchGroupOwners(token, groupId).catch(function () {
                    return [];
                });
                const cur = loadSettings();
                const applied = applyClassHeadFromGraphOwners(cur, area.refKey, owners);
                if (applied.changed) {
                    saveSettings(applied.settings);
                    toast(
                        'Klassenvorstand aus Gruppen-Besitzer: ' +
                            (applied.head && (applied.head.name || applied.head.email)),
                        'success'
                    );
                    if (getActiveOverlay() === overlay && typeof ae.renderPeopleList === 'function') ae.renderPeopleList();
                    notify();
                }
            } catch (e) {
                console.warn('[schul-org] Klassenvorstand aus Besitzern', e);
            }
        }

        const stammdatenOpenBtn = overlay.querySelector('[data-area-open-stammdaten]');
        if (stammdatenOpenBtn) {
            stammdatenOpenBtn.addEventListener('click', function () {
                if (isArgeArea && area.refKey) {
                    closeOverlay();
                    openArgeEditor(area.refKey);
                } else if (isClassArea && area.refKey) {
                    closeOverlay();
                    openClassEditor(area.refKey);
                } else if (isSgaBoardArea(area)) {
                    closeOverlay();
                    openSgaEditor(SCHUL_ORG_BOARD_REF);
                }
            });
        }

        /**
         * Warnung, wenn M365-Gruppe/Postfach schon an anderem Bereich hängt.
         * @param {'group'|'mailbox'} kind
         * @param {string} resourceId
         * @param {string} resourceLabel
         * @returns {Promise<boolean>} true = fortfahren
         */
        async function confirmM365ReuseIfNeeded(kind, resourceId, resourceLabel) {
            const layout = loadSettings().schulOrganigramm;
            const conflicts =
                kind === 'mailbox'
                    ? findSchulOrgNodesUsingM365Mailbox(layout, resourceId, {
                          excludeNodeId: area.id
                      })
                    : findSchulOrgNodesUsingM365Group(layout, resourceId, {
                          excludeNodeId: area.id
                      });
            if (!conflicts.length) return true;
            const list = conflicts
                .map(function (c) {
                    return schulOrgM365ConflictLabel(c);
                })
                .join(', ');
            const what =
                kind === 'mailbox'
                    ? 'Postfach „' + (resourceLabel || resourceId) + '“'
                    : 'Gruppe „' + (resourceLabel || resourceId) + '“';
            const ok = await confirmDlg(
                what +
                    ' ist auf der Landkarte bereits verknüpft mit: ' +
                    list +
                    '.\n\nTrotzdem diesem Bereich zuordnen? (Doppelte Nutzung kann Abgleich und Mitglieder vermischen.)',
                {
                    title: 'M365-Objekt bereits verwendet',
                    okText: 'Trotzdem zuordnen',
                    cancelText: 'Abbrechen'
                }
            );
            return !!ok;
        }

        const pickGroupBtn = overlay.querySelector('[data-area-pick-group]');
        if (pickGroupBtn) {
            pickGroupBtn.addEventListener('click', function () {
                if (groupModeEl && groupModeEl.value === 'none') groupModeEl.value = 'match';
                void toggleInlineEntraGroupPicker(
                    inlineEgpHost,
                    inlineEgpState,
                    {
                        hint:
                            'Gruppe für „' +
                            area.label +
                            '“ suchen (M365, Sicherheit, Verteiler) und antippen.',
                        onPick: function (picked) {
                            const gid = String((picked && picked.id) || '').trim();
                            const gLabel = (picked && (picked.displayName || picked.name)) || gid;
                            void confirmM365ReuseIfNeeded('group', gid, gLabel).then(function (ok) {
                                if (!ok) {
                                    toast(
                                        'Zuordnung abgebrochen – Gruppe wird schon woanders genutzt.',
                                        'info'
                                    );
                                    return;
                                }
                                const slotKind =
                                    ae.activePickKind ||
                                    inferSchulOrgGroupKind(picked) ||
                                    'unified';
                                writeBinding(slotKind, {
                                    mode: 'match',
                                    id: gid || null,
                                    label: gLabel,
                                    mail: normEmail(picked.mail || picked.email || '')
                                });
                                area.m365.groupMode = 'match';
                                if (groupModeEl) groupModeEl.value = 'match';
                                area.m365.groupKind = slotKind;
                                if (groupKindEl) groupKindEl.value = slotKind;
                                area.m365.groupId = gid || null;
                                area.m365.groupLabel = gLabel;
                                area.m365.groupMail = normEmail(picked.mail || picked.email || '');
                                refreshM365Sums();
                                toast(
                                    schulOrgM365ObjectKindLabel(slotKind) + ' zugeordnet.',
                                    'success'
                                );
                                if (isArgeArea && area.m365.groupId) {
                                    void syncArgeLeadFromLinkedGroup(area.m365.groupId);
                                }
                                if (isClassArea && area.m365.groupId) {
                                    void syncClassKvFromLinkedGroup(area.m365.groupId);
                                }
                            });
                        }
                    },
                    inlinePickers
                ).catch(function (e) {
                    toast(
                        'Gruppenauswahl fehlgeschlagen: ' + (e && e.message ? e.message : e),
                        'warning'
                    );
                });
            });
        }
        const clearGroupBtn = overlay.querySelector('[data-area-clear-group]');
        if (clearGroupBtn) {
            clearGroupBtn.addEventListener('click', function () {
                if (groupModeEl) groupModeEl.value = 'none';
                area.m365.groupMode = 'none';
                area.m365.groupId = null;
                area.m365.groupLabel = null;
                area.m365.groupMail = null;
                refreshM365Sums();
                // Sofort speichern – sonst bleibt die falsche Zuordnung auf der Landkarte
                const cur = loadSettings();
                let nextLayout = normalizeSchulOrganigramm(cur.schulOrganigramm);
                const labelEl = /** @type {HTMLInputElement} */ (
                    overlay.querySelector('#schulOrgAreaLabel')
                );
                const nextLabel = normStr(labelEl && labelEl.value) || area.label;
                const node = nextLayout.nodes.find(function (n) {
                    return n.id === area.id;
                });
                if (node) node.label = nextLabel;
                nextLayout = setSchulOrganigrammAreaM365(nextLayout, area.id, {
                    groupMode: 'none',
                    groupKind: normalizeSchulOrgGroupKind(
                        (groupKindEl && groupKindEl.value) || area.m365.groupKind
                    ),
                    groupId: null,
                    groupLabel: null,
                    groupMail: null,
                    mailboxMode: area.m365.mailboxMode || 'none',
                    mailboxUpn: area.m365.mailboxUpn || null,
                    mailboxLabel: area.m365.mailboxLabel || null,
                    mailboxId: area.m365.mailboxId || null
                });
                cur.schulOrganigramm = nextLayout;
                try {
                    window.__ms365SchulOrgSkipSyncOnce = true;
                } catch {
                    /* ignore */
                }
                saveSettings(cur);
                const savedNode = nextLayout.nodes.find(function (n) {
                    return n.id === area.id;
                });
                if (savedNode && savedNode.m365) {
                    area.m365 = Object.assign({}, savedNode.m365);
                }
                refreshM365Sums();
                toast('Gruppe gelöst und gespeichert.', 'success');
                notify();
            });
        }

        const pickOwnerBtn = overlay.querySelector('[data-area-pick-owner]');
        if (pickOwnerBtn) {
            pickOwnerBtn.addEventListener('click', function () {
                void toggleInlineEntraUserPicker(
                    inlineOwnerHost,
                    inlineOwnerState,
                    {
                        hint: 'Besitzer der neuen Gruppe in Microsoft 365 suchen und antippen (muss nicht Mitglied sein).',
                        onPick: function (picked) {
                            const mail = normEmail(
                                picked.mail || picked.userPrincipalName || picked.email
                            );
                            const name = normStr(picked.displayName || picked.name);
                            if (!mail || mail.indexOf('@') === -1) {
                                toast('Keine gültige E-Mail in der Auswahl.', 'warning');
                                return;
                            }
                            ae.customGroupOwner = {
                                email: mail,
                                label: name || mail
                            };
                            refreshGroupPeopleUi();
                            if (groupOwnerEl) groupOwnerEl.value = mail;
                            toast('Besitzer aus Microsoft 365 übernommen.', 'success');
                        }
                    },
                    inlinePickers
                ).catch(function (e) {
                    toast('M365: ' + (e && e.message ? e.message : String(e)), 'error');
                });
            });
        }

        const createBtn = overlay.querySelector('[data-area-create-group]');
        if (createBtn) {
            createBtn.addEventListener('click', function () {
                const kind = normalizeSchulOrgGroupKind(
                    (groupKindEl && groupKindEl.value) || area.m365.groupKind
                );
                if (!schulOrgGroupKindIsEntraGroup(kind)) {
                    toast(
                        'Teams-Gruppenchats werden im Werkzeug KlassenChats angelegt, nicht hier.',
                        'info'
                    );
                    return;
                }
                const displayName =
                    normStr(createNameEl && createNameEl.value) || normStr(area.label) || 'Bereich';
                const nick =
                    suggestSchulOrgMailNickname(
                        normStr(createNickEl && createNickEl.value) || displayName
                    );
                if (createNickEl) createNickEl.value = nick;
                const gug = window.ms365GraphUnifiedGroups;
                if (!gug || typeof gug.createGroupByKind !== 'function') {
                    toast('Graph-Modul für Gruppenanlage nicht geladen.', 'warning');
                    return;
                }
                createBtn.disabled = true;
                void (async function () {
                    try {
                        let token = '';
                        if (typeof window.ms365AuthAcquireTokenSilent === 'function') {
                            try {
                                token = (await window.ms365AuthAcquireTokenSilent()) || '';
                            } catch {
                                token = '';
                            }
                        }
                        if (!token && typeof gug.getGraphToken === 'function') {
                            token = await gug.getGraphToken();
                        }
                        if (!token && typeof window.ms365AuthAcquireToken === 'function') {
                            token = await window.ms365AuthAcquireToken();
                        }
                        if (!token) throw new Error('Anmeldung für Microsoft Graph fehlgeschlagen.');
                        const people = areaM365People();
                        const memberEmails = people.map(function (p) {
                            return p.email;
                        });
                        let ownerEmail = normEmail(groupOwnerEl && groupOwnerEl.value);
                        if (!ownerEmail && memberEmails.length) ownerEmail = memberEmails[0];
                        const created = await gug.createGroupByKind(
                            token,
                            kind,
                            displayName,
                            nick,
                            'Schul-Organigramm: ' + displayName
                        );
                        const gid = created && created.id ? String(created.id) : '';
                        if (!gid) throw new Error('Gruppe angelegt, aber keine ID zurückgegeben.');
                        let memberOk = 0;
                        let ownerOk = false;
                        if (typeof gug.sleep === 'function') await gug.sleep(1200);
                        if (ownerEmail) {
                            try {
                                if (typeof gug.resolveUserByEmail === 'function') {
                                    const ownerUser = await gug.resolveUserByEmail(token, ownerEmail);
                                    if (ownerUser && ownerUser.id) {
                                        if (typeof gug.addOwnerWithMemberFallback === 'function') {
                                            await gug.addOwnerWithMemberFallback(
                                                token,
                                                gid,
                                                ownerUser.id
                                            );
                                            ownerOk = true;
                                        } else if (typeof gug.ensureOwners === 'function') {
                                            await gug.ensureOwners(token, gid, [ownerEmail]);
                                            ownerOk = true;
                                        }
                                    }
                                } else if (typeof gug.ensureOwners === 'function') {
                                    await gug.ensureOwners(token, gid, [ownerEmail]);
                                    ownerOk = true;
                                }
                            } catch (ownerErr) {
                                console.warn(
                                    '[schul-org] Besitzer setzen fehlgeschlagen',
                                    ownerErr
                                );
                            }
                        } else if (typeof gug.ensureOwners === 'function') {
                            await gug.ensureOwners(token, gid, []);
                        }
                        if (memberEmails.length && typeof gug.syncEmailsToGroup === 'function') {
                            const sync = await gug.syncEmailsToGroup(
                                token,
                                gid,
                                memberEmails,
                                displayName
                            );
                            memberOk = (sync && sync.ok) || 0;
                        }
                        writeBinding(kind, {
                            mode: 'match',
                            id: gid,
                            label: (created && created.displayName) || displayName,
                            mail: normEmail((created && created.mail) || nick)
                        });
                        area.m365.groupMode = 'match';
                        if (groupModeEl) groupModeEl.value = 'match';
                        area.m365.groupKind = kind;
                        area.m365.groupId = gid;
                        area.m365.groupLabel =
                            (created && created.displayName) || displayName;
                        area.m365.groupMail = normEmail(
                            (created && created.mail) || nick
                        );
                        refreshM365Sums();
                        const bits = [schulOrgM365ObjectKindLabel(kind) + ' angelegt und verknüpft'];
                        if (memberEmails.length) {
                            bits.push(
                                memberOk +
                                    '/' +
                                    memberEmails.length +
                                    ' Mitglied(er)'
                            );
                        }
                        if (ownerEmail) {
                            bits.push(
                                ownerOk
                                    ? 'Besitzer: ' + ownerEmail
                                    : 'Besitzer konnte nicht gesetzt werden'
                            );
                        }
                        toast(bits.join(' · '), ownerOk || !ownerEmail ? 'success' : 'warning');
                    } catch (e) {
                        toast(
                            'Anlegen fehlgeschlagen: ' + (e && e.message ? e.message : e),
                            'warning'
                        );
                    } finally {
                        createBtn.disabled = false;
                    }
                })();
            });
        }

    ae.acquireAreaGraphToken = acquireAreaGraphToken;
    ae.syncArgeLeadFromLinkedGroup = syncArgeLeadFromLinkedGroup;
    ae.syncClassKvFromLinkedGroup = syncClassKvFromLinkedGroup;
    ae.confirmM365ReuseIfNeeded = confirmM365ReuseIfNeeded;
}
