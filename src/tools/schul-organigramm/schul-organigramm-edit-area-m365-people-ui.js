/**
 * Area M365 people/sum UI helpers.
 */
import { listM365PeopleForArea, listMailboxDelegatePeopleForArea } from './schul-organigramm-edit-people.js';
import {
    suggestSchulOrgMailNickname,
    schulOrgM365ObjectKindLabel,
    schulOrgGroupKindIsEntraGroup,
    normalizeSchulOrgGroupKind,
    schulOrgGroupKindIcon,
    schulOrgGroupKindLabel,
    SCHUL_ORG_M365_OBJECT_KINDS
} from './schul-organigramm-logic.js';

/**
 * @param {Record<string, any>} ae
 */
export function wireAreaEditorM365People(ae) {

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

        function schoolDomain() {
            const cur = loadSettings();
            return normStr(cur && cur.domain).replace(/^@/, '').toLowerCase();
        }

        function areaM365People() {
            return listM365PeopleForArea(loadSettings(), area);
        }

        function refreshGroupPeopleUi() {
            const people = areaM365People();
            if (groupPeopleEl) {
                if (!people.length) {
                    groupPeopleEl.textContent =
                        'Noch keine Personen mit M365-E-Mail in diesem Bereich, seinen Rollen oder Unterbereichen – die Gruppe wird ohne Mitglieder angelegt.';
                } else {
                    const nested = people.filter(function (p) {
                        return p.source === 'nested';
                    }).length;
                    groupPeopleEl.textContent =
                        'Als Mitglieder vorgesehen (' +
                        people.length +
                        (nested
                            ? ', davon ' + nested + ' aus Unterbereichen'
                            : '') +
                        '): ' +
                        people
                            .map(function (p) {
                                return p.label + (p.email ? ' <' + p.email + '>' : '');
                            })
                            .join(', ');
                }
            }
            if (groupOwnerEl) {
                const prev = normEmail(groupOwnerEl.value);
                const inPeople = function (em) {
                    return people.some(function (p) {
                        return p.email === em;
                    });
                };
                if (ae.customGroupOwner && inPeople(ae.customGroupOwner.email)) {
                    ae.customGroupOwner = null;
                }
                const custom =
                    ae.customGroupOwner && ae.customGroupOwner.email && !inPeople(ae.customGroupOwner.email)
                        ? ae.customGroupOwner
                        : null;
                groupOwnerEl.innerHTML =
                    '<option value="">— aus Bereich oder Microsoft 365 —</option>' +
                    (custom
                        ? '<option value="' +
                          escapeHtml(custom.email) +
                          '">' +
                          escapeHtml(custom.label + ' (' + custom.email + ') · M365') +
                          '</option>'
                        : '') +
                    people
                        .map(function (p) {
                            return (
                                '<option value="' +
                                escapeHtml(p.email) +
                                '">' +
                                escapeHtml(p.label + ' (' + p.email + ')') +
                                '</option>'
                            );
                        })
                        .join('');
                if (prev && (inPeople(prev) || (custom && custom.email === prev))) {
                    groupOwnerEl.value = prev;
                } else if (custom) {
                    groupOwnerEl.value = custom.email;
                } else if (!prev && people.length) {
                    groupOwnerEl.value = people[0].email;
                }
            }
        }

        function refreshMailDelegatesHint() {
            if (!mailDelegatesEl) return;
            const dels = areaM365People();
            if (!dels.length) {
                mailDelegatesEl.textContent =
                    'Noch keine Personen mit M365-E-Mail in diesem Bereich, seinen Rollen oder Unterbereichen – der Befehl legt nur das Postfach an.';
                return;
            }
            const nestedMail = dels.filter(function (p) {
                return p.source === 'nested';
            }).length;
            mailDelegatesEl.textContent =
                'Berechtigungen (FullAccess + SendAs) für ' +
                dels.length +
                (dels.length === 1 ? ' Person' : ' Personen') +
                (nestedMail
                    ? ' (davon ' + nestedMail + ' aus Unterbereichen)'
                    : '') +
                ': ' +
                dels
                    .map(function (p) {
                        return p.label + (p.email ? ' <' + p.email + '>' : '');
                    })
                    .join(', ');
        }

        function ensureMailCreateDefaults() {
            if (!mailCreateNameEl || !mailCreateAliasEl || !mailCreateSmtpEl) return;
            if (!normStr(mailCreateNameEl.value)) mailCreateNameEl.value = area.label || '';
            refreshMailDelegatesHint();
            void import('../../shared/exchange-shared-mailbox-script.js').then(function (mod) {
                const sug = mod.suggestSharedMailboxIdentity(
                    mailCreateNameEl.value || area.label,
                    schoolDomain()
                );
                if (!normStr(mailCreateAliasEl.value)) mailCreateAliasEl.value = sug.alias;
                if (!normStr(mailCreateSmtpEl.value) && sug.primarySmtp) {
                    mailCreateSmtpEl.value = sug.primarySmtp;
                }
            });
        }

        function syncM365ModeUi() {
            const kind = normalizeSchulOrgGroupKind(
                (groupKindEl && groupKindEl.value) || area.m365.groupKind || 'unified'
            );
            const isChat = kind === 'groupChat';
            const kindIcon = overlay.querySelector('[data-area-kind-icon]');
            if (kindIcon) {
                kindIcon.innerHTML = '<i class="bi ' + schulOrgGroupKindIcon(kind) + '"></i>';
            }
            // Gruppenchat: keine Entra-Match/Create-UI (Anlage über KlassenChats)
            if (groupModeEl) {
                if (isChat) {
                    if (groupModeEl.value === 'create' || groupModeEl.value === 'match') {
                        groupModeEl.value = 'none';
                        area.m365.groupMode = 'none';
                        area.m365.groupId = null;
                        area.m365.groupLabel = null;
                        area.m365.groupMail = null;
                    }
                    Array.prototype.forEach.call(groupModeEl.options, function (opt) {
                        const v = String(opt.value || '');
                        opt.disabled = v === 'match' || v === 'create';
                        if (v === 'none') opt.textContent = 'Kein Entra-Objekt (Chat separat)';
                        else if (v === 'match') {
                            opt.textContent = 'Bestehende zuordnen (nur Entra-Gruppen)';
                        } else if (v === 'create') {
                            opt.textContent = 'Neu anlegen (nur Entra-Gruppen)';
                        }
                    });
                } else {
                    Array.prototype.forEach.call(groupModeEl.options, function (opt) {
                        opt.disabled = false;
                        const v = String(opt.value || '');
                        if (v === 'none') opt.textContent = 'Keine Gruppe';
                        else if (v === 'match') opt.textContent = 'Bestehende zuordnen';
                        else if (v === 'create') opt.textContent = 'Neu in Microsoft 365 anlegen';
                    });
                }
            }
            const mode = (groupModeEl && groupModeEl.value) || area.m365.groupMode || 'none';
            const matchRow = overlay.querySelector('[data-area-group-match-row]');
            const createRow = overlay.querySelector('[data-area-group-create-row]');
            const hint = overlay.querySelector('[data-area-create-hint]');
            if (matchRow) matchRow.hidden = isChat || mode === 'create';
            if (createRow) createRow.hidden = isChat || mode !== 'create';
            if (mode === 'create' || mode === 'none' || isChat) {
                if (inlineEgpState.api) {
                    inlineEgpState.api.destroy();
                    inlineEgpState.api = null;
                }
                if (inlineEgpHost) inlineEgpHost.hidden = true;
            }
            if (mode !== 'create' || isChat) {
                if (inlineOwnerState.api) {
                    inlineOwnerState.api.destroy();
                    inlineOwnerState.api = null;
                }
                if (inlineOwnerHost) inlineOwnerHost.hidden = true;
            }
            if (hint) {
                const memberNote =
                    ' Personen aus dem Bereich, Rollen und Unterbereichen werden flach als Mitglieder gesetzt; Besitzer aus Bereich oder Microsoft 365.';
                if (kind === 'groupChat') {
                    hint.textContent =
                        'Teams-Gruppenchat: Anlage und Sync laufen über das Werkzeug KlassenChats. Hier nur Objekttyp/Intent – auf der Landkarte erscheint das Chat-Icon, sobald ein Chat verknüpft ist.';
                } else if (kind === 'distribution') {
                    hint.textContent =
                        'Verteilerlisten: Graph legt eine mailaktivierte Cloud-Gruppe an (ohne Team). Bei Tenant-Einschränkungen ggf. Exchange nutzen.' +
                        memberNote;
                } else if (kind === 'security') {
                    hint.textContent = 'Reine Sicherheitsgruppe (ohne E-Mail) in Entra.' + memberNote;
                } else if (kind === 'mailSecurity') {
                    hint.textContent = 'E-Mail-aktivierte Sicherheitsgruppe in Entra.' + memberNote;
                } else {
                    hint.textContent =
                        'Microsoft 365-Gruppe (Unified) – optional später Team möglich.' + memberNote;
                }
            }
            const mailMode = (mailboxModeEl && mailboxModeEl.value) || area.m365.mailboxMode || 'none';
            const mailMatchRow = overlay.querySelector('[data-area-mail-match-row]');
            const mailCreateRow = overlay.querySelector('[data-area-mail-create-row]');
            if (mailMatchRow) mailMatchRow.hidden = mailMode === 'create';
            if (mailCreateRow) mailCreateRow.hidden = mailMode !== 'create';
            if (mode === 'create' && !isChat) {
                refreshGroupPeopleUi();
            }
            if (mailMode === 'create') {
                ensureMailCreateDefaults();
                if (inlineMailState.api) {
                    inlineMailState.api.destroy();
                    inlineMailState.api = null;
                }
                if (inlineMailHost) inlineMailHost.hidden = true;
            } else if (mailMode === 'none') {
                if (inlineMailState.api) {
                    inlineMailState.api.destroy();
                    inlineMailState.api = null;
                }
                if (inlineMailHost) inlineMailHost.hidden = true;
            }
            refreshM365Sums();
        }

        function refreshM365Sums() {
            const gSum = overlay.querySelector('[data-area-group-sum]');
            const cSum = overlay.querySelector('[data-area-chat-sum]');
            const mSum = overlay.querySelector('[data-area-mail-sum]');
            const kind = normalizeSchulOrgGroupKind(
                (groupKindEl && groupKindEl.value) || area.m365.groupKind
            );
            const chatId = normStr(area.meta && area.meta.chatId);
            const chatTopic = normStr(area.meta && area.meta.chatTopic);
            const chatUrl = normStr(area.meta && area.meta.chatWebUrl);
            if (cSum) {
                if (kind === 'groupChat' || chatId) {
                    cSum.hidden = false;
                    if (chatId) {
                        cSum.textContent = [
                            'Teams-Gruppenchat verknüpft',
                            chatTopic,
                            chatUrl || chatId
                        ]
                            .filter(Boolean)
                            .join(' · ');
                    } else if (kind === 'groupChat') {
                        cSum.textContent =
                            'Noch kein Chat verknüpft – im Werkzeug KlassenChats anlegen oder synchronisieren.';
                    } else {
                        cSum.hidden = true;
                        cSum.textContent = '';
                    }
                } else {
                    cSum.hidden = true;
                    cSum.textContent = '';
                }
            }
            if (gSum) {
                if (kind === 'groupChat') {
                    gSum.textContent =
                        'Objekttyp Teams-Gruppenchat – Entra-Gruppe optional separat wählbar (anderer Objekttyp).';
                } else if (area.m365.groupMode === 'none' || !area.m365.groupId) {
                    gSum.textContent =
                        (groupModeEl && groupModeEl.value) === 'create'
                            ? 'Noch nicht angelegt – Name/Alias prüfen und anlegen.'
                            : 'Keine Gruppe verknüpft.';
                } else {
                    gSum.textContent = [
                        schulOrgGroupKindLabel(area.m365.groupKind),
                        area.m365.groupLabel,
                        area.m365.groupMail,
                        area.m365.groupId
                    ]
                        .filter(Boolean)
                        .join(' · ');
                }
            }
            if (mSum) {
                const mailModeUi = (mailboxModeEl && mailboxModeEl.value) || area.m365.mailboxMode;
                if (mailModeUi === 'create' && area.m365.mailboxUpn) {
                    mSum.textContent =
                        'Zum Anlegen vorgemerkt: ' +
                        [area.m365.mailboxLabel, area.m365.mailboxUpn].filter(Boolean).join(' · ') +
                        ' (.cmd-Assistent in Exchange ausführen)';
                } else if (area.m365.mailboxMode === 'none' || !area.m365.mailboxUpn) {
                    mSum.textContent =
                        mailModeUi === 'create'
                            ? 'Noch nicht vorgemerkt – Name/Alias/E-Mail prüfen und Assistent erzeugen.'
                            : 'Kein Postfach verknüpft.';
                } else {
                    mSum.textContent = [area.m365.mailboxLabel, area.m365.mailboxUpn]
                        .filter(Boolean)
                        .join(' · ');
                }
            }
        }

    ae.schoolDomain = schoolDomain;
    ae.areaM365People = areaM365People;
    ae.refreshGroupPeopleUi = refreshGroupPeopleUi;
    ae.refreshMailDelegatesHint = refreshMailDelegatesHint;
    ae.ensureMailCreateDefaults = ensureMailCreateDefaults;
    ae.syncM365ModeUi = syncM365ModeUi;
    ae.refreshM365Sums = refreshM365Sums;
}
