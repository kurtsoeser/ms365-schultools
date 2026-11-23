/**
 * Area M365 mailbox listeners.
 */
import {
    suggestSchulOrgMailNickname,
    findSchulOrgNodesUsingM365Mailbox,
    schulOrgM365ConflictLabel,
    normalizeSchulOrgM365Binding,
    schulOrgM365ObjectKindLabel
} from './schul-organigramm-logic.js';

/**
 * @param {Record<string, any>} ae
 */
export function wireAreaEditorM365Mail(ae) {

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
    const refreshM365Sums = ae.refreshM365Sums;
    const refreshMailDelegatesHint = ae.refreshMailDelegatesHint;
    const areaM365People = ae.areaM365People;
    const schoolDomain = ae.schoolDomain;
    const ensureMailCreateDefaults = ae.ensureMailCreateDefaults;
    const suggestedExchangeAdminUpn = ae.suggestedExchangeAdminUpn;
    const confirmM365ReuseIfNeeded = ae.confirmM365ReuseIfNeeded;

        if (mailboxModeEl) {
            mailboxModeEl.addEventListener('change', function () {
                const mode = mailboxModeEl.value || 'none';
                area.m365.mailboxMode = mode === 'create' || mode === 'match' ? mode : 'none';
                if (mode === 'none') {
                    area.m365.mailboxUpn = null;
                    area.m365.mailboxLabel = null;
                    area.m365.mailboxId = null;
                    if (mailScriptEl) {
                        mailScriptEl.value = '';
                        mailScriptEl.hidden = true;
                    }
                    ae.mailCmdDownloadText = '';
                    if (mailScriptCopyBtn) mailScriptCopyBtn.hidden = true;
                    if (mailScriptDlBtn) mailScriptDlBtn.hidden = true;
                }
                syncM365ModeUi();
            });
        }

        const pickMailBtn = overlay.querySelector('[data-area-pick-mail]');
        if (pickMailBtn) {
            pickMailBtn.addEventListener('click', function () {
                if (mailboxModeEl && mailboxModeEl.value === 'none') mailboxModeEl.value = 'match';
                void toggleInlineEntraUserPicker(
                    inlineMailHost,
                    inlineMailState,
                    {
                        hint:
                            'Postfach für „' +
                            area.label +
                            '“ suchen (freigegebene Postfächer oft als Benutzer) und antippen.',
                        onPick: function (picked) {
                            const upn = normEmail(picked.mail || picked.userPrincipalName || '');
                            const mid = String((picked && picked.id) || '').trim();
                            const mLabel = (picked && picked.displayName) || upn || mid;
                            void confirmM365ReuseIfNeeded('mailbox', upn || mid, mLabel).then(
                                function (ok) {
                                    if (!ok) {
                                        toast(
                                            'Zuordnung abgebrochen – Postfach wird schon woanders genutzt.',
                                            'info'
                                        );
                                        return;
                                    }
                                    writeBinding('mailbox', {
                                        mode: 'match',
                                        id: mid || null,
                                        label: mLabel,
                                        upn: upn
                                    });
                                    area.m365.mailboxMode = 'match';
                                    if (mailboxModeEl) mailboxModeEl.value = 'match';
                                    area.m365.mailboxUpn = upn;
                                    area.m365.mailboxLabel = mLabel;
                                    area.m365.mailboxId = mid || null;
                                    refreshM365Sums();
                                    toast('Postfach zugeordnet.', 'success');
                                }
                            );
                        }
                    },
                    inlinePickers
                ).catch(function (e) {
                    toast(
                        'Postfachauswahl fehlgeschlagen: ' + (e && e.message ? e.message : e),
                        'warning'
                    );
                });
            });
        }
        const clearMailBtn = overlay.querySelector('[data-area-clear-mail]');
        if (clearMailBtn) {
            clearMailBtn.addEventListener('click', function () {
                if (mailboxModeEl) mailboxModeEl.value = 'none';
                area.m365.mailboxMode = 'none';
                area.m365.mailboxUpn = null;
                area.m365.mailboxLabel = null;
                area.m365.mailboxId = null;
                if (mailScriptEl) {
                    mailScriptEl.value = '';
                    mailScriptEl.hidden = true;
                }
                ae.mailCmdDownloadText = '';
                if (mailScriptCopyBtn) mailScriptCopyBtn.hidden = true;
                if (mailScriptDlBtn) mailScriptDlBtn.hidden = true;
                refreshM365Sums();
            });
        }

        const createMailBtn = overlay.querySelector('[data-area-create-mail]');
        if (createMailBtn) {
            createMailBtn.addEventListener('click', function () {
                void import('../../shared/exchange-shared-mailbox-script.js')
                    .then(function (mod) {
                        const displayName =
                            normStr(mailCreateNameEl && mailCreateNameEl.value) ||
                            normStr(area.label) ||
                            'Postfach';
                        let alias = suggestSchulOrgMailNickname(
                            normStr(mailCreateAliasEl && mailCreateAliasEl.value) || displayName
                        );
                        if (mailCreateAliasEl) mailCreateAliasEl.value = alias;
                        let smtp = normEmail(mailCreateSmtpEl && mailCreateSmtpEl.value);
                        if (!smtp) {
                            const sug = mod.suggestSharedMailboxIdentity(displayName, schoolDomain());
                            smtp = sug.primarySmtp;
                            if (mailCreateSmtpEl && smtp) mailCreateSmtpEl.value = smtp;
                        }
                        if (!smtp || smtp.indexOf('@') === -1) {
                            toast(
                                'Bitte eine gültige E-Mail angeben (z. B. alias@' +
                                    (schoolDomain() || 'schule.at') +
                                    ').',
                                'warning'
                            );
                            return;
                        }
                        const delegates = areaM365People();
                        const mailOpts = {
                            displayName: displayName,
                            alias: alias,
                            primarySmtp: smtp,
                            description: 'Schul-Organigramm: ' + displayName,
                            delegates: delegates.map(function (p) {
                                return p.email;
                            }),
                            rights: ['FullAccess', 'SendAs'],
                            autoMapping: true,
                            suggestedAdminUpn: suggestedExchangeAdminUpn()
                        };
                        const script = mod.buildSharedMailboxCreateScript(mailOpts);
                        ae.mailCmdDownloadText =
                            typeof mod.buildSharedMailboxCreateCmd === 'function'
                                ? mod.buildSharedMailboxCreateCmd(mailOpts)
                                : script;
                        writeBinding('mailbox', {
                            mode: 'create',
                            id: null,
                            label: displayName,
                            upn: smtp
                        });
                        area.m365.mailboxMode = 'create';
                        if (mailboxModeEl) mailboxModeEl.value = 'create';
                        area.m365.mailboxUpn = smtp;
                        area.m365.mailboxLabel = displayName;
                        area.m365.mailboxId = null;
                        refreshMailDelegatesHint();
                        if (mailScriptEl) {
                            mailScriptEl.value = script;
                            mailScriptEl.hidden = false;
                        }
                        if (mailScriptCopyBtn) mailScriptCopyBtn.hidden = false;
                        if (mailScriptDlBtn) mailScriptDlBtn.hidden = false;
                        refreshM365Sums();
                        toast(
                            delegates.length
                                ? 'Postfach vorgemerkt. .cmd herunterladen und per Doppelklick starten – inkl. Rechte für ' +
                                      delegates.length +
                                      (delegates.length === 1 ? ' Person.' : ' Personen.')
                                : 'Postfach vorgemerkt. .cmd herunterladen und per Doppelklick starten.',
                            'success'
                        );
                    })
                    .catch(function (e) {
                        toast(
                            'Assistent fehlgeschlagen: ' + (e && e.message ? e.message : e),
                            'warning'
                        );
                    });
            });
        }
        if (mailScriptCopyBtn) {
            mailScriptCopyBtn.addEventListener('click', function () {
                const text = mailScriptEl ? mailScriptEl.value : '';
                if (!text) return;
                void navigator.clipboard.writeText(text).then(
                    function () {
                        toast('Geführten PowerShell-Assistenten kopiert.', 'success');
                    },
                    function () {
                        if (mailScriptEl) {
                            mailScriptEl.focus();
                            mailScriptEl.select();
                        }
                        toast('Kopieren manuell: Text markieren und Strg+C.', 'info');
                    }
                );
            });
        }
        if (mailScriptDlBtn) {
            mailScriptDlBtn.addEventListener('click', function () {
                const text = ae.mailCmdDownloadText || (mailScriptEl ? mailScriptEl.value : '');
                if (!text) return;
                const alias =
                    suggestSchulOrgMailNickname(
                        (mailCreateAliasEl && mailCreateAliasEl.value) || area.label || 'postfach'
                    ) || 'postfach';
                // UTF-8 BOM: Windows PowerShell 5.1 liest -Encoding UTF8 sonst falsch
                const blob = new Blob(['\uFEFF' + text], {
                    type: 'application/x-bat;charset=utf-8'
                });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'shared-mailbox-' + alias + '.cmd';
                document.body.appendChild(a);
                a.click();
                a.remove();
                URL.revokeObjectURL(url);
                toast('.cmd gespeichert – per Doppelklick starten und Admin-Konto wählen.', 'success');
            });
        }
}
