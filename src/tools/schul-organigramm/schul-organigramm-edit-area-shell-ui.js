/**
 * Area editor overlay shell (mount + DOM refs).
 */
import { isArgeAreaNode, isClassAreaNode } from './schul-organigramm-logic.js';
import { buildAreaEditorOverlay } from './schul-organigramm-edit-area-shell-html-ui.js';

/**
 * @param {Record<string, any>} ctx
 * @param {string|object} areaIdOrNode
 * @param {{ removeArea: function }} api
 * @returns {Record<string, any>|null}
 */
export function mountAreaEditorShell(ctx, areaIdOrNode, api) {
    const built = buildAreaEditorOverlay(ctx, areaIdOrNode);
    if (!built) return null;
    const overlay = built.overlay;
    const area = built.area;
    const settings = built.settings;
    const sammelKey = built.sammelKey;
    const isKlassenchatArea = built.isKlassenchatArea;
    const closeOverlay = ctx.closeOverlay;
    const onEsc = ctx.onEsc;
    const setActiveOverlay = ctx.setActiveOverlay;
    const removeArea = api.removeArea;

    document.body.appendChild(overlay);
    setActiveOverlay(overlay);
    document.addEventListener('keydown', onEsc, true);

        // Tabs: Personen | Microsoft 365-Objekte
        (function wireAreaEditorTabs() {
            const tabRoot = overlay.querySelector('[data-area-tabs]');
            if (!tabRoot) return;
            const buttons = tabRoot.querySelectorAll('[data-area-tab]');
            const panels = tabRoot.querySelectorAll('[data-area-panel]');
            function activate(tabId) {
                const id = tabId === 'm365' ? 'm365' : 'people';
                buttons.forEach(function (btn) {
                    const on = String(btn.getAttribute('data-area-tab') || '') === id;
                    btn.classList.toggle('is-active', on);
                    btn.setAttribute('aria-selected', on ? 'true' : 'false');
                    btn.tabIndex = on ? 0 : -1;
                });
                panels.forEach(function (panel) {
                    const on = String(panel.getAttribute('data-area-panel') || '') === id;
                    panel.classList.toggle('is-active', on);
                    panel.hidden = !on;
                });
            }
            buttons.forEach(function (btn) {
                btn.addEventListener('click', function () {
                    activate(String(btn.getAttribute('data-area-tab') || 'people'));
                });
            });
            tabRoot.addEventListener('keydown', function (ev) {
                const keys = ['ArrowLeft', 'ArrowRight', 'Home', 'End'];
                if (keys.indexOf(ev.key) === -1) return;
                const list = Array.prototype.slice.call(buttons);
                if (!list.length) return;
                let idx = list.findIndex(function (b) {
                    return b.classList.contains('is-active');
                });
                if (idx < 0) idx = 0;
                if (ev.key === 'ArrowRight') idx = (idx + 1) % list.length;
                else if (ev.key === 'ArrowLeft') idx = (idx - 1 + list.length) % list.length;
                else if (ev.key === 'Home') idx = 0;
                else if (ev.key === 'End') idx = list.length - 1;
                ev.preventDefault();
                list[idx].focus();
                activate(String(list[idx].getAttribute('data-area-tab') || 'people'));
            });
        })();

        const inlineEupHost = /** @type {HTMLElement|null} */ (
            overlay.querySelector('[data-schul-org-inline-eup]')
        );
        const inlineEgpHost = /** @type {HTMLElement|null} */ (
            overlay.querySelector('[data-schul-org-inline-egp]')
        );
        const inlineMailHost = /** @type {HTMLElement|null} */ (
            overlay.querySelector('[data-schul-org-inline-eup-mail]')
        );
        const inlineOwnerHost = /** @type {HTMLElement|null} */ (
            overlay.querySelector('[data-schul-org-inline-eup-owner]')
        );
        /** @type {{ api: { destroy: () => void, focus: () => void }|null }} */
        const inlineEupState = { api: null };
        /** @type {{ api: { destroy: () => void, focus: () => void }|null }} */
        const inlineEgpState = { api: null };
        /** @type {{ api: { destroy: () => void, focus: () => void }|null }} */
        const inlineMailState = { api: null };
        /** @type {{ api: { destroy: () => void, focus: () => void }|null }} */
        const inlineOwnerState = { api: null };
        const inlinePickers = [inlineEupState, inlineEgpState, inlineMailState, inlineOwnerState];
        /** @type {{ email: string, label: string }|null} */
        let customGroupOwner = null;

        const delAreaBtn = overlay.querySelector('[data-schul-org-delete-area]');
        if (delAreaBtn) {
            delAreaBtn.addEventListener('click', function () {
                void removeArea(area).then(function (ok) {
                    if (ok) closeOverlay();
                });
            });
        }

        const groupModeEl = /** @type {HTMLSelectElement|null} */ (
            overlay.querySelector('#schulOrgAreaGroupMode')
        );
        const groupKindEl = /** @type {HTMLSelectElement|null} */ (
            overlay.querySelector('#schulOrgAreaGroupKind')
        );
        const mailboxModeEl = /** @type {HTMLSelectElement|null} */ (
            overlay.querySelector('#schulOrgAreaMailboxMode')
        );
        const createNameEl = /** @type {HTMLInputElement|null} */ (
            overlay.querySelector('#schulOrgAreaCreateName')
        );
        const createNickEl = /** @type {HTMLInputElement|null} */ (
            overlay.querySelector('#schulOrgAreaCreateNick')
        );
        const mailCreateNameEl = /** @type {HTMLInputElement|null} */ (
            overlay.querySelector('#schulOrgAreaMailCreateName')
        );
        const mailCreateAliasEl = /** @type {HTMLInputElement|null} */ (
            overlay.querySelector('#schulOrgAreaMailCreateAlias')
        );
        const mailCreateSmtpEl = /** @type {HTMLInputElement|null} */ (
            overlay.querySelector('#schulOrgAreaMailCreateSmtp')
        );
        const mailScriptEl = /** @type {HTMLTextAreaElement|null} */ (
            overlay.querySelector('[data-area-mail-script]')
        );
        const mailScriptCopyBtn = /** @type {HTMLButtonElement|null} */ (
            overlay.querySelector('[data-area-mail-script-copy]')
        );
        const mailScriptDlBtn = /** @type {HTMLButtonElement|null} */ (
            overlay.querySelector('[data-area-mail-script-dl]')
        );
        const mailDelegatesEl = /** @type {HTMLElement|null} */ (
            overlay.querySelector('[data-area-mail-delegates]')
        );
        const groupPeopleEl = /** @type {HTMLElement|null} */ (
            overlay.querySelector('[data-area-group-people]')
        );
        const groupOwnerEl = /** @type {HTMLSelectElement|null} */ (
            overlay.querySelector('[data-area-group-owner]')
        );
        /** @type {string} */
        let mailCmdDownloadText = '';

    return {
        overlay: overlay,
        area: area,
        settings: settings,
        sammelKey: sammelKey,
        isKlassenchatArea: isKlassenchatArea,
        isArgeArea: isArgeAreaNode(area),
        isClassArea: isClassAreaNode(area),
        loadSettings: ctx.loadSettings,
        saveSettings: ctx.saveSettings,
        toast: ctx.toast,
        confirmDlg: ctx.confirmDlg,
        notify: ctx.notify,
        closeOverlay: ctx.closeOverlay,
        escapeHtml: ctx.escapeHtml,
        normStr: ctx.normStr,
        normEmail: ctx.normEmail,
        normCode: ctx.normCode,
        toggleInlineEntraUserPicker: ctx.toggleInlineEntraUserPicker,
        toggleInlineEntraGroupPicker: ctx.toggleInlineEntraGroupPicker,
        stripEphemeralSchoolSettings: ctx.stripEphemeralSchoolSettings,
        plainJsonClone: ctx.plainJsonClone,
        settingsForPersist: ctx.settingsForPersist,
        suggestLocalPersonEmail: ctx.suggestLocalPersonEmail,
        persistOrganigrammLayout: ctx.persistOrganigrammLayout,
        removeArea: removeArea,
        openArgeEditor: ctx.openArgeEditor,
        openClassEditor: ctx.openClassEditor,
        openSgaEditor: ctx.openSgaEditor,
        inlineEupHost: inlineEupHost,
        inlineEgpHost: inlineEgpHost,
        inlineMailHost: inlineMailHost,
        inlineOwnerHost: inlineOwnerHost,
        inlineEupState: inlineEupState,
        inlineEgpState: inlineEgpState,
        inlineMailState: inlineMailState,
        inlineOwnerState: inlineOwnerState,
        inlinePickers: inlinePickers,
        groupModeEl: groupModeEl,
        groupKindEl: groupKindEl,
        mailboxModeEl: mailboxModeEl,
        createNameEl: createNameEl,
        createNickEl: createNickEl,
        mailCreateNameEl: mailCreateNameEl,
        mailCreateAliasEl: mailCreateAliasEl,
        mailCreateSmtpEl: mailCreateSmtpEl,
        mailScriptEl: mailScriptEl,
        mailScriptCopyBtn: mailScriptCopyBtn,
        mailScriptDlBtn: mailScriptDlBtn,
        mailDelegatesEl: mailDelegatesEl,
        groupPeopleEl: groupPeopleEl,
        groupOwnerEl: groupOwnerEl,
        customGroupOwner: customGroupOwner,
        mailCmdDownloadText: mailCmdDownloadText
    };
}
