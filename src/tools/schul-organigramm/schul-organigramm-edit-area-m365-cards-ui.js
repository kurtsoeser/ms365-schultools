/**
 * Area M365 card/mode helpers.
 */
import {
    normalizeSchulOrgGroupKind,
    schulOrgGroupKindLabel,
    schulOrgGroupKindIcon,
    schulOrgGroupKindIsEntraGroup,
    inferSchulOrgGroupKind,
    suggestSchulOrgMailNickname,
    SCHUL_ORG_M365_OBJECT_KINDS,
    normalizeSchulOrgM365,
    normalizeSchulOrgM365Binding,
    emptySchulOrgM365Bindings,
    schulOrgM365ObjectKindLabel,
    schulOrgM365ObjectKindIcon,
    isArgeAreaNode,
    isClassAreaNode
} from './schul-organigramm-logic.js';

/**
 * @param {Record<string, any>} ae
 */
export function wireAreaEditorM365Cards(ae) {

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


        function suggestedExchangeAdminUpn() {
            try {
                if (typeof window.ms365AuthGetAccountInfo === 'function') {
                    const info = window.ms365AuthGetAccountInfo();
                    const upn = normEmail((info && (info.upn || info.username)) || '');
                    if (upn && upn.indexOf('@') !== -1) return upn;
                }
            } catch {
                /* ignore */
            }
            return '';
        }
        if (groupModeEl) groupModeEl.value = area.m365.groupMode || 'none';
        if (groupKindEl) {
            groupKindEl.value = normalizeSchulOrgGroupKind(area.m365.groupKind || 'unified');
        }
        if (mailboxModeEl) {
            mailboxModeEl.value =
                area.m365.mailboxMode === 'create' ||
                area.m365.mailboxMode === 'match' ||
                area.m365.mailboxMode === 'none'
                    ? area.m365.mailboxMode
                    : 'none';
        }

        /** Welcher Objekttyp gerade per Entra-Picker gewählt wird */
        let activePickKind = normalizeSchulOrgGroupKind(area.m365.groupKind || 'unified');
        ae.activePickKind = activePickKind;

        /**
         * Mode aus Checkbox + Aktions-Select (match|create) lesen.
         * @param {Element} card
         * @returns {'none'|'match'|'create'}
         */
        function modeFromCard(card) {
            const useEl = /** @type {HTMLInputElement|null} */ (card.querySelector('[data-m365-use]'));
            if (!useEl || !useEl.checked) return 'none';
            const modeEl = /** @type {HTMLSelectElement|null} */ (card.querySelector('[data-m365-mode]'));
            const action = modeEl ? modeEl.value || 'match' : 'match';
            return action === 'create' ? 'create' : 'match';
        }

        /**
         * Checkbox / Aktions-Select / Match-Zeile an Mode anpassen.
         * @param {Element} card
         * @param {'none'|'match'|'create'} mode
         */
        function applyModeToCardUi(card, mode) {
            const used = mode !== 'none';
            const useEl = /** @type {HTMLInputElement|null} */ (card.querySelector('[data-m365-use]'));
            const modeEl = /** @type {HTMLSelectElement|null} */ (card.querySelector('[data-m365-mode]'));
            const modeWrap = card.querySelector('[data-m365-mode-wrap]');
            if (useEl) useEl.checked = used;
            if (modeWrap) modeWrap.hidden = !used;
            if (modeEl) {
                if (used && (mode === 'match' || mode === 'create')) {
                    const opt = modeEl.querySelector('option[value="' + mode + '"]');
                    if (opt && !/** @type {HTMLOptionElement} */ (opt).disabled) {
                        modeEl.value = mode;
                    } else if (!modeEl.value || modeEl.value === 'none') {
                        modeEl.value = 'match';
                    }
                }
            }
            card.classList.toggle('is-active', used);
            const matchRow = card.querySelector('[data-m365-match-row]');
            if (matchRow) matchRow.hidden = mode !== 'match';
            const chatHint = card.querySelector('[data-m365-chat-hint]');
            if (chatHint) chatHint.hidden = !used;
        }

        /**
         * @param {{ applyUi?: boolean }} [opts] applyUi=false beim Speichern (keine Side-Effects)
         */
        function readBindingsFromCards(opts) {
            const applyUi = !(opts && opts.applyUi === false);
            if (!area.m365.bindings) area.m365.bindings = emptySchulOrgM365Bindings();
            /** @type {Record<string, object>} */
            const nextBindings = emptySchulOrgM365Bindings();
            overlay.querySelectorAll('[data-m365-kind]').forEach(function (card) {
                const kind = String(card.getAttribute('data-m365-kind') || '');
                if (!kind) return;
                const mode = modeFromCard(card);
                const prev = area.m365.bindings[kind] || {};
                const next = normalizeSchulOrgM365Binding(
                    {
                        mode: mode,
                        id: prev.id,
                        label: prev.label,
                        mail: prev.mail,
                        upn: prev.upn,
                        webUrl: prev.webUrl
                    },
                    kind
                );
                if (mode === 'none') {
                    next.id = null;
                    next.label = null;
                    next.mail = null;
                    next.upn = null;
                    next.webUrl = null;
                }
                nextBindings[kind] = next;
                if (applyUi) applyModeToCardUi(card, mode);
            });
            area.m365.bindings = nextBindings;
            area.m365 = normalizeSchulOrgM365({ bindings: nextBindings }) || area.m365;
            if (!applyUi) return;
            // Create-Panel: erste Entra-Binding mit mode=create steuert Legacy-Selects
            let createKind = '';
            ['unified', 'security', 'mailSecurity', 'distribution'].forEach(function (kind) {
                if (createKind) return;
                const b = area.m365.bindings[kind];
                if (b && b.mode === 'create') createKind = kind;
            });
            if (createKind) {
                if (groupKindEl) groupKindEl.value = createKind;
                if (groupModeEl) groupModeEl.value = 'create';
                activePickKind = createKind;
                ae.activePickKind = activePickKind;
                const createLabel = overlay.querySelector('[data-area-create-kind-label]');
                if (createLabel) {
                    createLabel.textContent = schulOrgM365ObjectKindLabel(createKind);
                }
            } else {
                if (groupKindEl) {
                    groupKindEl.value = normalizeSchulOrgGroupKind(area.m365.groupKind || 'unified');
                }
                if (groupModeEl) groupModeEl.value = area.m365.groupMode || 'none';
            }
            if (mailboxModeEl) mailboxModeEl.value = area.m365.mailboxMode || 'none';
        }

        function writeBinding(kind, patch) {
            if (!area.m365.bindings) area.m365.bindings = emptySchulOrgM365Bindings();
            area.m365.bindings[kind] = normalizeSchulOrgM365Binding(
                Object.assign({}, area.m365.bindings[kind] || {}, patch),
                kind
            );
            area.m365 = normalizeSchulOrgM365({ bindings: area.m365.bindings }) || area.m365;
            if (groupKindEl) {
                groupKindEl.value = normalizeSchulOrgGroupKind(area.m365.groupKind || 'unified');
            }
            if (groupModeEl) groupModeEl.value = area.m365.groupMode || 'none';
            if (mailboxModeEl) mailboxModeEl.value = area.m365.mailboxMode || 'none';
            const card = /** @type {Element|null} */ (
                overlay.querySelector('[data-m365-kind="' + kind + '"]')
            );
            if (card) {
                const b = area.m365.bindings[kind];
                const mode = (b && b.mode) || 'none';
                applyModeToCardUi(card, mode);
                const sum = card.querySelector('[data-m365-sum]');
                if (sum && b) {
                    if (kind === 'mailbox') {
                        sum.textContent = [b.label, b.upn || b.id].filter(Boolean).join(' · ') ||
                            (mode === 'none' ? 'Nicht verknüpft.' : 'Zugewiesen.');
                    } else if (kind === 'groupChat') {
                        sum.textContent = [b.label, b.id].filter(Boolean).join(' · ') ||
                            'Chat-Status aktualisiert.';
                    } else {
                        sum.textContent = [b.label, b.mail, b.id].filter(Boolean).join(' · ') ||
                            (mode === 'none' ? 'Nicht verknüpft.' : 'Zugewiesen.');
                    }
                }
            }
        }

        function onCardModeChanged(card) {
            readBindingsFromCards();
            const kind = card ? String(card.getAttribute('data-m365-kind') || '') : '';
            const mode = card ? modeFromCard(card) : 'none';
            if (
                kind &&
                kind !== 'mailbox' &&
                kind !== 'groupChat' &&
                mode === 'create'
            ) {
                activePickKind = kind;
                ae.activePickKind = activePickKind;
                if (groupKindEl) groupKindEl.value = kind;
                if (groupModeEl) groupModeEl.value = 'create';
                const createLabel = overlay.querySelector('[data-area-create-kind-label]');
                if (createLabel) {
                    createLabel.textContent = schulOrgM365ObjectKindLabel(kind);
                }
            }
            if (kind === 'mailbox') {
                if (mailboxModeEl) mailboxModeEl.value = mode;
                area.m365.mailboxMode = mode === 'create' || mode === 'match' ? mode : 'none';
            }
            if (typeof ae.syncM365ModeUi === 'function') ae.syncM365ModeUi();
        }

        overlay.querySelectorAll('[data-m365-use]').forEach(function (chk) {
            chk.addEventListener('change', function () {
                const card = chk.closest('[data-m365-kind]');
                if (!card) return;
                const useEl = /** @type {HTMLInputElement} */ (chk);
                const modeEl = /** @type {HTMLSelectElement|null} */ (
                    card.querySelector('[data-m365-mode]')
                );
                const modeWrap = card.querySelector('[data-m365-mode-wrap]');
                if (useEl.checked) {
                    if (modeWrap) modeWrap.hidden = false;
                    if (modeEl) {
                        if (!modeEl.value || modeEl.value === 'none') modeEl.value = 'match';
                        // Gruppenchat: create ist disabled → immer match
                        const createOpt = modeEl.querySelector('option[value="create"]');
                        if (
                            createOpt &&
                            /** @type {HTMLOptionElement} */ (createOpt).disabled &&
                            modeEl.value === 'create'
                        ) {
                            modeEl.value = 'match';
                        }
                    }
                } else if (modeWrap) {
                    modeWrap.hidden = true;
                }
                onCardModeChanged(card);
            });
        });
        overlay.querySelectorAll('[data-m365-mode]').forEach(function (sel) {
            sel.addEventListener('change', function () {
                const card = sel.closest('[data-m365-kind]');
                if (!card) return;
                onCardModeChanged(card);
            });
        });
        overlay.querySelectorAll('[data-m365-pick]').forEach(function (btn) {
            btn.addEventListener('click', function () {
                const card = btn.closest('[data-m365-kind]');
                const kind = card ? String(card.getAttribute('data-m365-kind') || '') : '';
                if (kind === 'mailbox') {
                    const legacy = overlay.querySelector('[data-area-pick-mail]');
                    if (legacy) legacy.click();
                    return;
                }
                if (kind === 'groupChat') {
                    toast('Gruppenchats werden im Werkzeug KlassenChats verknüpft.', 'info');
                    return;
                }
                activePickKind = kind || 'unified';
                ae.activePickKind = activePickKind;
                if (groupKindEl) groupKindEl.value = activePickKind;
                if (groupModeEl) groupModeEl.value = 'match';
                writeBinding(activePickKind, { mode: 'match' });
                const legacy = overlay.querySelector('[data-area-pick-group]');
                if (legacy) legacy.click();
            });
        });
        overlay.querySelectorAll('[data-m365-clear]').forEach(function (btn) {
            btn.addEventListener('click', function () {
                const card = btn.closest('[data-m365-kind]');
                const kind = card ? String(card.getAttribute('data-m365-kind') || '') : '';
                if (!kind || !card) return;
                writeBinding(kind, {
                    mode: 'none',
                    id: null,
                    label: null,
                    mail: null,
                    upn: null,
                    webUrl: null
                });
                applyModeToCardUi(card, 'none');
                if (typeof ae.syncM365ModeUi === 'function') ae.syncM365ModeUi();
                toast(schulOrgM365ObjectKindLabel(kind) + ' gelöst.', 'success');
            });
        });

    ae.suggestedExchangeAdminUpn = suggestedExchangeAdminUpn;
    ae.modeFromCard = modeFromCard;
    ae.applyModeToCardUi = applyModeToCardUi;
    ae.readBindingsFromCards = readBindingsFromCards;
    ae.writeBinding = writeBinding;
    ae.onCardModeChanged = onCardModeChanged;
}
