/**
 * Dialog HTML builders for Schul-Organigramm edit.
 */
import { escapeHtml, normStr } from './schul-organigramm-edit-utils.js';
import {
    SCHUL_ORG_M365_OBJECT_KINDS,
    schulOrgM365ObjectKindLabel,
    schulOrgM365ObjectKindIcon,
    normalizeSchulOrgM365,
    normalizeSchulOrgM365Binding,
    emptySchulOrgM365Bindings
} from './schul-organigramm-logic.js';

export function editDialogHtml(opts) {
    const icon = opts.icon || 'bi-pencil-square';
    const title = opts.title || 'Bearbeiten';
    const sub = opts.sub || '';
    const body = opts.body || '';
    const foot = opts.foot || '';
    const wide = !!(opts && opts.wide);
    return (
        '<div class="modal-box schul-org-edit-box' +
        (wide ? ' schul-org-edit-box--wide' : '') +
        '" role="dialog" aria-modal="true" aria-labelledby="schulOrgEditTitle" tabindex="-1">' +
        '<header class="schul-org-edit-head">' +
        '<div class="schul-org-edit-head-main">' +
        '<span class="schul-org-edit-icon" aria-hidden="true"><i class="bi ' +
        icon +
        '"></i></span>' +
        '<div class="schul-org-edit-titles">' +
        '<h3 id="schulOrgEditTitle" class="schul-org-edit-title">' +
        escapeHtml(title) +
        '</h3>' +
        (sub ? '<p class="schul-org-edit-sub">' + escapeHtml(sub) + '</p>' : '') +
        '</div></div>' +
        '<button type="button" class="schul-org-edit-close" data-schul-org-close aria-label="Schließen">' +
        '<i class="bi bi-x-lg" aria-hidden="true"></i></button>' +
        '</header>' +
        '<div class="schul-org-edit-body">' +
        body +
        '</div>' +
        '<footer class="schul-org-edit-foot">' +
        foot +
        '</footer></div>'
    );
}

/**
 * Sechs parallel nutzbare M365-Objektkarten für „Bereich bearbeiten“.
 * @param {object} area
 * @param {{ isKlassenchatArea?: boolean }} [opts]
 */
export function areaM365ObjectsHtml(area, opts) {
    const o = opts && typeof opts === 'object' ? opts : {};
    const m365 = normalizeSchulOrgM365(area && area.m365 ? area.m365 : {}) || {
        bindings: emptySchulOrgM365Bindings()
    };
    const chatId = normStr(area && area.meta && area.meta.chatId);
    const chatTopic = normStr(area && area.meta && area.meta.chatTopic);
    return SCHUL_ORG_M365_OBJECT_KINDS.map(function (kind) {
        const b = m365.bindings[kind] || normalizeSchulOrgM365Binding(null, kind);
        const isChat = kind === 'groupChat';
        const isMail = kind === 'mailbox';
        const isEntra = !isChat && !isMail;
        let mode = b.mode || 'none';
        if (isChat && o.isKlassenchatArea && chatId && mode === 'none') mode = 'match';
        let sum = 'Nicht verknüpft.';
        if (isChat) {
            if (chatId || b.id) {
                sum = ['Verknüpft', chatTopic || b.label, chatId || b.id].filter(Boolean).join(' · ');
            } else if (mode !== 'none') {
                sum =
                    'Noch kein Chat – im Werkzeug KlassenChats anlegen oder synchronisieren.';
            } else if (o.isKlassenchatArea) {
                sum = 'KlassenChat-Bereich · Chat über das Werkzeug KlassenChats.';
            }
        } else if (isMail) {
            if (mode === 'create' && b.upn) {
                sum = 'Zum Anlegen vorgemerkt: ' + [b.label, b.upn].filter(Boolean).join(' · ');
            } else if (mode !== 'none' && (b.upn || b.id)) {
                sum = [b.label, b.upn || b.id].filter(Boolean).join(' · ');
            } else if (mode === 'create') {
                sum = 'Noch nicht vorgemerkt – Assistent erzeugen.';
            }
        } else if (mode === 'create' && !b.id) {
            sum = 'Noch nicht angelegt – Name/Alias prüfen und anlegen.';
        } else if (mode !== 'none' && b.id) {
            sum = [b.label, b.mail, b.id].filter(Boolean).join(' · ');
        }
        const used = mode !== 'none';
        /** Aktion nur match|create – „nicht nutzen“ steuert die Checkbox */
        const action = mode === 'create' ? 'create' : 'match';
        function opt(val, label, disabled) {
            return (
                '<option value="' +
                val +
                '"' +
                (action === val ? ' selected' : '') +
                (disabled ? ' disabled' : '') +
                '>' +
                label +
                '</option>'
            );
        }
        const modeOpts = isChat
            ? opt('match', 'Aus KlassenChats') +
              opt('create', 'Neu anlegen', true)
            : isMail
              ? opt('match', 'Bestehendes zuordnen') + opt('create', 'Neu anlegen')
              : opt('match', 'Bestehende zuordnen') + opt('create', 'Neu anlegen');
        const kindLabel = schulOrgM365ObjectKindLabel(kind);
        const chatToolHref = klassenchatsToolHref();
        const actionsHtml = isChat
            ? '<p class="muted schul-org-m365-obj__hint"' +
              (used ? '' : ' hidden') +
              ' data-m365-chat-hint>' +
              '<a class="schul-org-m365-obj__tool-link" href="' +
              escapeHtml(chatToolHref) +
              '" data-m365-klassenchats-link>' +
              '<i class="bi bi-box-arrow-up-right" aria-hidden="true"></i> Werkzeug KlassenChats öffnen</a></p>'
            : '<div class="schul-org-m365-obj__actions" data-m365-match-row' +
              (mode === 'match' ? '' : ' hidden') +
              '>' +
              '<button type="button" class="btn btn-sm schul-org-m365-obj__pick" data-m365-pick>' +
              '<i class="bi bi-search" aria-hidden="true"></i> In Entra wählen</button>' +
              '<button type="button" class="btn btn-sm alt schul-org-m365-obj__clear" data-m365-clear>Lösen</button>' +
              '</div>';
        return (
            '<div class="schul-org-m365-obj' +
            (used ? ' is-active' : '') +
            (isChat ? ' is-chat' : '') +
            '" data-m365-kind="' +
            kind +
            '">' +
            '<label class="schul-org-m365-obj__use">' +
            '<input type="checkbox" class="schul-org-m365-obj__check" data-m365-use' +
            (used ? ' checked' : '') +
            ' aria-label="' +
            escapeHtml(kindLabel) +
            ' nutzen" />' +
            '<span class="schul-org-m365-obj__check-ui" aria-hidden="true"></span>' +
            '</label>' +
            '<span class="schul-org-m365-obj__icon" aria-hidden="true"><i class="bi ' +
            schulOrgM365ObjectKindIcon(kind) +
            '"></i></span>' +
            '<div class="schul-org-m365-obj__meta">' +
            '<strong class="schul-org-m365-obj__name">' +
            escapeHtml(kindLabel) +
            '</strong>' +
            '<span class="muted schul-org-m365-obj__sum" data-m365-sum>' +
            escapeHtml(sum) +
            '</span></div>' +
            '<div class="schul-org-m365-obj__mode-wrap"' +
            (used ? '' : ' hidden') +
            ' data-m365-mode-wrap>' +
            '<select class="schul-org-m365-obj__mode" data-m365-mode aria-label="Aktion ' +
            escapeHtml(kindLabel) +
            '">' +
            modeOpts +
            '</select></div>' +
            actionsHtml +
            '</div>'
        );
    }).join('');
}

/** Relativer Link zum Werkzeug KlassenChats (embedded vs. tools/). */
export function klassenchatsToolHref() {
    try {
        const path = String(
            (typeof location !== 'undefined' && location.pathname) || ''
        );
        if (/\/tools\//i.test(path)) return 'klassenchats.html';
    } catch {
        /* ignore */
    }
    return 'tools/klassenchats.html';
}

export function editFootHtml(delAttr, delLabel, saveAttr) {
    return (
        '<button type="button" class="btn btn-danger schul-org-edit-btn-danger" ' +
        delAttr +
        '><i class="bi bi-trash" aria-hidden="true"></i> ' +
        escapeHtml(delLabel) +
        '</button>' +
        '<span class="schul-org-edit-foot-spacer"></span>' +
        '<button type="button" class="btn schul-org-edit-btn-ghost" data-schul-org-close>Schließen</button>' +
        '<button type="button" class="btn btn-success schul-org-edit-btn-save" ' +
        saveAttr +
        '><i class="bi bi-check-lg" aria-hidden="true"></i> Speichern</button>'
    );
}
