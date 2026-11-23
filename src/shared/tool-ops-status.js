/**
 * Eine rollengerechte Statuskarte für Alltag-Tools (Elternbrief, Einsammler, Krankmeldung).
 * Ersetzt gestapelte Login-/SharePoint-/Empty-Hinweise.
 */

/**
 * @typedef {'need_login' | 'site_missing' | 'spo_local' | 'scope' | ''} ToolOpsKind
 *
 * @typedef {object} ToolOpsSync
 * @property {string} [source]
 * @property {boolean} [ok]
 * @property {string} [siteUrl]
 * @property {string} [error]
 * @property {boolean} [connected]
 *
 * @typedef {object} ToolOpsInput
 * @property {string} nounPlural
 * @property {string} [listWord]
 * @property {boolean} [requireLogin]
 * @property {boolean} [loggedIn]
 * @property {string} [accountLabel]
 * @property {boolean} [isAdmin]
 * @property {boolean} [showAll]
 * @property {string} [roleLabelLehrer]
 * @property {string} [roleLabelAdmin]
 * @property {string} [scopeOwnLabel]
 * @property {string} [scopeAllLabel]
 * @property {string} [scopeToggleId]
 * @property {string} [scopeToggleLabel]
 * @property {ToolOpsSync} [sync]
 * @property {string} [setupHref]
 * @property {string} [setupLabel]
 * @property {(s: string) => string} [esc]
 */

function defaultEsc(s) {
    return String(s == null ? '' : s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

/**
 * Priorität: Login → fehlende Site → Lokal/Demo → Scope (wenn Login-Tool) → leer.
 * @param {ToolOpsInput} input
 * @returns {ToolOpsKind}
 */
export function resolveToolOpsKind(input) {
    const sync = (input && input.sync) || {};
    const requireLogin = input && input.requireLogin !== false;
    const loggedIn = !!(input && input.loggedIn);
    const source = String(sync.source || '');
    const liveOk = source === 'live' && !!sync.ok;
    const siteUrl = String(sync.siteUrl || '').trim();

    if (requireLogin && !loggedIn) return 'need_login';
    if (!liveOk) {
        if (!siteUrl && source !== 'demo') return 'site_missing';
        if (source === 'demo' || source === 'local' || source === 'none' || sync.error) {
            return 'spo_local';
        }
    }
    if (requireLogin && loggedIn) return 'scope';
    return '';
}

/**
 * @param {ToolOpsInput} input
 * @returns {string}
 */
export function renderToolOpsStatusCard(input) {
    const esc = (input && input.esc) || defaultEsc;
    const kind = resolveToolOpsKind(input || {});
    if (!kind) return '';

    const noun = esc((input && input.nounPlural) || 'Einträge');
    const listWord = esc((input && input.listWord) || 'Listen');
    const isAdmin = !!(input && input.isAdmin);
    const sync = (input && input.sync) || {};
    const setupHref = String((input && input.setupHref) || '').trim();
    const setupLabel = esc((input && input.setupLabel) || (listWord === 'Liste' ? 'Liste einrichten' : 'Listen einrichten'));

    if (kind === 'need_login') {
        return (
            '<section class="ms365-tool-ops ms365-tool-ops--login" role="status">' +
            '<div class="ms365-tool-ops__text">' +
            '<p class="ms365-tool-ops__eyebrow">Anmeldung</p>' +
            '<p>Bitte mit Microsoft-Konto anmelden – dann sehen Sie nur <strong>Ihre</strong> ' +
            noun +
            '.</p>' +
            '</div></section>'
        );
    }

    if (kind === 'site_missing') {
        const cta =
            isAdmin && setupHref
                ? '<p class="ms365-tool-ops__actions"><a class="ms365-tool-ops__link" href="' +
                  esc(setupHref) +
                  '">' +
                  setupLabel +
                  '</a></p>'
                : '';
        return (
            '<section class="ms365-tool-ops ms365-tool-ops--ops" role="status">' +
            '<div class="ms365-tool-ops__text">' +
            '<p class="ms365-tool-ops__eyebrow">SharePoint</p>' +
            '<p>Noch keine Site – IT richtet die ' +
            listWord +
            ' ein. Bis dahin lokal/Demo.</p>' +
            cta +
            '</div>' +
            renderScopeRow(input, esc) +
            '</section>'
        );
    }

    if (kind === 'spo_local') {
        const body = sync.error
            ? esc(sync.error)
            : 'SharePoint nicht verbunden – ' +
              noun +
              ' bleiben lokal, bis die ' +
              listWord +
              ' greifen.';
        const cta =
            isAdmin && setupHref
                ? '<p class="ms365-tool-ops__actions"><a class="ms365-tool-ops__link" href="' +
                  esc(setupHref) +
                  '">' +
                  setupLabel +
                  '</a></p>'
                : '';
        return (
            '<section class="ms365-tool-ops ms365-tool-ops--ops" role="status">' +
            '<div class="ms365-tool-ops__text">' +
            '<p class="ms365-tool-ops__eyebrow">Lokal / Demo</p>' +
            '<p>' +
            body +
            '</p>' +
            cta +
            '</div>' +
            renderScopeRow(input, esc) +
            '</section>'
        );
    }

    // scope
    const scope = renderScopeRow(input, esc);
    if (!scope) return '';
    return (
        '<section class="ms365-tool-ops ms365-tool-ops--scope" role="status">' + scope + '</section>'
    );
}

/**
 * @param {ToolOpsInput} input
 * @param {(s: string) => string} esc
 */
function renderScopeRow(input, esc) {
    if (!(input && input.requireLogin !== false && input.loggedIn)) return '';
    const who = esc(input.accountLabel || '');
    const isAdmin = !!input.isAdmin;
    const role = esc(
        isAdmin
            ? input.roleLabelAdmin || 'Admin'
            : input.roleLabelLehrer || 'Lehrkraft'
    );
    const showAll = !!input.showAll;
    const own = esc(input.scopeOwnLabel || 'nur eigene');
    const all = esc(input.scopeAllLabel || 'alle');
    const toggleId = String(input.scopeToggleId || '').trim();
    const toggleLabel = esc(input.scopeToggleLabel || 'Alle anzeigen (Admin)');
    return (
        '<div class="ms365-tool-ops__scope">' +
        '<span>Angemeldet: <strong>' +
        who +
        '</strong> · ' +
        role +
        (showAll && isAdmin ? ' · <em>' + all + '</em>' : ' · ' + own) +
        '</span>' +
        (isAdmin && toggleId
            ? '<label class="ms365-tool-ops__toggle"><input type="checkbox" id="' +
              esc(toggleId) +
              '"' +
              (showAll ? ' checked' : '') +
              '> ' +
              toggleLabel +
              '</label>'
            : '') +
        '</div>'
    );
}
