/**
 * Shared Systemadministration shell (Vorbild Schulaktivitäten).
 * Tabs: Datenquelle · Berechtigungen · Kalender · Wartung · Danger Zone
 */

export const SYSTEM_ADMIN_DEFAULT_TABS = /** @type {const} */ ([
    'site',
    'perms',
    'kalender',
    'wartung',
    'danger'
]);

/**
 * @param {string} raw
 * @param {readonly string[]} [allowed]
 * @returns {string}
 */
export function normalizeSystemTab(raw, allowed = SYSTEM_ADMIN_DEFAULT_TABS) {
    const t = String(raw || '')
        .trim()
        .toLowerCase();
    return allowed.includes(t) ? t : 'site';
}

/**
 * @param {{
 *   prefix: string,
 *   activeTab: string,
 *   tabs?: Array<{ id: string, label: string, danger?: boolean }>,
 *   ariaLabel?: string,
 *   dataAttr?: string
 * }} opts
 */
export function renderSystemTabBar(opts) {
    const prefix = String(opts.prefix || 'ms').trim() || 'ms';
    const active = normalizeSystemTab(opts.activeTab);
    const dataAttr = String(opts.dataAttr || `data-${prefix}-sys-tab`);
    const ariaLabel = String(opts.ariaLabel || 'Systemadministration');
    const tabs =
        Array.isArray(opts.tabs) && opts.tabs.length
            ? opts.tabs
            : [
                  { id: 'site', label: '1 · Datenquelle' },
                  { id: 'perms', label: '2 · Berechtigungen' },
                  { id: 'kalender', label: '3 · Kalender' },
                  { id: 'wartung', label: '4 · Wartung' },
                  { id: 'danger', label: 'Danger Zone', danger: true }
              ];
    const btns = tabs
        .map((tab) => {
            const id = String(tab.id || '').trim();
            if (!id) return '';
            const danger = !!tab.danger;
            const selected = active === id;
            return `<button type="button" class="ms-sys-tabs__btn${selected ? ' is-active' : ''}${
                danger ? ' ms-sys-tabs__btn--danger' : ''
            }" role="tab" id="${prefix}SysTab-${id}" ${dataAttr}="${id}" aria-selected="${
                selected ? 'true' : 'false'
            }" aria-controls="${prefix}SysPanel-${id}">${escapeSys(tab.label)}</button>`;
        })
        .join('');
    return `<div class="ms-sys-tabs" role="tablist" aria-label="${escapeSys(ariaLabel)}">${btns}</div>`;
}

/**
 * @param {string} activeTab
 * @param {string} panelId
 */
export function systemPanelHiddenAttr(activeTab, panelId) {
    return normalizeSystemTab(activeTab) === panelId ? '' : ' hidden';
}

const DEFAULT_TAB_LABEL = {
    site: '1 · Datenquelle',
    perms: '2 · Berechtigungen',
    kalender: '3 · Kalender',
    wartung: '4 · Wartung',
    danger: 'Danger Zone'
};

/**
 * Komplette Systemadministration: Tab-Leiste + eine Karte pro Panel.
 * `body` ist vertrauenswürdiges Markup des Tools (bestehende Felder, gleiche IDs).
 * `leadHtml` ebenfalls (bereits escapet oder statisch).
 *
 * @param {{
 *   prefix: string,
 *   activeTab?: string,
 *   ariaLabel?: string,
 *   dataAttr?: string,
 *   panels: Array<{
 *     id: string,
 *     title: string,
 *     step?: string,
 *     tabLabel?: string,
 *     leadHtml?: string,
 *     body?: string,
 *     danger?: boolean
 *   }>
 * }} opts
 */
export function renderSystemAdminPage(opts) {
    const prefix = String(opts.prefix || 'ms').trim() || 'ms';
    const dataAttr = String(opts.dataAttr || `data-${prefix}-sys-tab`);
    const panels = Array.isArray(opts.panels) ? opts.panels : [];
    const tabs = panels
        .map((p) => {
            const id = String(p.id || '').trim();
            if (!id) return null;
            const danger = !!p.danger || id === 'danger';
            return {
                id,
                danger,
                label:
                    p.tabLabel ||
                    DEFAULT_TAB_LABEL[id] ||
                    (danger ? 'Danger Zone' : String(p.title || id))
            };
        })
        .filter(Boolean);
    const allowed = tabs.map((t) => t.id);
    const active = normalizeSystemTab(opts.activeTab, allowed.length ? allowed : undefined);
    const bar = renderSystemTabBar({
        prefix,
        activeTab: active,
        tabs,
        dataAttr,
        ariaLabel: opts.ariaLabel
    });
    const cards = panels
        .map((p) => {
            const id = String(p.id || '').trim();
            if (!id) return '';
            const danger = !!p.danger || id === 'danger';
            const hidden = active === id ? '' : ' hidden';
            const stepInner = danger
                ? '<i class="bi bi-exclamation-triangle"></i>'
                : escapeSys(p.step || '');
            const lead = p.leadHtml ? `<p>${p.leadHtml}</p>` : '';
            return `<section class="ms-sys-card${danger ? ' ms-sys-card--danger' : ''}" id="${prefix}SysPanel-${id}" role="tabpanel" aria-labelledby="${prefix}SysTab-${id}"${hidden}>
        <header class="ms-sys-card__head">
          <span class="ms-sys-card__step${danger ? ' ms-sys-card__step--danger' : ''}"${
              danger ? ' aria-hidden="true"' : ''
          }>${stepInner}</span>
          <div>
            <h3>${escapeSys(p.title || '')}</h3>
            ${lead}
          </div>
        </header>
        ${p.body || ''}
      </section>`;
        })
        .join('');
    return `<div class="ms-sys">${bar}${cards}</div>`;
}

/**
 * @param {ParentNode|null|undefined} root
 * @param {{ dataAttr: string, onSelect: (tabId: string) => void, allowed?: readonly string[] }} opts
 */
export function bindSystemTabClicks(root, opts) {
    if (!root || !opts || typeof opts.onSelect !== 'function') return;
    const attr = String(opts.dataAttr || '').trim();
    if (!attr) return;
    root.querySelectorAll('[' + attr + ']').forEach((btn) => {
        btn.addEventListener('click', () => {
            opts.onSelect(normalizeSystemTab(btn.getAttribute(attr), opts.allowed));
        });
    });
}

/**
 * @param {string} s
 */
function escapeSys(s) {
    return String(s == null ? '' : s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}
