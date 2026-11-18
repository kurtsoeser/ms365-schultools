/**
 * Detail-Drawer für sRDP-Anmeldungen (eine Quelle für Liste/Kalender/Dashboard).
 */
import { buildAnmeldungSummary, teacherRolesOnItem } from './srdp-anmeldung-logic.js';
import { isSrdpStaffRole } from './srdp-anmeldung-state.js';

function esc(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

/**
 * @param {object|null} item
 * @param {object} state
 */
export function renderAnmeldungDrawer(item, state) {
    if (!item) return '';
    const summary = buildAnmeldungSummary(item, state.options && state.options.profileId);
    const roles = teacherRolesOnItem(item, state.teacherLabel);
    const staff = isSrdpStaffRole(state.role);
    const confirmed = !!item.bestaetigung;
    const nacharbeit = !!item.nacharbeitErledigt;
    return `
    <div class="srdp-drawer-backdrop" data-srdp-drawer-close tabindex="-1"></div>
    <aside class="srdp-drawer" role="dialog" aria-modal="true" aria-label="Anmeldungsdetail">
      <header class="srdp-drawer__head">
        <div>
          <p class="srdp-drawer__eyebrow">${esc(item.klasse || '–')} · ${esc(item.variante || '–')}</p>
          <h2>${esc(summary.title || item.title || 'Anmeldung')}</h2>
        </div>
        <button type="button" class="btn" data-srdp-drawer-close aria-label="Schließen"><i class="bi bi-x-lg"></i></button>
      </header>
      <div class="srdp-drawer__body">
        <p class="srdp-panel__hint">${esc(summary.pruefplan || 'Prüfplan folgt der gewählten Variante.')}</p>
        <div class="srdp-drawer__badges">
          <span class="srdp-badge ${confirmed ? 'srdp-badge--ok' : 'srdp-badge--warn'}">${
              confirmed ? 'Bestätigt' : 'Bestätigung offen'
          }</span>
          <span class="srdp-badge ${nacharbeit ? 'srdp-badge--ok' : 'srdp-badge--info'}">${
              nacharbeit ? 'Prüfungsliste erledigt' : 'Nacharbeit offen'
          }</span>
        </div>
        <dl class="srdp-summary" style="border:0;box-shadow:none;margin:0;">
          ${summary.rows
              .map((r) => `<dt>${esc(r.label)}</dt><dd>${esc(r.value)}</dd>`)
              .join('')}
          ${
              item.schuelerEmail
                  ? `<dt>Schüler-E-Mail</dt><dd>${esc(item.schuelerEmail)}</dd>`
                  : ''
          }
        </dl>
        ${
            roles.length
                ? `<p class="muted">Deine Rollen: ${roles
                      .map((r) => esc(r.label))
                      .join(', ')}</p>`
                : ''
        }
      </div>
      <footer class="srdp-drawer__foot">
        ${
            staff
                ? `<button type="button" class="btn ${nacharbeit ? '' : 'btn-success'}" data-srdp-toggle-nacharbeit="${esc(
                      item.itemId
                  )}">
            <i class="bi bi-clipboard-check"></i>${
                nacharbeit ? 'Nacharbeit wieder öffnen' : 'Prüfungsliste erledigt'
            }
          </button>`
                : ''
        }
        <button type="button" class="btn" data-srdp-drawer-close>Schließen</button>
      </footer>
    </aside>`;
}

/**
 * @param {object} item
 * @param {object} state
 * @param {{ heading?: string, sub?: string }} [opts]
 */
export function renderSubmitSuccess(item, state, opts) {
    const o = opts || {};
    const summary = buildAnmeldungSummary(item, state.options && state.options.profileId);
    return `<div class="srdp-alert srdp-alert--ok" role="status">Anmeldung gespeichert.</div>
      <div class="srdp-summary">
        <div class="srdp-summary__head">
          <h3>${esc(o.heading || 'Deine Anmeldung')}</h3>
          <p>${esc(o.sub || summary.pruefplan || 'Deine gewählten Prüfungsteile')}</p>
        </div>
        <dl>
          ${summary.rows.map((r) => `<dt>${esc(r.label)}</dt><dd>${esc(r.value)}</dd>`).join('')}
        </dl>
      </div>
      <div class="srdp-actions">
        <button type="button" class="btn" data-srdp-view-jump="meine"><i class="bi bi-person-check"></i>Zur Anmeldung</button>
        <button type="button" class="btn" data-srdp-dismiss-success>Schließen</button>
      </div>`;
}
