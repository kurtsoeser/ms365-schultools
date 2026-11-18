/**
 * Freistellung Detail-Drawer (Übersicht / Liste / Kalender).
 * Extrahiert aus renderDetailModal; Layout analog AKT-Drawer.
 */
import {
    formatDeDate,
    formatDeDateTime,
    statusLabel,
    canDecide,
    canDeleteFreistellung,
    itemVisibleForRole
} from './freistellung-planer-state.js';

function esc(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function statusBadge(status) {
    const s = String(status || 'Ausstehend').trim();
    const cls =
        s === 'Genehmigt'
            ? 'fr-badge fr-badge--ok'
            : s === 'Abgelehnt'
              ? 'fr-badge fr-badge--bad'
              : 'fr-badge fr-badge--warn';
    return `<span class="${cls}">${esc(statusLabel(s))}</span>`;
}

/**
 * Detail-Drawer für eine Freistellung (state.detailId).
 * @param {object} state
 */
export function renderFreistellungDrawer(state) {
    if (!state || !state.detailId) return '';
    const it = (state.items || []).find((x) => String(x.itemId) === String(state.detailId));
    if (!it || !itemVisibleForRole(state, it, { scopeAll: state.role === 'direktion' })) return '';

    const allowDelete = canDeleteFreistellung(state);
    const allowDecide = canDecide(state) && String(it.status || '').trim() === 'Ausstehend';
    const title = it.schuelerName || it.titel || 'Freistellung';
    const dayCount = it.dayCount;
    const dayLabel =
        dayCount == null ? '–' : String(dayCount) + ' Tag' + (dayCount === 1 ? '' : 'e');

    const primaryActions = [];
    const secondaryActions = [];

    if (allowDecide) {
        primaryActions.push(
            `<button type="button" class="btn fr-btn-primary" data-fr-approve="${esc(
                it.itemId
            )}" title="Status Genehmigt (Fallback)">Genehmigen</button>`
        );
        secondaryActions.push(
            `<button type="button" class="btn" data-fr-reject="${esc(
                it.itemId
            )}" title="Status Abgelehnt (Fallback)">Ablehnen</button>`
        );
    }
    if (allowDelete) {
        secondaryActions.push(
            `<button type="button" class="btn fr-btn-delete" data-fr-delete="${esc(
                it.itemId
            )}"><i class="bi bi-trash"></i>Antrag löschen</button>`
        );
    }
    const nachweiseHtml =
        (it.nachweise || []).length
            ? (it.nachweise || [])
                  .map(
                      (n) =>
                          `<a href="${esc(n.url)}" target="_blank" rel="noopener noreferrer">${esc(
                              n.name
                          )}</a>`
                  )
                  .join('<br>')
            : '–';

    return `
    <div class="fr-drawer-root" role="dialog" aria-modal="true" aria-labelledby="frDetailTitle">
      <button type="button" class="fr-drawer__backdrop" data-fr-close-detail aria-label="Schließen"></button>
      <aside class="fr-drawer">
        <header class="fr-drawer__head">
          <div>
            <h2 class="fr-drawer__title" id="frDetailTitle">${esc(title)}</h2>
            <p class="fr-drawer__sub">${esc(it.klasse || '')}${
                it.kategorie ? ` · ${esc(it.kategorie)}` : ''
            }</p>
            <div class="fr-drawer__badges">${statusBadge(it.status)}</div>
          </div>
          <button type="button" class="btn fr-drawer__close" data-fr-close-detail aria-label="Schließen"><i class="bi bi-x-lg"></i></button>
        </header>

        <div class="fr-drawer__core">
          <p class="fr-drawer__when"><i class="bi bi-calendar3" aria-hidden="true"></i> ${esc(
              formatDeDateTime(it.beginn)
          )} – ${esc(formatDeDateTime(it.ende))} (${esc(dayLabel)})</p>
          <p class="fr-drawer__meta muted">${esc(it.approvalLabel || '')}</p>
        </div>

        <div class="fr-drawer__body">
          <section class="fr-drawer__section">
            <h4>Antrag</h4>
            <dl class="fr-dl">
              <dt>Status</dt><dd>${statusBadge(it.status)}</dd>
              <dt>Klasse</dt><dd>${esc(it.klasse)}</dd>
              <dt>Zeitraum</dt><dd>${esc(formatDeDateTime(it.beginn))} – ${esc(
                  formatDeDateTime(it.ende)
              )} (${esc(dayLabel)})</dd>
              <dt>Genehmigung</dt><dd>${esc(it.approvalLabel)}</dd>
              <dt>Kategorie</dt><dd>${esc(it.kategorie)}</dd>
              <dt>KV</dt><dd>${esc(it.kvName || '–')} &lt;${esc(it.kvEmail || '')}&gt;</dd>
              <dt>Beschreibung</dt><dd>${esc(it.beschreibung || '–')}</dd>
              <dt>Nachweise</dt><dd>${nachweiseHtml}</dd>
              <dt>Bemerkungen</dt><dd>${esc(it.bemerkungen || '–')}</dd>
              <dt>Beantragt von</dt><dd>${esc(it.authorName || '–')}${
                  it.authorEmail ? ` &lt;${esc(it.authorEmail)}&gt;` : ''
              }</dd>
            </dl>
          </section>

          ${
              it.genehmigtVonKv ||
              it.genehmigtAmKv ||
              it.genehmigtVonDirektion ||
              it.genehmigtAmDirektion ||
              it.abgelehntVon ||
              it.abgelehntAm
                  ? `<section class="fr-drawer__section">
            <h4>Audit / Freigaben</h4>
            <dl class="fr-dl">
              ${
                  it.genehmigtVonKv || it.genehmigtAmKv
                      ? `<dt>Freigabe KV</dt><dd>${esc(it.genehmigtVonKv || '–')}${
                            it.genehmigtAmKv ? ' · ' + esc(formatDeDate(it.genehmigtAmKv)) : ''
                        }</dd>`
                      : ''
              }
              ${
                  it.genehmigtVonDirektion || it.genehmigtAmDirektion
                      ? `<dt>Freigabe Direktion</dt><dd>${esc(it.genehmigtVonDirektion || '–')}${
                            it.genehmigtAmDirektion
                                ? ' · ' + esc(formatDeDate(it.genehmigtAmDirektion))
                                : ''
                        }</dd>`
                      : ''
              }
              ${
                  it.abgelehntVon || it.abgelehntAm
                      ? `<dt>Abgelehnt</dt><dd>${esc(it.abgelehntVon || '–')}${
                            it.abgelehntAm ? ' · ' + esc(formatDeDate(it.abgelehntAm)) : ''
                        }</dd>`
                      : ''
              }
            </dl>
          </section>`
                  : ''
          }

          ${
              allowDecide
                  ? `<section class="fr-drawer__section fr-drawer__section--action">
            <h4>Entscheidung</h4>
            <p class="fr-drawer__lead">Manuelles ✓/✗ nur Fallback ohne Microsoft Approvals Flow.</p>
          </section>`
                  : ''
          }
        </div>

        <footer class="fr-drawer__foot">
          ${primaryActions.join('')}
          ${secondaryActions.join('')}
        </footer>
      </aside>
    </div>`;
}

/** Alias für schrittweisen Umbau aus freistellung-planer-ui.js */
export { renderFreistellungDrawer as renderDetailModal };
