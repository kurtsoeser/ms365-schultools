/**
 * Lehrer-Freistellung Detail-Drawer (Übersicht / Liste / Kalender / Freigabe).
 */
import { canDecide } from './lfr-state.js';
import {
    formatDeDateTime,
    formatDeDateTimeRange,
    formatDisplayTextWithDeDates,
    normalizeStatus
} from './lfr-logic.js';

function esc(s) {
    return String(s == null ? '' : s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function statusBadge(status) {
    const s = normalizeStatus(status);
    const cls =
        s === 'Genehmigt' ? 'fr-badge--ok' : s === 'Abgelehnt' ? 'fr-badge--bad' : 'fr-badge--warn';
    return `<span class="fr-badge ${cls}">${esc(s)}</span>`;
}

function dayCount(beginn, ende) {
    const b = beginn ? new Date(beginn) : null;
    const e = ende ? new Date(ende) : null;
    if (!b || !e || isNaN(b) || isNaN(e)) return null;
    return Math.max(1, Math.round((e.getTime() - b.getTime()) / 86400000) + 1);
}

function findItem(state) {
    const k = String((state && state.detailId) || '');
    if (!k) return null;
    return (
        (state.items || []).find((it) => String(it.antragId) === k || String(it.itemId) === k) ||
        null
    );
}

/**
 * @param {object} state
 */
export function renderLfrDrawer(state) {
    if (!state || !state.detailId) return '';
    const it = findItem(state);
    if (!it) return '';

    const id = it.antragId || it.itemId;
    const st = normalizeStatus(it.status);
    const allowDecide = canDecide(state) && st === 'Ausstehend';
    const days = dayCount(it.beginn, it.ende);
    const dayLabel =
        days == null ? '–' : String(days) + ' Tag' + (days === 1 ? '' : 'e');
    const title = formatDisplayTextWithDeDates(it.titel || it.lehrerName || 'Freistellung');
    const bemerkungDir = formatDisplayTextWithDeDates(it.bemerkungDirektion || '');

    const primaryActions = [];
    const secondaryActions = [];
    if (allowDecide) {
        primaryActions.push(
            `<button type="button" class="btn fr-btn-primary" data-lfr-approve="${esc(
                id
            )}">Genehmigen</button>`
        );
        secondaryActions.push(
            `<button type="button" class="btn" data-lfr-reject="${esc(id)}">Ablehnen</button>`
        );
    }

    return `
    <div class="fr-drawer-root" role="dialog" aria-modal="true" aria-labelledby="lfrDetailTitle">
      <button type="button" class="fr-drawer__backdrop" data-lfr-close-detail aria-label="Schließen"></button>
      <aside class="fr-drawer">
        <header class="fr-drawer__head">
          <div>
            <h2 class="fr-drawer__title" id="lfrDetailTitle">${esc(title)}</h2>
            <p class="fr-drawer__sub">${esc(it.lehrerName || it.lehrerEmail || '')}${
                it.kategorie ? ` · ${esc(it.kategorie)}` : ''
            }</p>
            <div class="fr-drawer__badges">${statusBadge(it.status)}</div>
          </div>
          <button type="button" class="btn fr-drawer__close" data-lfr-close-detail aria-label="Schließen"><i class="bi bi-x-lg"></i></button>
        </header>

        <div class="fr-drawer__core">
          <p class="fr-drawer__when"><i class="bi bi-calendar3" aria-hidden="true"></i> ${esc(
              formatDeDateTimeRange(it.beginn, it.ende)
          )} (${esc(dayLabel)})</p>
          <p class="fr-drawer__meta muted">Genehmigung über Microsoft Approvals (Power Automate).</p>
        </div>

        <div class="fr-drawer__body">
          <section class="fr-drawer__section">
            <h4>Antrag</h4>
            <dl class="fr-dl">
              <dt>Status</dt><dd>${statusBadge(it.status)}</dd>
              <dt>Lehrkraft</dt><dd>${esc(it.lehrerName || '–')}${
                  it.lehrerEmail ? ` &lt;${esc(it.lehrerEmail)}&gt;` : ''
              }</dd>
              <dt>Zeitraum</dt><dd>${esc(formatDeDateTimeRange(it.beginn, it.ende))} (${esc(
                  dayLabel
              )})</dd>
              <dt>Kategorie</dt><dd>${esc(it.kategorie || '–')}</dd>
              <dt>Beschreibung</dt><dd>${esc(it.beschreibung || '–')}</dd>
              <dt>Bemerkung Direktion</dt><dd>${esc(bemerkungDir || '–')}</dd>
            </dl>
          </section>

          ${
              it.genehmigtVon ||
              it.genehmigtAm ||
              it.abgelehntVon ||
              it.abgelehntAm
                  ? `<section class="fr-drawer__section">
            <h4>Audit / Freigaben</h4>
            <dl class="fr-dl">
              ${
                  it.genehmigtVon || it.genehmigtAm
                      ? `<dt>Freigabe Direktion</dt><dd>${esc(it.genehmigtVon || '–')}${
                            it.genehmigtAm ? ' · ' + esc(formatDeDateTime(it.genehmigtAm)) : ''
                        }</dd>`
                      : ''
              }
              ${
                  it.abgelehntVon || it.abgelehntAm
                      ? `<dt>Abgelehnt</dt><dd>${esc(it.abgelehntVon || '–')}${
                            it.abgelehntAm ? ' · ' + esc(formatDeDateTime(it.abgelehntAm)) : ''
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
