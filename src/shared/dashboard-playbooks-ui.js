/**
 * Dashboard: Playbooks als Fortschritts-Workflows.
 */
import { loadPlaybookState } from './playbook-store.js';
import {
    DASHBOARD_PLAYBOOKS,
    computePlaybookProgress,
    formatPlaybookProgressLabel,
    playbookProgressBarSegments
} from './dashboard-playbooks-catalog.js';

/**
 * @param {import('./dashboard-playbooks-catalog.js').DashboardPlaybookDef} def
 * @param {Storage} [storage]
 */
export function getDashboardPlaybookProgress(def, storage) {
    const st = loadPlaybookState(def.storageKey, storage);
    return computePlaybookProgress(st, def.stepIds);
}

/**
 * @param {import('./dashboard-playbooks-catalog.js').DashboardPlaybookDef} def
 * @param {ReturnType<typeof computePlaybookProgress>} progress
 */
function renderPlaybookRow(def, progress) {
    const label = formatPlaybookProgressLabel(progress);
    const barChars = playbookProgressBarSegments(progress.ratio, Math.min(14, progress.total || 12));
    const pct = progress.total ? Math.round(progress.ratio * 100) : 0;
    const statusClass =
        progress.status === 'complete'
            ? 'dash-playbook-flow__row--complete'
            : progress.status === 'in-progress'
              ? 'dash-playbook-flow__row--active'
              : 'dash-playbook-flow__row--idle';

    const completeBadge =
        progress.status === 'complete' && progress.total > 0
            ? '<span class="dash-playbook-flow__done" aria-hidden="true"><i class="bi bi-check-circle-fill"></i></span>'
            : '';

    return (
        '<a class="dash-playbook-flow__row ' +
        statusClass +
        '" role="listitem" href="' +
        def.href +
        '" data-playbook-id="' +
        def.id +
        '">' +
        '<div class="dash-playbook-flow__head">' +
        '<h3 class="dash-playbook-flow__title"><i class="bi ' +
        def.icon +
        '" aria-hidden="true"></i>' +
        escapeHtml(def.title) +
        '</h3>' +
        (def.blurb ? '<p class="dash-playbook-flow__blurb">' + escapeHtml(def.blurb) + '</p>' : '') +
        '</div>' +
        '<div class="dash-playbook-flow__track" aria-hidden="true">' +
        '<span class="dash-playbook-flow__track-glyphs">' +
        barChars +
        '</span>' +
        '<div class="dash-playbook-flow__bar"><span class="dash-playbook-flow__bar-fill" style="width:' +
        pct +
        '%"></span></div>' +
        '</div>' +
        '<div class="dash-playbook-flow__meta">' +
        '<span class="dash-playbook-flow__status">' +
        escapeHtml(label) +
        '</span>' +
        completeBadge +
        '<span class="dash-playbook-flow__cta">Weiter <i class="bi bi-chevron-right" aria-hidden="true"></i></span>' +
        '</div>' +
        '</a>'
    );
}

/**
 * @param {string} s
 */
function escapeHtml(s) {
    return String(s || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

export function mountDashboardPlaybooksFlows() {
    const wrap = document.getElementById('dashPlaybooksList');
    if (!wrap || wrap.dataset.mounted === '1') return;
    wrap.dataset.mounted = '1';

    function refresh() {
        wrap.innerHTML = DASHBOARD_PLAYBOOKS
            .map(function (def) {
                const progress = getDashboardPlaybookProgress(def);
                return renderPlaybookRow(def, progress);
            })
            .join('');
    }

    refresh();
    window.addEventListener('ms365-app-local-data-changed', function (e) {
        const d = e && e.detail;
        if (!d || d.source === 'playbook') refresh();
    });
    window.addEventListener('storage', refresh);

    return { refresh };
}

if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', mountDashboardPlaybooksFlows);
    } else {
        mountDashboardPlaybooksFlows();
    }
}
