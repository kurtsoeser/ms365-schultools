/**
 * Dashboard-Footer „Ihr Stand“ – Fortschrittsringe und Metrik-Karten.
 */

function metricFillColor(tone) {
    if (tone === 'ok') return '#0d8050';
    if (tone === 'warn') return '#e67700';
    return '#8b93a7';
}

/**
 * @param {HTMLElement | null} el
 * @param {{
 *   title?: string,
 *   done?: number,
 *   total?: number,
 *   valueText?: string,
 *   hint?: string,
 *   tone?: string,
 *   jumpTaskId?: string
 * }} [options]
 */
export function updateStandMetricCard(el, options) {
    if (!el) return;
    const o = options && typeof options === 'object' ? options : {};
    const done = Math.max(0, Number(o.done) || 0);
    const total = Math.max(0, Number(o.total) || 0);
    const tone = String(o.tone || 'muted').trim() || 'muted';
    el.setAttribute('data-tone', tone);

    let pct = 0;
    if (total > 0) pct = Math.min(100, Math.round((done / total) * 100));
    else if (tone === 'ok') pct = 100;

    const titleEl = el.querySelector('.dash-stand__metric-title');
    const valueEl = el.querySelector('.dash-stand__metric-value');
    const hintEl = el.querySelector('.dash-stand__metric-hint');
    const ringInner = el.querySelector('.dash-stand__metric-ring-inner');
    const ring = el.querySelector('.dash-stand__metric-ring');
    const jump = el.querySelector('.dash-stand__metric-jump');

    if (titleEl && o.title) titleEl.textContent = o.title;
    if (valueEl) {
        valueEl.textContent = total > 0 ? done + ' von ' + total : String(o.valueText || '–');
    }
    if (hintEl) hintEl.textContent = String(o.hint || '');

    if (ring) {
        ring.style.setProperty('--dash-metric-pct', pct + '%');
        ring.style.setProperty('--dash-metric-fill', metricFillColor(tone));
    }
    if (ringInner) {
        ringInner.textContent = total > 0 ? pct + '%' : tone === 'ok' ? '✓' : '–';
    }

    if (jump) {
        const taskId = String(o.jumpTaskId || '').trim();
        if (taskId) {
            jump.hidden = false;
            jump.href = '#' + taskId;
            if (!jump.dataset.bound) {
                jump.dataset.bound = '1';
                jump.addEventListener('click', function (ev) {
                    ev.preventDefault();
                    if (typeof window.__ms365DashScrollToTask === 'function') {
                        window.__ms365DashScrollToTask(taskId);
                    }
                });
            }
        } else {
            jump.hidden = true;
        }
    }
}

/**
 * @param {HTMLElement | null} el
 * @param {{ ok: number, warn: number, label?: string }} summary
 */
export function updateStandOverallBanner(el, summary) {
    if (!el) return;
    const ok = Math.max(0, Number(summary && summary.ok) || 0);
    const warn = Math.max(0, Number(summary && summary.warn) || 0);
    const total = ok + warn;
    if (!total) {
        el.hidden = true;
        el.textContent = '';
        return;
    }
    el.hidden = false;
    if (warn === 0) {
        el.setAttribute('data-tone', 'ok');
        el.textContent = ok + ' Bereiche erledigt – Register und M365 im Gleichschritt.';
    } else {
        el.setAttribute('data-tone', 'warn');
        el.textContent = warn + ' offen · ' + ok + ' erledigt – Details in den Kacheln oben.';
    }
}
