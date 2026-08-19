/**
 * DOM-Aufbau für den Mitglieder-Abgleich (Richtung Lokal ↔ Cloud).
 * @file
 */
import { memberDisplayNameFromGraph } from './membership-reconcile.js';

const SECTION_META = {
    onlyLocal: {
        variant: 'local',
        direction: 'Lokal → Cloud'
    },
    onlyGraph: {
        variant: 'cloud',
        direction: 'Cloud → Lokal'
    }
};

const ACTION_BUTTON_IDS = {
    onlyLocal: ['slgMrAddToGroup'],
    onlyGraph: ['slgMrImportLocal', 'slgMrRemoveFromGroup']
};

/**
 * Buttons vor body.replaceChildren() zurück in den Staging-Container legen.
 * @param {HTMLElement|null} stagingEl
 */
export function collectMembershipReviewActionButtons(stagingEl) {
    if (!stagingEl) return;
    ['slgMrAddToGroup', 'slgMrImportLocal', 'slgMrRemoveFromGroup'].forEach(function (id) {
        const btn = document.getElementById(id);
        if (btn && btn.parentElement !== stagingEl) stagingEl.appendChild(btn);
    });
}

/**
 * Aktionsbuttons unter die jeweilige Auswahl-Spalte setzen.
 * @param {HTMLElement} body
 * @param {object} [counts]
 */
export function attachMembershipReviewSectionActions(body, counts) {
    if (!body) return;
    const c = counts && typeof counts === 'object' ? counts : {};

    function mountFooter(sectionKey, buttonIds, count) {
        const section = body.querySelector('[data-mr-section-panel="' + sectionKey + '"]');
        if (!section) return;
        let footer = section.querySelector('.mr-section-actions');
        if (!footer) {
            footer = document.createElement('div');
            footer.className = 'mr-section-actions';
            section.appendChild(footer);
        }
        footer.replaceChildren();
        buttonIds.forEach(function (id) {
            const btn = document.getElementById(id);
            if (!btn) return;
            btn.disabled = !count;
            footer.appendChild(btn);
        });
    }

    mountFooter('onlyLocal', ACTION_BUTTON_IDS.onlyLocal, c.onlyLocal || 0);
    mountFooter('onlyGraph', ACTION_BUTTON_IDS.onlyGraph, c.onlyGraph || 0);
}

export function renderMembershipReviewLegend() {
    const legend = document.createElement('div');
    legend.className = 'mr-flow-legend';
    legend.innerHTML =
        '<p class="mr-flow-legend__intro">Es werden zwei Quellen verglichen – Abweichungen können in eine Richtung übernommen werden:</p>' +
        '<div class="mr-flow-legend__sources">' +
        '<div class="mr-source mr-source--local">' +
        '<i class="bi bi-journal-text" aria-hidden="true"></i>' +
        '<div><strong>Schul-Liste</strong><span>lokal (Stammdaten im Browser)</span></div>' +
        '</div>' +
        '<div class="mr-flow-legend__mid" aria-hidden="true"><i class="bi bi-arrow-left-right"></i></div>' +
        '<div class="mr-source mr-source--cloud">' +
        '<i class="bi bi-cloud" aria-hidden="true"></i>' +
        '<div><strong>Microsoft-365-Gruppe</strong><span>online in der Cloud</span></div>' +
        '</div>' +
        '</div>';
    return legend;
}

/**
 * @param {number} listCount
 * @param {object} diff
 * @param {number} [groupCountOverride]
 */
export function renderMembershipReviewStats(listCount, diff, groupCountOverride) {
    const groupCount =
        typeof groupCountOverride === 'number'
            ? groupCountOverride
            : diff.onlyLocal.length + diff.onlyGraph.length + diff.both.length + (diff.preserved ? diff.preserved.length : 0);

    const bar = document.createElement('div');
    bar.className = 'mr-stats-bar';
    bar.innerHTML =
        '<div class="mr-stat mr-stat--local"><span class="mr-stat__n">' +
        listCount +
        '</span><span class="mr-stat__l">Schul-Liste</span></div>' +
        '<div class="mr-stat mr-stat--cloud"><span class="mr-stat__n">' +
        groupCount +
        '</span><span class="mr-stat__l">M365-Gruppe</span></div>' +
        '<div class="mr-stat mr-stat--ok"><span class="mr-stat__n">' +
        diff.both.length +
        '</span><span class="mr-stat__l">Übereinstimmend</span></div>' +
        '<div class="mr-stat mr-stat--diff"><span class="mr-stat__n">' +
        (diff.onlyLocal.length + diff.onlyGraph.length) +
        '</span><span class="mr-stat__l">Abweichungen</span></div>';
    return bar;
}

function renderReviewTable(sectionKey, emails, graphByEmail, opts) {
    const o = opts && typeof opts === 'object' ? opts : {};
    const wrap = document.createElement('div');
    wrap.className = 'slg-deviation-table-wrap';
    if (!emails.length) {
        const p = document.createElement('p');
        p.className = 'muted slg-deviation-list__empty';
        p.style.margin = '0';
        p.textContent = o.emptyText || 'Keine';
        wrap.appendChild(p);
        return wrap;
    }
    const table = document.createElement('table');
    table.className = 'slg-deviation-table';
    const thead = document.createElement('thead');
    const headRow = document.createElement('tr');
    ['', 'E-Mail', o.nameColumn ? 'Name (Graph)' : ''].forEach(function (label, idx) {
        if (idx === 2 && !o.nameColumn) return;
        const th = document.createElement('th');
        th.className = idx === 0 ? 'col-check' : '';
        th.textContent = label;
        headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    table.appendChild(thead);
    const tbody = document.createElement('tbody');
    emails.forEach(function (em) {
        const tr = document.createElement('tr');
        const tdCheck = document.createElement('td');
        tdCheck.className = 'col-check';
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = o.checkedByDefault !== false;
        cb.setAttribute('data-mr-section', sectionKey);
        cb.setAttribute('data-mr-email', em);
        tdCheck.appendChild(cb);
        tr.appendChild(tdCheck);
        const tdEm = document.createElement('td');
        tdEm.textContent = em;
        tr.appendChild(tdEm);
        if (o.nameColumn) {
            const tdName = document.createElement('td');
            const person = graphByEmail.get(em);
            tdName.textContent = person ? memberDisplayNameFromGraph(person, em) : '–';
            tr.appendChild(tdName);
        }
        tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    wrap.appendChild(table);
    return wrap;
}

/**
 * @param {string} sectionKey
 * @param {string} title
 * @param {string} hint
 * @param {string[]} emails
 * @param {Map<string, object>} graphByEmail
 * @param {object} [opts]
 */
export function renderMembershipReviewSection(sectionKey, title, hint, emails, graphByEmail, opts) {
    const o = opts && typeof opts === 'object' ? opts : {};
    const meta = SECTION_META[sectionKey] || {};
    const section = document.createElement('section');
    section.className =
        'slg-deviation-section mr-diff-section mr-diff-section--' + (meta.variant || 'neutral');
    section.setAttribute('data-mr-section-panel', sectionKey);

    const head = document.createElement('div');
    head.className = 'mr-diff-section__head';
    const h = document.createElement('h4');
    h.className = 'slg-deviation-section__title';
    if (meta.direction) {
        const badge = document.createElement('span');
        badge.className = 'mr-direction-badge mr-direction-badge--' + meta.variant;
        badge.textContent = meta.direction;
        h.appendChild(badge);
    }
    const titleSpan = document.createElement('span');
    titleSpan.textContent = title + ' (' + emails.length + ')';
    h.appendChild(titleSpan);
    head.appendChild(h);
    section.appendChild(head);

    if (hint) {
        const p = document.createElement('p');
        p.className = 'muted slg-deviation-section__hint';
        p.textContent = hint;
        section.appendChild(p);
    }
    if (emails.length) {
        const toolbar = document.createElement('div');
        toolbar.className = 'slg-deviation-section__toolbar';
        const lbl = document.createElement('label');
        lbl.className = 'slg-deviation-section__select-all';
        const allCb = document.createElement('input');
        allCb.type = 'checkbox';
        allCb.checked = true;
        const span = document.createElement('span');
        span.textContent = 'Alle auswählen';
        lbl.appendChild(allCb);
        lbl.appendChild(span);
        toolbar.appendChild(lbl);
        section.appendChild(toolbar);
        allCb.addEventListener('change', function () {
            section.querySelectorAll('input[data-mr-section="' + sectionKey + '"]').forEach(function (cb) {
                cb.checked = allCb.checked;
            });
        });
    }
    section.appendChild(renderReviewTable(sectionKey, emails, graphByEmail, o));
    return section;
}

/**
 * @param {string[]} emails
 * @param {string} [title]
 */
export function renderMembershipReviewBoth(emails, title) {
    const wrap = document.createElement('section');
    wrap.className = 'slg-deviation-section slg-deviation-both mr-diff-section mr-diff-section--ok';
    const details = document.createElement('details');
    const summary = document.createElement('summary');
    summary.textContent = title || 'In Liste und Gruppe (' + emails.length + ') – bereits konsistent';
    details.appendChild(summary);
    if (emails.length) {
        const ul = document.createElement('ul');
        ul.className = 'slg-deviation-list';
        emails.forEach(function (em) {
            const li = document.createElement('li');
            li.textContent = em;
            ul.appendChild(li);
        });
        details.appendChild(ul);
    } else {
        const p = document.createElement('p');
        p.className = 'muted';
        p.style.margin = '8px 0 0';
        p.textContent = 'Keine gemeinsamen Einträge.';
        details.appendChild(p);
    }
    wrap.appendChild(details);
    return wrap;
}

/**
 * @param {object} cfg
 * @returns {DocumentFragment}
 */
export function buildMembershipReviewBody(cfg) {
    const c = cfg && typeof cfg === 'object' ? cfg : {};
    const diff = c.diff;
    const labels = c.labels || {};
    const graphByEmail = c.graphByEmail || new Map();
    const frag = document.createDocumentFragment();

    frag.appendChild(renderMembershipReviewLegend());
    frag.appendChild(renderMembershipReviewStats(c.listCount || 0, diff, c.groupCount));

    const columns = document.createElement('div');
    columns.className = 'mr-diff-columns';
    columns.appendChild(
        renderMembershipReviewSection(
            'onlyLocal',
            labels.onlyLocalTitle || 'Nur in der Schul-Liste',
            labels.onlyLocalHint || 'In den Stammdaten, aber nicht in der Microsoft-365-Gruppe.',
            diff.onlyLocal,
            graphByEmail,
            {
                emptyText: labels.onlyLocalEmpty || 'Keine – alle Listeneinträge sind auch in der Gruppe.',
                nameColumn: false
            }
        )
    );
    columns.appendChild(
        renderMembershipReviewSection(
            'onlyGraph',
            labels.onlyGraphTitle || 'Nur in der Microsoft-365-Gruppe',
            labels.onlyGraphHint ||
                'In der Gruppe online, aber nicht in den lokalen Stammdaten.',
            diff.onlyGraph,
            graphByEmail,
            {
                emptyText: labels.onlyGraphEmpty || 'Keine – alle Gruppenmitglieder stehen auch in der Liste.',
                nameColumn: true
            }
        )
    );
    frag.appendChild(columns);
    frag.appendChild(renderMembershipReviewBoth(diff.both));
    if (diff.preserved && diff.preserved.length) {
        frag.appendChild(
            renderMembershipReviewBoth(
                diff.preserved,
                'In Gruppe, aber weder Klassen-Schüler noch zu entfernende Schüler (' +
                    diff.preserved.length +
                    ') – z. B. Lehrkräfte'
            )
        );
    }
    return frag;
}

const api = {
    collectMembershipReviewActionButtons: collectMembershipReviewActionButtons,
    attachMembershipReviewSectionActions: attachMembershipReviewSectionActions,
    renderMembershipReviewLegend: renderMembershipReviewLegend,
    renderMembershipReviewStats: renderMembershipReviewStats,
    renderMembershipReviewSection: renderMembershipReviewSection,
    renderMembershipReviewBoth: renderMembershipReviewBoth,
    buildMembershipReviewBody: buildMembershipReviewBody
};

if (typeof window !== 'undefined') {
    window.ms365MembershipReviewRender = api;
}

export default api;
