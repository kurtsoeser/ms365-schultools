/**
 * Tab „Schulstruktur“ – Stammdaten-/Organigramm-Zuordnungen einer Person.
 */
import {
    loadPersonSchulstrukturSettings,
    collectPersonSchulstrukturAssignments
} from './personen-verwaltung-schulstruktur-logic.js';

/** @type {any} */
let host = null;

export function init(a) {
    host = a;
}

function personEmail(u) {
    if (!u) return '';
    return String(u.mail || u.userPrincipalName || '')
        .trim()
        .toLowerCase();
}

/**
 * @param {object|null|undefined} u Graph-User
 */
export function renderSchulstrukturTab(u) {
    const root = document.getElementById('pvSchulstrukturRoot');
    const meta = document.getElementById('pvSchulstrukturMeta');
    if (!root) return;

    root.replaceChildren();
    if (!u) {
        if (meta) meta.textContent = '';
        const p = document.createElement('p');
        p.className = 'muted';
        p.style.margin = '0';
        p.textContent = 'Keine Person ausgewählt.';
        root.appendChild(p);
        return;
    }

    const email = personEmail(u);
    const settings = loadPersonSchulstrukturSettings();
    const data = collectPersonSchulstrukturAssignments(settings, email, {
        name: u.displayName || '',
        upn: u.userPrincipalName || ''
    });

    if (meta) {
        meta.textContent = data.total
            ? data.total +
              ' Zuordnung' +
              (data.total === 1 ? '' : 'en') +
              ' in den Stammdaten / im Schul-Organigramm'
            : 'Keine Stammdaten-Zuordnungen für diese Person gefunden';
    }

    if (data.empty) {
        const empty = document.createElement('div');
        empty.className = 'pv-schulstruktur-empty';
        empty.innerHTML =
            '<p style="margin:0 0 8px;"><strong>Noch keine Zuordnungen</strong></p>' +
            '<p class="muted" style="margin:0;font-size:0.9em;line-height:1.45;">' +
            'Hier erscheinen Funktionen, Verwaltungsbereiche, Leitungen (ARGE/KV), ' +
            'Lehrer-/Schülerstatus und Unterrichtsteams aus den Tenant-Stammdaten – ' +
            'unabhängig von Microsoft-365-Gruppenmitgliedschaften (Tab „Gruppen“).' +
            '</p>';
        root.appendChild(empty);
        return;
    }

    data.categories.forEach(function (cat) {
        const section = document.createElement('section');
        section.className = 'pv-schulstruktur-cat';
        section.setAttribute('aria-label', cat.title);

        const h = document.createElement('h4');
        h.className = 'pv-schulstruktur-cat__title';
        h.textContent = cat.title + ' (' + cat.items.length + ')';
        section.appendChild(h);

        const wrap = document.createElement('div');
        wrap.className = 'pv-groups-wrap pv-schulstruktur-table-wrap';
        const table = document.createElement('table');
        table.setAttribute('aria-label', cat.title);
        const thead = document.createElement('thead');
        thead.innerHTML = '<tr><th>Zuordnung</th><th>Art / Quelle</th></tr>';
        table.appendChild(thead);
        const tbody = document.createElement('tbody');
        cat.items.forEach(function (item) {
            const tr = document.createElement('tr');
            const tdL = document.createElement('td');
            tdL.textContent = item.label || '–';
            const tdD = document.createElement('td');
            tdD.className = 'pv-schulstruktur-detail';
            tdD.textContent = item.detail || item.kind || '–';
            tr.appendChild(tdL);
            tr.appendChild(tdD);
            tbody.appendChild(tr);
        });
        table.appendChild(tbody);
        wrap.appendChild(table);
        section.appendChild(wrap);
        root.appendChild(section);
    });
}

export function reloadSchulstrukturTab() {
    const u = host && typeof host.getSelectedUser === 'function' ? host.getSelectedUser() : null;
    renderSchulstrukturTab(u || null);
}

export default {
    init,
    renderSchulstrukturTab,
    reloadSchulstrukturTab
};
