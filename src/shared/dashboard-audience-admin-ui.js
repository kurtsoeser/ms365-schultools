/**
 * Admin-Matrix: welches Dashboard-Werkzeug für Lehrkraft / Schüler sichtbar ist.
 */
import {
    listDashboardToolsByCluster,
    toolLabel,
    DASHBOARD_TOOL_RULES
} from './dashboard-audience-catalog.js';
import {
    getToolAccessFlags,
    loadDashboardToolAccessConfig,
    resetDashboardToolAccessToDefaults,
    setToolAccessFlags
} from './dashboard-audience-store.js';

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function plannerBadge(toolId) {
    const p = (DASHBOARD_TOOL_RULES[toolId] || {}).planner;
    if (!p) return '';
    const name = p.app === 'schularbeiten' ? 'Schularbeiten' : 'Freistellungen';
    return `<span class="dash-access-planner" title="Zusätzlich Planer-Berechtigung nötig">${escapeHtml(name)}</span>`;
}

function renderToolRow(id, cfg) {
    const flags = getToolAccessFlags(id, cfg);
    return `
            <tr data-dash-access-row="${escapeHtml(id)}">
                <th scope="row">${escapeHtml(toolLabel(id))}${plannerBadge(id)}</th>
                <td class="dash-access-matrix__cell">
                    <input type="checkbox" data-dash-access-lehrer="${escapeHtml(id)}" ${flags.lehrer ? 'checked' : ''} aria-label="Lehrkraft: ${escapeHtml(toolLabel(id))}" />
                </td>
                <td class="dash-access-matrix__cell">
                    <input type="checkbox" data-dash-access-schueler="${escapeHtml(id)}" ${flags.schueler ? 'checked' : ''} aria-label="Schüler/in: ${escapeHtml(toolLabel(id))}" />
                </td>
            </tr>`;
}

function renderTable(root) {
    const cfg = loadDashboardToolAccessConfig();
    const clusters = listDashboardToolsByCluster();
    const rows = clusters
        .map((cluster) => {
            const toolRows = cluster.toolIds.map((id) => renderToolRow(id, cfg)).join('');
            return `
            <tr class="dash-access-matrix__cluster">
                <th colspan="3" scope="colgroup">
                    <span class="dash-access-matrix__cluster-label">
                        <i class="bi ${escapeHtml(cluster.icon)}" aria-hidden="true"></i>
                        ${escapeHtml(cluster.label)}
                    </span>
                </th>
            </tr>
            ${toolRows}`;
        })
        .join('');

    root.innerHTML = `
        <section class="dash-access-matrix" aria-labelledby="dashAccessMatrixTitle">
            <div class="dash-access-matrix__head">
                <div>
                    <h3 id="dashAccessMatrixTitle">Dashboard: Werkzeug-Zugriff</h3>
                    <p class="dash-access-matrix__lead">
                        Legt fest, welche Kacheln Mitglieder der <strong>Lehrer-Entra-Gruppe</strong> bzw. <strong>Schüler-Entra-Gruppe</strong> auf dem Start-Dashboard sehen
                        (Gruppen oben auf dieser Seite). Global Administratorinnen/Administratoren sehen immer den vollen Katalog.
                        Schularbeiten- und Freistellungs-Planer haben zusätzlich eigene Berechtigungen im jeweiligen Tool.
                    </p>
                </div>
                <button type="button" class="btn btn-sm" id="dashAccessResetDefaults">Standard wiederherstellen</button>
            </div>
            <div class="dash-access-matrix__scroll">
                <table class="dash-access-matrix__table">
                    <thead>
                        <tr>
                            <th scope="col">Werkzeug</th>
                            <th scope="col">Lehrkraft</th>
                            <th scope="col">Schüler/in</th>
                        </tr>
                    </thead>
                    <tbody>${rows}</tbody>
                </table>
            </div>
            <p class="dash-access-matrix__note">Änderungen werden lokal gespeichert (<code>ms365-dashboard-tool-access-v1</code>) und mit dem Browser-Backup / SharePoint-Sicherung mitgesichert.</p>
        </section>`;

    root.querySelector('#dashAccessResetDefaults')?.addEventListener('click', () => {
        resetDashboardToolAccessToDefaults();
        renderTable(root);
    });

    root.querySelectorAll('[data-dash-access-lehrer]').forEach((el) => {
        el.addEventListener('change', () => {
            const id = el.getAttribute('data-dash-access-lehrer');
            if (!id) return;
            const schEl = root.querySelector(`[data-dash-access-schueler="${id}"]`);
            setToolAccessFlags(id, {
                lehrer: el.checked,
                schueler: schEl ? schEl.checked : false
            });
        });
    });
    root.querySelectorAll('[data-dash-access-schueler]').forEach((el) => {
        el.addEventListener('change', () => {
            const id = el.getAttribute('data-dash-access-schueler');
            if (!id) return;
            const lehrEl = root.querySelector(`[data-dash-access-lehrer="${id}"]`);
            setToolAccessFlags(id, {
                lehrer: lehrEl ? lehrEl.checked : false,
                schueler: el.checked
            });
        });
    });
}

/**
 * @param {string} [mountSelector]
 */
export function mountDashboardAudienceAdmin(mountSelector) {
    const sel = mountSelector || '[data-ms365-dashboard-access-mount]';
    const root = document.querySelector(sel);
    if (!root) return;
    renderTable(root);
}
