/**
 * Schul-IT: Entra-Gruppen (aus Stammdaten) + Werkzeug-Matrix für das Dashboard.
 */
import { loadSchoolAudienceGroups } from './school-audience-groups.js';
import { mountDashboardAudienceAdmin } from './dashboard-audience-admin-ui.js';
import { clearDashboardPersonaCache } from './dashboard-persona-session.js';

function escapeHtml(s) {
    return String(s || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function groupLine(label, id, name) {
    if (!id) {
        return (
            '<p class="dash-access-groups__empty"><strong>' +
            escapeHtml(label) +
            ':</strong> noch nicht in den Stammdaten verknüpft.</p>'
        );
    }
    const title = name ? escapeHtml(name) : escapeHtml(id);
    return (
        '<p class="dash-access-groups__line"><strong>' +
        escapeHtml(label) +
        ':</strong> ' +
        title +
        ' <code class="dash-access-groups__gid">' +
        escapeHtml(id) +
        '</code></p>'
    );
}

function renderStammdatenGroupsPanel() {
    const cfg = loadSchoolAudienceGroups();
    return (
        groupLine('Lehrkräfte-Sammelgruppe', cfg.groupLehrerId, cfg.groupLehrerName) +
        groupLine('Schüler-Sammelgruppe', cfg.groupSchuelerId, cfg.groupSchuelerName)
    );
}

export function mountDashboardWerkzeugZugriffPage() {
    const root = document.getElementById('dashWerkzeugZugriffRoot');
    if (!root) return;

    root.innerHTML = `
        <section class="dash-access-groups tile" aria-labelledby="dashAudGroupsTitle">
            <div class="tile-head">
                <h2 id="dashAudGroupsTitle">Standard-Gruppen (Stammdaten)</h2>
                <p class="tile-subtitle">
                    Lehrkräfte- und Schüler-Sammelgruppe aus den <strong>Stammdaten</strong> (Einrichtung / MS&nbsp;365-Gruppenverwaltung).
                    Dashboard, Schularbeiten- und Freistellungs-Planer nutzen dieselben Gruppen.
                    Global Administratorinnen/Administratoren und eingetragene IT-Kontakte sehen den vollen Katalog.
                </p>
            </div>
            <div id="dashAudStammdatenGroups" class="dash-access-groups__readonly"></div>
            <p class="dash-access-groups__note">
                Pflege: <a href="tenant.html">Stammdaten</a> oder
                <a href="tools/schulstruktur-sync.html">MS&nbsp;365 Gruppenverwaltung</a>.
                Legacy-Fallback (<code>ms365-dashboard-audience-groups-v1</code>) gilt nur, wenn in den Stammdaten noch keine Sammelgruppen gesetzt sind.
            </p>
        </section>
        <div data-ms365-dashboard-access-mount></div>`;

    const panel = document.getElementById('dashAudStammdatenGroups');
    if (panel) panel.innerHTML = renderStammdatenGroupsPanel();

    const refresh = () => {
        clearDashboardPersonaCache();
        if (panel) panel.innerHTML = renderStammdatenGroupsPanel();
    };
    window.addEventListener('storage', (ev) => {
        if (ev && ev.key === 'ms365-schooltool-data-v2') refresh();
    });

    mountDashboardAudienceAdmin('[data-ms365-dashboard-access-mount]');
}
