/**
 * Schul-IT: „Portal vorbereiten“ – Schule, Gestalten, Zugriff, Freigeben.
 * Globale IT-Leiste links (Schul-Apps); hier nur horizontale Schritte.
 */
import { loadSchoolAudienceGroups } from './school-audience-groups.js';
import { mountDashboardAudienceAdmin } from './dashboard-audience-admin-ui.js';
import { mountDashboardPortalDefaultsAdmin } from './dashboard-portal-defaults-admin-ui.js';
import { mountDashboardSchoolProfileSetup } from './dashboard-school-profile-setup-ui.js';
import { mountPortalReleaseGateUi } from './portal-release-gate-ui.js';
import { clearDashboardPersonaCache } from './dashboard-persona-session.js';

const PANE_IDS = ['schule', 'gestalten', 'zugriff', 'freigeben'];
const STORAGE_KEY = 'ms365-portal-prep-pane-v1';

const PANE_META = {
    schule: {
        label: 'Schule',
        icon: 'bi-building',
        lead: 'Name, Logo und Adresse für die Karte „Meine Schule“.'
    },
    gestalten: {
        label: 'Gestalten',
        icon: 'bi-layout-sidebar',
        lead: 'Home-Pins, Sidebar und Vorschau für Lehrkräfte und Schüler:innen.'
    },
    zugriff: {
        label: 'Zugriff',
        icon: 'bi-sliders',
        lead: 'Sammelgruppen und welche Apps Lehrkräfte bzw. Schüler:innen sehen.'
    },
    freigeben: {
        label: 'Freigeben',
        icon: 'bi-door-open',
        lead: 'Go-Live: Lehrer-Portal freigeben, wenn Einrichtung steht.'
    }
};

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

function readStoredPane() {
    try {
        const v = String(localStorage.getItem(STORAGE_KEY) || '').trim();
        return PANE_IDS.includes(v) ? v : '';
    } catch {
        return '';
    }
}

function storePane(paneId) {
    try {
        localStorage.setItem(STORAGE_KEY, paneId);
    } catch {
        /* ignore */
    }
}

function paneFromHash() {
    try {
        const h = String(window.location.hash || '')
            .replace(/^#/, '')
            .trim()
            .toLowerCase();
        if (PANE_IDS.includes(h)) return h;
        if (h === 'portal-release' || h === 'release') return 'freigeben';
        if (h === 'profile' || h === 'school') return 'schule';
        if (h === 'design' || h === 'pins') return 'gestalten';
        if (h === 'access' || h === 'apps') return 'zugriff';
    } catch {
        /* ignore */
    }
    return '';
}

/**
 * @param {HTMLElement} setup
 * @param {string} paneId
 * @param {{ persist?: boolean, hash?: boolean }} [opts]
 */
function activatePane(setup, paneId, opts) {
    let next = String(paneId || '').trim();
    if (!PANE_IDS.includes(next)) next = 'schule';

    setup.querySelectorAll('[data-portal-prep-pane]').forEach(function (btn) {
        const id = String(btn.getAttribute('data-portal-prep-pane') || '');
        const on = id === next;
        btn.classList.toggle('is-active', on);
        btn.setAttribute('aria-current', on ? 'page' : 'false');
    });

    setup.querySelectorAll('[data-portal-prep-panel]').forEach(function (panel) {
        const id = String(panel.getAttribute('data-portal-prep-panel') || '');
        const on = id === next;
        panel.hidden = !on;
        if (on) panel.removeAttribute('aria-hidden');
        else panel.setAttribute('aria-hidden', 'true');
    });

    setup.setAttribute('data-active-pane', next);

    const meta = PANE_META[next];
    const lead = setup.querySelector('[data-portal-prep-lead]');
    if (lead && meta) lead.textContent = meta.lead;

    if (!opts || opts.persist !== false) storePane(next);
    if (!opts || opts.hash !== false) {
        try {
            const url = new URL(window.location.href);
            url.hash = next;
            history.replaceState(null, '', url.pathname + url.search + '#' + next);
        } catch {
            /* ignore */
        }
    }
}

function buildShellHtml() {
    const steps = PANE_IDS.map(function (id, i) {
        const m = PANE_META[id];
        return (
            '<button type="button" class="dash-portal-prep__step-btn' +
            (i === 0 ? ' is-active' : '') +
            '" data-portal-prep-pane="' +
            id +
            '" aria-current="' +
            (i === 0 ? 'page' : 'false') +
            '">' +
            '<span class="dash-portal-prep__step-num">' +
            String(i + 1) +
            '</span>' +
            '<span class="dash-portal-prep__step-label"><i class="bi ' +
            m.icon +
            '" aria-hidden="true"></i> ' +
            escapeHtml(m.label) +
            '</span>' +
            '</button>'
        );
    }).join('');

    return (
        '<div class="dash-portal-prep" id="dashPortalPrep" data-active-pane="schule">' +
        '<nav class="dash-portal-prep__steps" aria-label="Schritte Portal vorbereiten">' +
        steps +
        '</nav>' +
        '<p class="dash-portal-prep__lead" data-portal-prep-lead>' +
        escapeHtml(PANE_META.schule.lead) +
        '</p>' +
        '<div class="dash-portal-prep__panes">' +
        '<section class="dash-portal-prep__pane" data-portal-prep-panel="schule" aria-label="Schule im Portal">' +
        '<div data-ms365-dashboard-school-profile-mount></div>' +
        '</section>' +
        '<section class="dash-portal-prep__pane" data-portal-prep-panel="gestalten" aria-label="Portal gestalten" hidden>' +
        '<div data-ms365-dashboard-portal-defaults-mount></div>' +
        '</section>' +
        '<section class="dash-portal-prep__pane" data-portal-prep-panel="zugriff" aria-label="Werkzeug-Zugriff" hidden>' +
        '<section class="dash-pg-section dash-access-groups" aria-labelledby="dashAudGroupsTitle">' +
        '<div class="dash-pg-section__head"><div>' +
        '<h3 id="dashAudGroupsTitle">Standard-Gruppen</h3>' +
        '<p class="dash-pg-section__lead">Lehrkräfte- und Schüler-Sammelgruppe aus den Stammdaten – Dashboard und Planer nutzen dieselben Gruppen.</p>' +
        '</div></div>' +
        '<div id="dashAudStammdatenGroups" class="dash-access-groups__readonly"></div>' +
        '<p class="dash-access-groups__note">' +
        'Pflege: <a href="tenant.html">Stammdaten</a> oder ' +
        '<a href="tools/schulstruktur-sync.html">MS&nbsp;365 Gruppenverwaltung</a>.' +
        '</p>' +
        '</section>' +
        '<div data-ms365-dashboard-access-mount></div>' +
        '</section>' +
        '<section class="dash-portal-prep__pane" data-portal-prep-panel="freigeben" aria-label="Portal freigeben" hidden>' +
        '<div data-ms365-dashboard-portal-release-mount></div>' +
        '</section>' +
        '</div>' +
        '<div class="dash-portal-prep__footer">' +
        '<button type="button" class="btn alt" data-portal-prep-prev hidden>Zurück</button>' +
        '<button type="button" class="btn" data-portal-prep-next>Weiter: Gestalten</button>' +
        '</div>' +
        '</div>'
    );
}

/**
 * @param {HTMLElement} setup
 */
function syncStepButtons(setup) {
    const cur = String(setup.getAttribute('data-active-pane') || 'schule');
    const idx = PANE_IDS.indexOf(cur);
    const prev = setup.querySelector('[data-portal-prep-prev]');
    const next = setup.querySelector('[data-portal-prep-next]');
    if (prev) {
        prev.hidden = idx <= 0;
        prev.textContent = idx > 0 ? 'Zurück: ' + PANE_META[PANE_IDS[idx - 1]].label : 'Zurück';
    }
    if (next) {
        if (idx >= PANE_IDS.length - 1) {
            next.hidden = true;
        } else {
            next.hidden = false;
            next.textContent = 'Weiter: ' + PANE_META[PANE_IDS[idx + 1]].label;
        }
    }
}

export function mountDashboardWerkzeugZugriffPage() {
    const root = document.getElementById('dashWerkzeugZugriffRoot');
    if (!root) return;

    root.innerHTML = buildShellHtml();
    const setup = document.getElementById('dashPortalPrep');
    if (!setup) return;

    const panel = document.getElementById('dashAudStammdatenGroups');
    if (panel) panel.innerHTML = renderStammdatenGroupsPanel();

    const refresh = () => {
        clearDashboardPersonaCache();
        if (panel) panel.innerHTML = renderStammdatenGroupsPanel();
    };
    window.addEventListener('storage', (ev) => {
        if (ev && ev.key === 'ms365-schooltool-data-v2') refresh();
    });

    setup.querySelectorAll('[data-portal-prep-pane]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            const id = String(btn.getAttribute('data-portal-prep-pane') || '');
            activatePane(setup, id);
            syncStepButtons(setup);
        });
    });

    const prevBtn = setup.querySelector('[data-portal-prep-prev]');
    const nextBtn = setup.querySelector('[data-portal-prep-next]');
    if (prevBtn) {
        prevBtn.addEventListener('click', function () {
            const cur = String(setup.getAttribute('data-active-pane') || 'schule');
            const idx = PANE_IDS.indexOf(cur);
            if (idx > 0) {
                activatePane(setup, PANE_IDS[idx - 1]);
                syncStepButtons(setup);
            }
        });
    }
    if (nextBtn) {
        nextBtn.addEventListener('click', function () {
            const cur = String(setup.getAttribute('data-active-pane') || 'schule');
            const idx = PANE_IDS.indexOf(cur);
            if (idx < PANE_IDS.length - 1) {
                activatePane(setup, PANE_IDS[idx + 1]);
                syncStepButtons(setup);
            }
        });
    }

    window.addEventListener('hashchange', function () {
        const fromHash = paneFromHash();
        if (fromHash) {
            activatePane(setup, fromHash, { hash: false });
            syncStepButtons(setup);
        }
    });

    mountDashboardSchoolProfileSetup('[data-ms365-dashboard-school-profile-mount]');
    mountDashboardPortalDefaultsAdmin('[data-ms365-dashboard-portal-defaults-mount]');
    mountDashboardAudienceAdmin('[data-ms365-dashboard-access-mount]');
    mountPortalReleaseGateUi('[data-ms365-dashboard-portal-release-mount]');

    const initial = paneFromHash() || readStoredPane() || 'schule';
    activatePane(setup, initial);
    syncStepButtons(setup);
}
