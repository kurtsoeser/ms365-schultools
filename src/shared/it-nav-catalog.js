/**
 * Ein IT-Menü für Startseite und Werkzeuge.
 * Einträge, Icons und Reihenfolge stehen hier – nicht im HTML der Startseite
 * und nicht in einer Session-Kopie.
 */
import { stammdatenImportNavTools, stammdatenRegisterNavTools } from './stammdaten-nav-catalog.js';

/**
 * @param {string} href
 * @param {string} label
 * @param {string} icon
 * @param {boolean} [expert]
 */
function tool(href, label, icon, expert) {
    return { href: href, label: label, icon: icon, expert: !!expert };
}

function stammdatenTools() {
    const register = stammdatenRegisterNavTools().map(function (item) {
        return tool(item.href, item.label, item.icon, false);
    });
    const imported = stammdatenImportNavTools().map(function (item) {
        return tool(item.href, item.label, item.icon, true);
    });
    return register.concat(imported);
}

/**
 * @returns {{ id: string, label: string, icon: string, overviewHash: string, tools: { href: string, label: string, icon: string, expert: boolean }[] }[]}
 */
export function itNavCatalog() {
    return [
        {
            id: 'dashTaskImportVerknuepfen',
            label: 'Stammdaten',
            icon: 'bi bi-database',
            overviewHash: 'import',
            tools: stammdatenTools()
        },
        {
            id: 'dashTaskGruppen',
            label: 'MS365 Mitgliedschaft',
            icon: 'bi bi-people',
            overviewHash: 'mitgliedschaften',
            tools: [
                tool('tools/schueler-lehrer-gruppen.html?kind=schueler', 'Schüler:innen-Sammelgruppe', 'bi bi-mortarboard-fill'),
                tool('tools/schueler-lehrer-gruppen.html?kind=lehrer', 'Lehrer:innen-Sammelgruppe', 'bi bi-person-workspace'),
                tool('tools/verwaltung.html?kind=schulleitung', 'Schulleitung', 'bi bi-person-badge'),
                tool('tools/verwaltung.html?kind=verwaltung', 'Verwaltung (Personal)', 'bi bi-building-gear'),
                tool('tools/klassenvorstaende.html', 'Klassenvorstände', 'bi bi-person-vcard'),
                tool('tools/klassenchats.html', 'KlassenChats', 'bi bi-chat-dots'),
                tool('tools/arge-fachgruppen.html', 'Fächer und ARGEs', 'bi bi-collection'),
                tool('tools/datenhygiene.html', 'Datenhygiene (Abgleich M365)', 'bi bi-shield-check', true),
                tool('tools/weitere-teams-gruppen.html', 'Weitere Teams & Gruppen', 'bi bi-magic', true)
            ]
        },
        {
            id: 'dashTaskUnterricht',
            label: 'Klassen & Unterricht',
            icon: 'bi bi-mortarboard',
            overviewHash: 'klassen',
            tools: [
                tool('tools/jahrgangsgruppen.html', 'Klassengruppen', 'bi bi-diagram-3'),
                tool('tools/kursteams.html', 'Unterrichtsteams', 'bi bi-rocket-takeoff'),
                tool('tools/kursteams.html?mode=single', 'Einzelne Unterrichtsteams', 'bi bi-plus-square'),
                tool('tools/playbook-kursteams.html', 'Playbook Unterrichtsteams', 'bi bi-list-check'),
                tool('tools/unterrichtsteams-katalog.html', 'Unterrichtsteams-Katalog', 'bi bi-table', true),
                tool('tools/klassen-merge.html', 'Klassen zusammenlegen', 'bi bi-intersect', true),
                tool('tools/onenote-verteilung.html', 'OneNote-Inhalte verteilen', 'bi bi-journal-richtext', true)
            ]
        },
        {
            id: 'dashTaskPersonen',
            label: 'Personen',
            icon: 'bi bi-person-badge',
            overviewHash: 'personen',
            tools: [
                tool('tools/personen-verwaltung.html', 'Personen suchen', 'bi bi-search'),
                tool('tools/personen-verwaltung.html?create=1', 'Konto anlegen', 'bi bi-person-plus'),
                tool('tools/schueler-lifecycle.html', 'Schüler-Lifecycle', 'bi bi-arrow-left-right'),
                tool('tools/gaeste-verwalten.html', 'Gäste', 'bi bi-shield-check', true),
                tool('tools/lizenzverwaltung.html', 'Lizenzverwaltung', 'bi bi-key', true),
                tool('tools/namenskonvention-audit.html', 'Namenskonvention prüfen', 'bi bi-spellcheck', true)
            ]
        },
        {
            id: 'dashTaskSchuljahr',
            label: 'Schuljahresstart',
            icon: 'bi bi-calendar2-range',
            overviewHash: 'schuljahresstart',
            tools: [
                tool('tools/playbook-schuljahresstart.html', 'Playbook Schuljahresstart', 'bi bi-rocket-takeoff'),
                tool('tools/organisations-assistent.html', 'Schuljahr-Assistent', 'bi bi-calendar2-check'),
                tool('tools/webuntis-stammdaten-import.html', 'WebUntis-Stammdaten importieren', 'bi bi-file-earmark-spreadsheet'),
                tool('tools/klassen-umbenennen.html', 'Klassen umbenennen', 'bi bi-input-cursor-text', true),
                tool('tools/cleanup-playbook.html', 'Cleanup zum Schuljahr', 'bi bi-broom', true)
            ]
        },
        {
            id: 'dashTaskIntranet',
            label: 'Intranet',
            icon: 'bi bi-house-door',
            overviewHash: 'intranet',
            tools: [
                tool('tools/playbook-intranet.html', 'Playbook Intranet', 'bi bi-list-check'),
                tool('tools/sharepoint-intranet-hub.html', 'Schul-Intranet', 'bi bi-house-door'),
                tool('tools/sharepoint-liste-stammdaten.html', 'Stammdaten-Listen (IT)', 'bi bi-list-columns'),
                tool('tools/sharepoint-liste-lehrer.html', 'Lehrerliste', 'bi bi-table', true),
                tool('tools/sharepoint-liste-schultermine.html', 'Schultermine-Liste', 'bi bi-calendar3', true)
            ]
        },
        {
            id: 'dashTaskSchulApps',
            label: 'Schul-Apps',
            icon: 'bi bi-window-stack',
            overviewHash: 'schulapps',
            tools: [
                tool('dashboard-werkzeug-zugriff.html', 'Portal vorbereiten', 'bi bi-sliders2-vertical'),
                tool('tools/schularbeiten-planer.html', 'Schularbeiten-Planer', 'bi bi-journal-check'),
                tool('tools/freistellung-planer.html', 'Freistellungen (Schüler)', 'bi bi-calendar2-check'),
                tool('tools/lehrer-freistellung-planer.html', 'Freistellungen Lehrkräfte', 'bi bi-briefcase'),
                tool('tools/schulsprecherwahl.html', 'Schulwahlen', 'bi bi-ui-checks-grid'),
                tool('tools/projektwochen.html', 'Projektwochen', 'bi bi-calendar-week'),
                tool('tools/schulaktivitaeten-planer.html', 'Schulaktivitäten', 'bi bi-calendar-event'),
                tool('tools/sharepoint-liste-schularbeiten.html', 'Schularbeiten-Listen (IT)', 'bi bi-list-ul'),
                tool('tools/playbook-freistellungen.html', 'Playbook Freistellungen', 'bi bi-list-check'),
                tool('tools/playbook-schularbeiten-planer.html', 'Playbook Schularbeiten', 'bi bi-list-check'),
                tool('tools/playbook-eltern.html', 'Playbook Elternkommunikation', 'bi bi-chat-heart'),
                tool('tools/elternsprechtag-bookings.html', 'Elternsprechtag (Bookings)', 'bi bi-calendar2-heart'),
                tool('tools/playbook-elternsprechtag.html', 'Playbook Elternsprechtag', 'bi bi-list-check')
            ]
        },
        {
            id: 'dashTaskOrdnung',
            label: 'Aufräumen',
            icon: 'bi bi-ui-checks-grid',
            overviewHash: 'aufraeumen',
            tools: [
                tool('tools/datenhygiene.html', 'Datenhygiene', 'bi bi-shield-check'),
                tool('tools/cleanup-playbook.html', 'Cleanup-Playbook', 'bi bi-list-check'),
                tool('tools/leere-gruppen-report.html?problem=no-owners', 'Leere Gruppen prüfen', 'bi bi-clipboard2-pulse'),
                tool('tools/teams-archiv.html', 'Teams archivieren (Bulk)', 'bi bi-archive'),
                tool('tools/schulstruktur-sync.html', 'Alle Gruppen', 'bi bi-diagram-2', true),
                tool('tools/datenlandkarte.html', 'Datenlandkarte', 'bi bi-diagram-3', true),
                tool('tools/stammdaten-backup-abgleich.html', 'Backup-Abgleich', 'bi bi-columns-gap', true)
            ]
        }
    ];
}

/**
 * @param {string} taskId
 * @returns {{ href: string, label: string, icon: string, expert: boolean }[]}
 */
export function itNavToolsForTask(taskId) {
    const id = String(taskId || '');
    const node = itNavCatalog().find(function (item) {
        return item.id === id;
    });
    return node && Array.isArray(node.tools) ? node.tools.map(function (t) {
        return Object.assign({}, t);
    }) : [];
}

/**
 * @param {string} taskId
 * @returns {string}
 */
export function itNavAreaLabel(taskId) {
    const node = itNavCatalog().find(function (item) {
        return item.id === String(taskId || '');
    });
    return node ? node.label : '';
}

/**
 * @param {string} taskId
 * @returns {string}
 */
export function itNavAreaIcon(taskId) {
    const node = itNavCatalog().find(function (item) {
        return item.id === String(taskId || '');
    });
    return node ? node.icon : '';
}
