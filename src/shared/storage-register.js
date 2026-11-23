/**
 * Storage-Register – einzige Quelle für bekannte App-Storage-Keys.
 *
 * Zweck (Plan c7 / R21):
 * - Inventar für Browser-Backup und „Alles löschen“
 * - CI scheitert, wenn im Quellcode neue ms365-/webuntis-Keys ohne Eintrag entstehen
 *
 * Laufzeit: Backup/Clear nutzen weiterhin das Präfix ms365-/webuntis- (auch unbekannte
 * Keys), damit nichts „verwaist“ liegen bleibt. Das Register ist die dokumentierte Wahrheit.
 *
 * Neue Keys: hier eintragen (passende Kategorie) – sonst schlägt
 * tests/browser-backup-storage-coverage.test.mjs fehl.
 */
(function (root) {
    'use strict';

    const SESSION_SKIP_KEYS = {
        'ms365-access-granted-v1': true,
        'ms365-admin-access-granted-v1': true,
        'ms365-post-login-url': true,
        'ms365-spo-compare-session-v1': true
    };

    const INDEXED_DB_CATALOG = [
        { name: 'ms365-einsammler-attachments', stores: ['files'], label: 'Einsammler-Anhänge' },
        { name: 'ms365-elternbrief-attachments', stores: ['files'], label: 'Elternbrief-Anhänge' },
        { name: 'ms365-krankmeldung-attachments', stores: ['files'], label: 'Krankmeldung-Anhänge' },
        {
            name: 'ms365-kollegium-umfragen-attachments',
            stores: ['files'],
            label: 'Kollegium-Umfragen-Anhänge'
        }
    ];

    const FINGERPRINT_SKIP_KEYS = {
        'ms365-stammdaten-spo-sync-v1': true,
        'ms365-stammdaten-spo-sync-by-tenant-v2': true,
        'ms365-last-backup-export-at': true
    };

    const STORAGE_CATALOG = [
        {
            label: 'Zentrale Schuldaten',
            keys: [
                // inkl. years.*.unterrichtsbelegung, core.schulOrganigramm,
                // core.verwaltungOrganigramm (Legacy), core.administration / adminRoles / admin
                'ms365-schooltool-data-v2',
                'ms365-stammdaten-saved-at',
                'ms365-tenant-settings-v1',
                'ms365-school-email-domain-v1',
                'ms365-class-nick-schema-v1'
            ]
        },
        {
            label: 'Demo & Onboarding-Hinweise',
            keys: [
                'ms365-demo-mode-v1',
                'ms365-onboarding-welcome-v1',
                'ms365-it-first-run-v1',
                'ms365-it-first-save-tip-v1',
                'ms365-lehrer-welcome-v1',
                'ms365-lehrer-first-run-v1',
                'ms365-it-first-run-chip-dismiss-v1',
                'ms365-lehrer-first-run-chip-dismiss-v1'
            ]
        },
        {
            label: 'Dashboard',
            keys: [
                'ms365-dashboard-audience-groups-v1',
                'ms365-dashboard-tool-access-v1',
                'ms365-dashboard-portal-defaults-v1',
                'ms365-dashboard-favorites-v1',
                'ms365-dashboard-recent-tools-v1',
                'ms365-dashboard-order-catalog-v1',
                'ms365-dashboard-order-m365-v3',
                'ms365-dashboard-order-sharepoint-v2',
                'ms365-dashboard-order-v1',
                'ms365-dash-catalog-fold-v1',
                'ms365-dashboard-category-tab-v1',
                'ms365-dashboard-expert-open-v1',
                'ms365-dashboard-view-v1',
                'ms365-dashboard-it-preview-v1',
                'ms365-dashboard-it-preview-v1:demo-tenant',
                'ms365-it-workspace-v1',
                'ms365-it-workspace-nav-v1',
                'ms365-it-workspace-task-v1',
                'ms365-portal-prep-pane-v1',
                'ms365-dashboard-setup-dismissed-v1',
                'ms365-portal-quick-links-v1',
                'ms365-dash-main-section-v1',
                'ms365-dash-it-focus-v1',
                'ms365-dash-tasks-split-v1',
                'ms365-dash-tasks-split-nav-width-v1',
                'ms365-dash-tasks-nav-groups-v1',
                'ms365-dash-tasks-nav-groups-v2',
                'ms365-dash-tasks-nav-more-v1',
                'ms365-dash-quick-search-recent-v1',
                'ms365-dash-tools-view-v1'
            ]
        },
        {
            label: 'Dashboard (Sitzung)',
            keys: [
                'ms365-dashboard-persona-cache-v2',
                'ms365-gaeste-verwalten-active-tab-v1',
                'ms365-live-school-stammdaten-v1',
                'ms365-live-stammdaten-ensure-v1'
            ]
        },
        {
            label: 'Schulstruktur (Legacy-Spiegel)',
            keys: [
                'ms365-schulstruktur-sync-v1',
                'ms365-schulstruktur-match-v1',
                'ms365-schulstruktur-tenant-cache-v1',
                'ms365-schulstruktur-sync-ui-mode-v1',
                'ms365-schulstruktur-ad-flags-v1'
            ]
        },
        {
            label: 'Personen-Verwaltung',
            keys: ['ms365-pv-ad-flags-v1', 'ms365-pv-users-cache-v1']
        },
        {
            label: 'Kursteams / WebUntis (UI-Cache; Belegung in v2)',
            keys: ['webuntis-teams-creator-state-v1']
        },
        {
            label: 'Klassen & Jahrgang (Graph-Arbeitsstand; Listen in v2)',
            keys: ['ms365-jahrgang-state-v1', 'ms365-class-create-team-v1', 'ms365-klassenvorstaende-v1']
        },
        {
            label: 'Gruppen & SLG',
            keys: [
                'ms365-schueler-lehrer-gruppen-v1',
                'ms365-schueler-lehrer-gruppen-v2',
                'ms365-verwaltung-gruppe-v1'
            ]
        },
        {
            label: 'Fächer / ARGE / Weitere Teams (Graph-Arbeitsstand; Stammdaten in v2)',
            keys: [
                'ms365-arge-state-v1',
                'ms365-arge-state-v2',
                'ms365-wtg-state-v1',
                'ms365-arge-mail-prefix-ag-upgraded-v1'
            ]
        },
        {
            label: 'Power Automate & Automationen',
            keys: [
                'ms365-pa-onboarding-v1',
                'ms365-freistellung-setup-v1',
                'ms365-freistellung-setup-step-v1',
                'ms365-freistellung-setup-step-v2',
                'ms365-freistellung-setup-step-v3',
                'ms365-freistellung-perms-v1',
                'ms365-freistellung-kategorien-extra-v1',
                'ms365-freistellung-planer-role-v1',
                'ms365-freistellung-student-klasse-pick-v1',
                'ms365-power-automate-recipes-v1',
                'ms365-pa-termine-sync-v1',
                'ms365-pa-antraege-v1',
                'ms365-pa-seminar-v1',
                'ms365-pa-gast-v1',
                'ms365-pa-diplom-v1',
                'ms365-pa-schilf-v1',
                'ms365-pa-schularbeiten-mail-v1',
                'ms365-pa-projektwochen-mail-v1',
                'ms365-pa-done-freistellung',
                'ms365-pa-done-lehrer-freistellung',
                'ms365-pa-done-termine-sync',
                'ms365-pa-done-antraege',
                'ms365-pa-done-seminar',
                'ms365-pa-done-gast-erinnerung',
                'ms365-pa-done-diplom-ordner',
                'ms365-pa-done-schilf',
                'ms365-pa-done-srdp-mail',
                'ms365-srdp-anmeldung-v1',
                'ms365-srdp-matura-termine-v1',
                'ms365-srdp-anmeldung-role-v1',
                'ms365-srdp-list-perms-v1',
                'ms365-srdp-mail-setup-v1',
                'ms365-srdp-planer-site-v1',
                'ms365-srdp-wizard-step-v1',
                'ms365-srdp-wizard-step-v2',
                'ms365-lfr-setup-v1',
                'ms365-lfr-setup-step-v1',
                'ms365-lfr-perms-v1',
                'ms365-lfr-planer-role-v1',
                'ms365-lfr-planer-site-v1',
                'ms365-lfr-demo-items-v1',
                'ms365-lfr-outlook-event-v1',
                'ms365-akt-setup-v1',
                'ms365-akt-setup-step-v1',
                'ms365-pa-done-schulaktivitaeten'
            ]
        },
        {
            label: 'SharePoint / Intranet',
            keys: [
                'ms365-intranet-starter-v1',
                'ms365-intranet-list-links-v1',
                'ms365-intranet-hub-wizard-step-v1',
                'ms365-sps-wizard-step-v1',
                'ms365-spakt-wizard-step-v1',
                'ms365-stammdaten-it-library-v1',
                'ms365-stammdaten-it-library-by-tenant-v2',
                'ms365-it-library-meta-v1',
                'ms365-stammdaten-spo-sync-v1',
                'ms365-stammdaten-spo-sync-by-tenant-v2',
                'ms365-stammdaten-listen-perms-v1',
                'ms365-stammdaten-setup-pane-v1',
                'ms365-school-site-bootstrap-done-v1',
                'ms365-it-library-prompt-done-v1',
                'ms365-su-form-draft-v1',
                'ms365-su-pending-action-v1',
                'ms365-pw-site-url',
                'ms365-pw-demo-role',
                'ms365-projektwochen-demo-v1',
                'ms365-sa-site-url',
                'ms365-sa-demo-role',
                'ms365-sa-settings-v1',
                'ms365-sa-demo-klasse',
                'ms365-sa-schuljahr',
                'ms365-sa-cal-show-v1',
                'ms365-sa-cal-mode-v1',
                'ms365-sa-personal-cal-v1',
                'ms365-schularbeiten-perms-v1',
                'ms365-schulaktivitaeten-perms-v1',
                'ms365-akt-planer-role-v1',
                'ms365-akt-planer-site-v1',
                'ms365-akt-ics-feed-v1',
                'ms365-akt-personal-cal-v1',
                'ms365-freistellung-planer-site-v1',
                'ms365-freistellung-planer-demo-klasse-v1'
            ]
        },
        {
            label: 'Schul-Apps (Einstellungen & Stand)',
            keys: [
                'ms365-einsammler-v1',
                'ms365-einsammler-settings-v1',
                'ms365-einsammler-sp-settings-v1',
                'ms365-einsammler-inbox-v1',
                'ms365-elternbrief-v1',
                'ms365-elternbrief-settings-v1',
                'ms365-elternbrief-sp-settings-v1',
                'ms365-elternbrief-inbox-v1',
                'ms365-krankmeldung-v1',
                'ms365-krankmeldung-settings-v1',
                'ms365-krankmeldung-sp-settings-v1',
                'ms365-kollegium-umfragen-v1',
                'ms365-kollegium-umfragen-settings-v1',
                'ms365-kollegium-umfragen-sp-settings-v1',
                'ms365-ausgaben-v1',
                'ms365-ausgaben-settings-v1',
                'ms365-ausgaben-sp-settings-v1',
                'ms365-wahlfach-zuteilung-v1',
                'ms365-wahlfach-zuteilung-settings-v1',
                'ms365-schulsprecherwahl-v1',
                'ms365-teams-antrag-v1',
                'ms365-teams-antrag-settings-v1',
                'ms365-teams-antrag-sp-settings-v1',
                'ms365-diplomarbeiten-state-v1',
                'ms365-diplomarbeiten-settings-v1',
                'ms365-raumplan-v1',
                'ms365-raumplan-sp-settings-v1',
                'ms365-aufgaben-v1',
                'ms365-aufgaben-sp-settings-v1',
                'ms365-technik-meldungen-v1',
                'ms365-technik-meldungen-sp-settings-v1',
                'ms365-termine-hub-v1',
                'ms365-termine-hub-settings-v1',
                'ms365-termine-hub-cal-mode-v1',
                'ms365-kalender-postfaecher-v1',
                'ms365-schueler-stammdaten-portal-v1',
                'ms365-schueler-stammdaten-sessions-v1',
                'ms365-nc-teams-migrate-v1'
            ]
        },
        {
            label: 'Personen, Gäste, Hygiene',
            keys: [
                'ms365-hygiene-scan-v2',
                'ms365-gast-einlader-policy-v1',
                'ms365-gast-zugaenge-snapshot-v1',
                'ms365-gruppenerstellung-policy-v1'
            ]
        },
        {
            label: 'Verteiler, Postfächer & Caches',
            keys: [
                'ms365-verteilerlisten-cache-v1',
                'ms365-postfaecher-cache-v1',
                'ms365-class-chats-mirror-v1',
                'ms365-arge-export-v1',
                'ms365-wtg-export-v1'
            ]
        },
        {
            label: 'Playbooks & Spielwiesen',
            keys: [
                'ms365-cleanup-playbook-v1',
                'ms365-playbook-daten-import-verknuepfen-v2',
                'ms365-playbook-eltern-v1',
                'ms365-playbook-elternsprechtag-v1',
                'ms365-playbook-freistellungen-v1',
                'ms365-playbook-intranet-v1',
                'ms365-playbook-kursteams-v1',
                'ms365-playbook-schularbeiten-v1',
                'ms365-playbook-schuljahresstart-v1',
                'ms365-playbook-schulaktivitaeten-v1',
                'ms365-playbook-projektwochen-v1',
                'ms365-playbook-einsammler-v1',
                'ms365-playbook-elternbrief-v1',
                'ms365-playbook-krankmeldung-v1',
                'ms365-playbook-srdp-v1',
                'ms365-playbook-wahlen-v1',
                'ms365-dash-playbooks-category-v1',
                'ms365-dash-playbooks-split-nav-width-v1',
                'ms365-spielwiesen-notebook-v1',
                'ms365-spielwiesen-demo-pool-v1',
                'ms365-spielwiesen-name-pattern-v1',
                'ms365-kursteam-templates-v1',
                'ms365-kursteam-catalog-draft-v1'
            ]
        },
        {
            label: 'Weitere Werkzeuge',
            keys: [
                'ms365-elternsprechtag-bookings-v1',
                'ms365-datenlandkarte-layout-v1',
                'ms365-schulgraph-options-v1',
                'ms365-schulgraph-options-v2',
                'ms365-schul-baseline-v1',
                'ms365-schul-org-view-mode-v1',
                'ms365-schul-org-last-role-parent-v1',
                'ms365-student-lifecycle-prev-v1',
                'ms365-student-lifecycle-sammel-v1',
                'ms365-lizenzverwaltung-v1',
                'ms365-tenant-main-active-tab-v1',
                'ms365-tenant-main-tab-order-v1',
                'ms365-tenant-main-tab-order-v2',
                'ms365-tenant-students-view-v1',
                'ms365-tenant-klassenlehrer-view-v1',
                'ms365-tenant-school-email-pattern-v1',
                'ms365-tenant-teachers-email-pattern-v1',
                'ms365-tenant-v1-mirror-disable',
                'ms365-admin-active-tab-v1',
                'ms365-admin-lic-filter-v1',
                'ms365-su-wizard-step-v1',
                'ms365-dash-catalog-split-nav-width-v1',
                'ms365-last-backup-export-at'
            ]
        },
        {
            label: 'UI & Hinweise',
            keys: [
                'ms365-theme-v1',
                'ms365-brand-v1',
                'ms365-brand-custom-v1',
                'ms365-ss-graph-collapsed-v1',
                'ms365-schooltool-access-override-v1',
                'ms365-schooltool-release-notes-v1',
                'ms365-schooltool-release-notes-last-seen-at-v1',
                'ms365-license-me-v1',
                'ms365-license-me-v2',
                'ms365-admin-operator-cache-v1',
                'ms365-reported-tool-site-url-v1',
                'ms365-reported-tool-site-url-at-v1',
                'ms365-school-profile-license-pull-v1',
                'ms365-school-profile-license-push-v1'
            ]
        },
        {
            label: 'Sitzung & Runtime-Marker',
            keys: [
                'ms365-access-granted-v1',
                'ms365-admin-access-granted-v1',
                'ms365-post-login-url',
                'ms365-spo-compare-session-v1',
                'ms365-spo-auto-pull-done-v1',
                'ms365-admin-license-packages-draft-v1',
                'ms365-dashboard-portal-nav-v1',
                'ms365-portal-workspace-nav-v1',
                'ms365-schulstruktur-sync-match-v1',
                'ms365-schulstruktur-sync-tenant-cache-v1',
                'ms365-spo-config-bundle-v1'
            ]
        }
    ];

    function allRegisteredKeys() {
        const keys = [];
        STORAGE_CATALOG.forEach(function (cat) {
            (cat.keys || []).forEach(function (k) {
                keys.push(k);
            });
        });
        return keys;
    }

    function registeredKeySet() {
        return new Set(allRegisteredKeys());
    }

    function isRegisteredKey(key) {
        return registeredKeySet().has(String(key || ''));
    }

    function isAppStorageKeyPrefix(key) {
        const k = String(key || '');
        return k.indexOf('ms365-') === 0 || k.indexOf('webuntis-') === 0;
    }

    function isSessionSkipKey(key) {
        return !!SESSION_SKIP_KEYS[String(key || '')];
    }

    function isFingerprintSkipKey(key) {
        return !!FINGERPRINT_SKIP_KEYS[String(key || '')];
    }

    function duplicateKeys() {
        const seen = new Map();
        const dupes = [];
        allRegisteredKeys().forEach(function (k) {
            const n = (seen.get(k) || 0) + 1;
            seen.set(k, n);
            if (n === 2) dupes.push(k);
        });
        return dupes.sort();
    }

    const api = {
        STORAGE_CATALOG: STORAGE_CATALOG,
        INDEXED_DB_CATALOG: INDEXED_DB_CATALOG,
        SESSION_SKIP_KEYS: SESSION_SKIP_KEYS,
        FINGERPRINT_SKIP_KEYS: FINGERPRINT_SKIP_KEYS,
        allRegisteredKeys: allRegisteredKeys,
        registeredKeySet: registeredKeySet,
        isRegisteredKey: isRegisteredKey,
        isAppStorageKeyPrefix: isAppStorageKeyPrefix,
        isSessionSkipKey: isSessionSkipKey,
        isFingerprintSkipKey: isFingerprintSkipKey,
        duplicateKeys: duplicateKeys
    };

    root.ms365StorageRegister = api;
})(typeof window !== 'undefined' ? window : globalThis);
