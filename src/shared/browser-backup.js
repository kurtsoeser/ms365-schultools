/**
 * Vollständiges lokales Browser-Backup (localStorage + wiederherstellbare
 * sessionStorage-Daten der App).
 * Zum Übertragen zwischen Browsern/PCs. Microsoft-Anmeldung (MSAL) bleibt
 * bewusst außen vor – im Zielbrowser erneut anmelden.
 */
(function () {
    'use strict';

    const KIND = 'ms365-browser-backup-v1';
    const VERSION = 7;
    const SESSION_SKIP_KEYS = {
        'ms365-access-granted-v1': true,
        'ms365-admin-access-granted-v1': true,
        'ms365-post-login-url': true,
        'ms365-spo-compare-session-v1': true
    };

    /**
     * IndexedDB-Anhänge der Schul-Apps (liegen nicht in localStorage).
     * Werden in Backup-Version ≥ 7 unter payload.indexedDB mitgesichert.
     */
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

    /** Keys, die den Inhalts-Fingerprint nicht beeinflussen (Sync-Meta selbst). */
    const FINGERPRINT_SKIP_KEYS = {
        'ms365-stammdaten-spo-sync-v1': true,
        'ms365-stammdaten-spo-sync-by-tenant-v2': true,
        'ms365-last-backup-export-at': true
    };

    /**
     * Bekannte Schlüssel – Inventar-Anzeige.
     * Export/Sync umfassen weiterhin alle ms365-* / webuntis-* Keys (auch unbekannte).
     */
    const STORAGE_CATALOG = [
        {
            label: 'Zentrale Schuldaten',
            keys: [
                // inkl. years.*.unterrichtsbelegung, core.schulOrganigramm,
                // core.verwaltungOrganigramm (Legacy), core.administration / adminRoles / admin
                'ms365-schooltool-data-v2',
                'ms365-tenant-settings-v1',
                'ms365-school-email-domain-v1',
                'ms365-class-nick-schema-v1'
            ]
        },
        {
            label: 'Demo & Onboarding-Hinweise',
            keys: ['ms365-demo-mode-v1', 'ms365-onboarding-welcome-v1']
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
                'ms365-dashboard-setup-dismissed-v1',
                'ms365-portal-quick-links-v1',
                'ms365-dash-main-section-v1',
                'ms365-dash-it-focus-v1',
                'ms365-dash-tasks-split-v1',
                'ms365-dash-tasks-split-nav-width-v1'
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
                'ms365-freistellung-planer-role-v1',
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
                'ms365-admin-operator-cache-v1'
            ]
        }
    ];

    function getStore(storage) {
        if (storage) return storage;
        try {
            if (typeof localStorage !== 'undefined') return localStorage;
        } catch {
            /* ignore */
        }
        return null;
    }

    function getSessionStore(storage) {
        if (storage) return storage;
        try {
            if (typeof sessionStorage !== 'undefined') return sessionStorage;
        } catch {
            /* ignore */
        }
        return null;
    }

    function openIndexedDb(name, storeNames) {
        return new Promise(function (resolve, reject) {
            if (typeof indexedDB === 'undefined') {
                resolve(null);
                return;
            }
            const req = indexedDB.open(String(name || ''), 1);
            req.onupgradeneeded = function () {
                const db = req.result;
                (storeNames || []).forEach(function (storeName) {
                    if (storeName && !db.objectStoreNames.contains(storeName)) {
                        db.createObjectStore(storeName);
                    }
                });
            };
            req.onsuccess = function () {
                resolve(req.result);
            };
            req.onerror = function () {
                reject(req.error || new Error('IndexedDB konnte nicht geöffnet werden.'));
            };
        });
    }

    function readObjectStoreRecords(db, storeName) {
        return new Promise(function (resolve, reject) {
            if (!db || !storeName || !db.objectStoreNames.contains(storeName)) {
                resolve([]);
                return;
            }
            const tx = db.transaction(storeName, 'readonly');
            const store = tx.objectStore(storeName);
            const rows = [];
            const req = store.openCursor();
            req.onsuccess = function () {
                const cursor = req.result;
                if (!cursor) {
                    resolve(rows);
                    return;
                }
                rows.push({ key: cursor.key, value: cursor.value });
                cursor.continue();
            };
            req.onerror = function () {
                reject(req.error || new Error('IndexedDB lesen fehlgeschlagen.'));
            };
        });
    }

    function writeObjectStoreRecords(db, storeName, rows) {
        return new Promise(function (resolve, reject) {
            if (!db || !storeName) {
                resolve(0);
                return;
            }
            if (!db.objectStoreNames.contains(storeName)) {
                resolve(0);
                return;
            }
            const list = Array.isArray(rows) ? rows : [];
            const tx = db.transaction(storeName, 'readwrite');
            const store = tx.objectStore(storeName);
            store.clear();
            let written = 0;
            list.forEach(function (row) {
                if (!row || row.key == null) return;
                store.put(row.value, row.key);
                written += 1;
            });
            tx.oncomplete = function () {
                resolve(written);
            };
            tx.onerror = function () {
                reject(tx.error || new Error('IndexedDB schreiben fehlgeschlagen.'));
            };
        });
    }

    function deleteIndexedDb(name) {
        return new Promise(function (resolve) {
            if (typeof indexedDB === 'undefined' || !name) {
                resolve(false);
                return;
            }
            const req = indexedDB.deleteDatabase(String(name));
            req.onsuccess = function () {
                resolve(true);
            };
            req.onerror = function () {
                resolve(false);
            };
            req.onblocked = function () {
                resolve(false);
            };
        });
    }

    /**
     * Liest alle bekannten App-IndexedDBs (Anhänge).
     * @returns {Promise<Record<string, { stores: Record<string, Array<{key:*, value:*}>>, recordCount: number }>>}
     */
    async function collectIndexedDb() {
        const out = {};
        for (let i = 0; i < INDEXED_DB_CATALOG.length; i++) {
            const entry = INDEXED_DB_CATALOG[i];
            try {
                const db = await openIndexedDb(entry.name, entry.stores);
                if (!db) continue;
                const stores = {};
                let recordCount = 0;
                for (let s = 0; s < entry.stores.length; s++) {
                    const storeName = entry.stores[s];
                    const rows = await readObjectStoreRecords(db, storeName);
                    stores[storeName] = rows;
                    recordCount += rows.length;
                }
                try {
                    db.close();
                } catch {
                    /* ignore */
                }
                if (recordCount > 0) {
                    out[entry.name] = {
                        label: entry.label,
                        stores: stores,
                        recordCount: recordCount
                    };
                }
            } catch {
                /* einzelne DB-Fehler sollen Export nicht stoppen */
            }
        }
        return out;
    }

    /**
     * Stellt IndexedDB-Anhänge aus einem Backup wieder her.
     * @param {object} indexed
     * @param {{ replace?: boolean }} [opts]
     */
    async function applyIndexedDb(indexed, opts) {
        const replace = !opts || opts.replace !== false;
        const incoming = indexed && typeof indexed === 'object' && !Array.isArray(indexed) ? indexed : {};
        const written = [];
        const cleared = [];
        const errors = [];
        const knownNames = INDEXED_DB_CATALOG.map(function (e) {
            return e.name;
        });

        for (let i = 0; i < INDEXED_DB_CATALOG.length; i++) {
            const entry = INDEXED_DB_CATALOG[i];
            const dump = Object.prototype.hasOwnProperty.call(incoming, entry.name) ? incoming[entry.name] : null;
            if (!dump) {
                if (replace) {
                    const ok = await deleteIndexedDb(entry.name);
                    if (ok) cleared.push(entry.name);
                }
                continue;
            }
            try {
                const db = await openIndexedDb(entry.name, entry.stores);
                if (!db) continue;
                const stores = dump.stores && typeof dump.stores === 'object' ? dump.stores : {};
                let n = 0;
                for (let s = 0; s < entry.stores.length; s++) {
                    const storeName = entry.stores[s];
                    n += await writeObjectStoreRecords(db, storeName, stores[storeName] || []);
                }
                try {
                    db.close();
                } catch {
                    /* ignore */
                }
                written.push(entry.name + ':' + n);
            } catch (e) {
                errors.push(entry.name + ': ' + (e && e.message ? e.message : String(e)));
            }
        }

        if (replace) {
            Object.keys(incoming).forEach(function (name) {
                if (knownNames.indexOf(name) >= 0) return;
                /* Unbekannte IDBs aus älteren/neuen Backups: best-effort anlegen */
            });
        }

        return { written: written, cleared: cleared, errors: errors };
    }

    async function clearIndexedDbCatalog() {
        const cleared = [];
        for (let i = 0; i < INDEXED_DB_CATALOG.length; i++) {
            const ok = await deleteIndexedDb(INDEXED_DB_CATALOG[i].name);
            if (ok) cleared.push(INDEXED_DB_CATALOG[i].name);
        }
        return cleared;
    }

    function indexedDbSummary(indexed) {
        const obj = indexed && typeof indexed === 'object' ? indexed : {};
        const names = Object.keys(obj);
        if (!names.length) return '';
        const bits = names.map(function (name) {
            const row = obj[name] || {};
            const label = row.label || name;
            const n = Number(row.recordCount) || 0;
            return label + (n ? ' (' + n + ')' : '');
        });
        return 'IndexedDB: ' + bits.join(', ');
    }

    function isAuthKey(key) {
        const k = String(key || '');
        if (/^msal[.\-]/i.test(k)) return true;
        if (/login\.microsoftonline\.com/i.test(k)) return true;
        if (/login\.windows\.net/i.test(k)) return true;
        if (/login\.microsoft\.com/i.test(k)) return true;
        return false;
    }

    function isAppStorageKey(key) {
        if (isAuthKey(key)) return false;
        const k = String(key || '');
        return k.indexOf('ms365-') === 0 || k.indexOf('webuntis-') === 0;
    }

    function isRestorableSessionKey(key) {
        const k = String(key || '');
        if (!isAppStorageKey(k)) return false;
        if (SESSION_SKIP_KEYS[k]) return false;
        if (k.indexOf('ms365-spo-auto-pull-done-v1') === 0) return false;
        return true;
    }

    function listAppKeys(storage) {
        const store = getStore(storage);
        if (!store || typeof store.key !== 'function') return [];
        const keys = [];
        const n = store.length || 0;
        for (let i = 0; i < n; i++) {
            const k = store.key(i);
            if (k && isAppStorageKey(k)) keys.push(k);
        }
        keys.sort();
        return keys;
    }

    function encodeValue(raw) {
        if (raw == null) return '';
        const s = String(raw);
        try {
            const parsed = JSON.parse(s);
            if (parsed && typeof parsed === 'object') return parsed;
        } catch {
            /* keep raw string */
        }
        return s;
    }

    function decodeValue(value) {
        if (value == null) return '';
        if (typeof value === 'string') return value;
        try {
            return JSON.stringify(value);
        } catch {
            return String(value);
        }
    }

    function collectLocalStorage(storage) {
        const store = getStore(storage);
        const out = {};
        if (!store) return out;
        listAppKeys(store).forEach(function (k) {
            try {
                const raw = store.getItem(k);
                if (raw == null) return;
                out[k] = encodeValue(raw);
            } catch {
                /* ignore unreadable keys */
            }
        });
        return out;
    }

    function collectSessionStorage(storage) {
        const store = getSessionStore(storage);
        const out = {};
        if (!store || typeof store.key !== 'function') return out;
        const n = store.length || 0;
        for (let i = 0; i < n; i++) {
            const k = store.key(i);
            if (!k || !isRestorableSessionKey(k)) continue;
            try {
                const raw = store.getItem(k);
                if (raw == null) continue;
                out[k] = encodeValue(raw);
            } catch {
                /* ignore unreadable keys */
            }
        }
        return out;
    }

    function asDate(d) {
        if (d && typeof d.getFullYear === 'function' && typeof d.toISOString === 'function') return d;
        return new Date();
    }

    function localDateStamp(d) {
        const x = asDate(d);
        const y = x.getFullYear();
        const m = String(x.getMonth() + 1).padStart(2, '0');
        const day = String(x.getDate()).padStart(2, '0');
        return y + '-' + m + '-' + day;
    }

    function localDateTimeStamp(d) {
        const x = asDate(d);
        const hh = String(x.getHours()).padStart(2, '0');
        const mm = String(x.getMinutes()).padStart(2, '0');
        return localDateStamp(x) + '_' + hh + mm;
    }

    function sanitizeForFilename(raw) {
        return String(raw || '')
            .trim()
            .replace(/[^a-zA-Z0-9äöüÄÖÜß._-]/g, '_')
            .replace(/_+/g, '_')
            .replace(/^_|_$/g, '')
            .slice(0, 48);
    }

    function readSchoolMeta(storage) {
        try {
            const store = getStore(storage);
            if (!store) return { schoolName: '', domain: '' };
            const raw = store.getItem('ms365-schooltool-data-v2');
            if (!raw) return { schoolName: '', domain: '' };
            const parsed = JSON.parse(raw);
            const core = parsed && parsed.core ? parsed.core : {};
            return {
                schoolName: String(core.schoolName || '').trim(),
                domain: String(core.domain || '').trim()
            };
        } catch {
            return { schoolName: '', domain: '' };
        }
    }

    /**
     * Kurzstatistik aus ms365-schooltool-data-v2 für Inventory/Toast.
     * @param {object} local collectLocalStorage-Ergebnis
     */
    function summarizeSchooltoolForBackup(local) {
        const raw = local && local['ms365-schooltool-data-v2'];
        if (raw == null || raw === '') return null;
        try {
            const v2 = typeof raw === 'object' ? raw : JSON.parse(String(raw));
            const core = v2 && v2.core && typeof v2.core === 'object' ? v2.core : {};
            const org = core.schulOrganigramm && typeof core.schulOrganigramm === 'object' ? core.schulOrganigramm : null;
            const nodes = org && Array.isArray(org.nodes) ? org.nodes : [];
            const cur = v2 && v2.years ? String(v2.years.current || '').trim() : '';
            const bucket =
                cur && v2.years && v2.years.byLabel && v2.years.byLabel[cur] ? v2.years.byLabel[cur] : null;
            const setup = v2 && v2.setup && typeof v2.setup === 'object' ? v2.setup : {};
            const profile =
                setup.schoolProfile && typeof setup.schoolProfile === 'object' ? setup.schoolProfile : {};
            const listTitles =
                setup.intranetListTitles && typeof setup.intranetListTitles === 'object'
                    ? setup.intranetListTitles
                    : {};
            const logoDataUrl = String(profile.logoDataUrl || '').trim();
            const logoUrl = String(profile.logoUrl || '').trim();
            return {
                schoolName: String(core.schoolName || '').trim(),
                domain: String(core.domain || '').trim(),
                schoolCode: String(profile.schoolCode || '').trim(),
                hasSchoolProfileLogo: !!(logoUrl || logoDataUrl),
                schoolProfileStreet: String(profile.street || '').trim(),
                schoolProfileCity: String(profile.city || '').trim(),
                schoolProfileEmail: String(profile.email || '').trim(),
                schoolProfileWebsite: String(profile.website || '').trim(),
                teacherCount: Array.isArray(core.teachers) ? core.teachers.length : 0,
                studentCount: bucket && Array.isArray(bucket.students) ? bucket.students.length : 0,
                classCount: bucket && Array.isArray(bucket.classes) ? bucket.classes.length : 0,
                subjectCount: Array.isArray(core.subjects) ? core.subjects.length : 0,
                argeCount: Array.isArray(core.arges) ? core.arges.length : 0,
                intranetSiteUrl: String(setup.intranetSiteUrl || '').trim(),
                intranetListTitleCount: Object.keys(listTitles).filter(function (k) {
                    return String(listTitles[k] || '').trim();
                }).length,
                schulOrganigrammNodeCount: nodes.length,
                schulOrganigrammContentCount: nodes.filter(function (n) {
                    return n && String(n.type || '').toLowerCase() !== 'section';
                }).length
            };
        } catch {
            return null;
        }
    }

    /**
     * Vor dem Export: offene Änderungen aus app-data-v2 / Stammdaten in localStorage spiegeln,
     * damit das Backup möglichst vollständig und konsistent ist.
     */
    /**
     * Stichprobe: welche Demo-kritischen Blöcke sind im Export wirklich enthalten?
     * @param {object} local
     */
    function buildCriticalCoverage(local, indexed) {
        const has = function (k) {
            return !!(local && Object.prototype.hasOwnProperty.call(local, k) && local[k] != null && local[k] !== '');
        };
        const schoolApps = [
            { id: 'schularbeiten', label: 'Schularbeiten', keys: ['ms365-sa-site-url', 'ms365-sa-settings-v1'] },
            { id: 'projektwochen', label: 'Projektwochen', keys: ['ms365-pw-site-url', 'ms365-projektwochen-demo-v1'] },
            {
                id: 'schulaktivitaeten',
                label: 'Schulaktivitäten',
                keys: [
                    'ms365-akt-planer-site-v1',
                    'ms365-akt-setup-v1',
                    'ms365-schulaktivitaeten-perms-v1'
                ]
            },
            { id: 'freistellung', label: 'Freistellung', keys: ['ms365-freistellung-setup-v1', 'ms365-freistellung-planer-site-v1'] },
            { id: 'lehrer-freistellung', label: 'Lehrer-Freistellung', keys: ['ms365-lfr-setup-v1', 'ms365-lfr-planer-site-v1'] },
            { id: 'einsammler', label: 'Einsammler', keys: ['ms365-einsammler-v1', 'ms365-einsammler-settings-v1', 'ms365-einsammler-sp-settings-v1'] },
            { id: 'elternbrief', label: 'Elternbrief', keys: ['ms365-elternbrief-v1', 'ms365-elternbrief-settings-v1', 'ms365-elternbrief-sp-settings-v1'] },
            { id: 'krankmeldung', label: 'Krankmeldung', keys: ['ms365-krankmeldung-v1', 'ms365-krankmeldung-settings-v1', 'ms365-krankmeldung-sp-settings-v1'] },
            { id: 'kollegium-umfragen', label: 'Kollegium-Umfragen', keys: ['ms365-kollegium-umfragen-v1', 'ms365-kollegium-umfragen-settings-v1'] },
            { id: 'ausgaben', label: 'Ausgaben', keys: ['ms365-ausgaben-v1', 'ms365-ausgaben-settings-v1', 'ms365-ausgaben-sp-settings-v1'] },
            { id: 'wahlfach', label: 'Wahlfach-Zuteilung', keys: ['ms365-wahlfach-zuteilung-v1', 'ms365-wahlfach-zuteilung-settings-v1'] },
            { id: 'schulsprecherwahl', label: 'Schulsprecherwahl', keys: ['ms365-schulsprecherwahl-v1'] },
            { id: 'teams-antrag', label: 'Teams-Antrag', keys: ['ms365-teams-antrag-v1', 'ms365-teams-antrag-settings-v1'] },
            { id: 'diplomarbeiten', label: 'Diplomarbeiten', keys: ['ms365-diplomarbeiten-state-v1', 'ms365-diplomarbeiten-settings-v1'] },
            { id: 'raumplan', label: 'Raumplan', keys: ['ms365-raumplan-v1', 'ms365-raumplan-sp-settings-v1'] },
            { id: 'aufgaben', label: 'Aufgaben', keys: ['ms365-aufgaben-v1', 'ms365-aufgaben-sp-settings-v1'] },
            { id: 'technik-meldungen', label: 'Technik-Meldungen', keys: ['ms365-technik-meldungen-v1', 'ms365-technik-meldungen-sp-settings-v1'] },
            { id: 'termine-hub', label: 'Termine-Hub', keys: ['ms365-termine-hub-v1', 'ms365-termine-hub-settings-v1', 'ms365-kalender-postfaecher-v1'] },
            { id: 'stammdaten-portal', label: 'Schüler-Stammdaten-Portal', keys: ['ms365-schueler-stammdaten-portal-v1'] }
        ].map(function (app) {
            const presentKeys = app.keys.filter(has);
            return {
                id: app.id,
                label: app.label,
                present: presentKeys.length > 0,
                keys: presentKeys
            };
        });
        const indexedObj = indexed && typeof indexed === 'object' ? indexed : {};
        const indexedPresent = Object.keys(indexedObj).filter(function (name) {
            const row = indexedObj[name];
            return row && (Number(row.recordCount) > 0 || (row.stores && Object.keys(row.stores).length));
        });
        const checks = [
            { id: 'stammdaten', label: 'Stammdaten (v2)', ok: has('ms365-schooltool-data-v2') },
            { id: 'emailPattern', label: 'E-Mail-Muster', ok: has('ms365-tenant-school-email-pattern-v1') || has('ms365-tenant-teachers-email-pattern-v1') },
            { id: 'intranetLinks', label: 'Intranet-Listen-Links', ok: has('ms365-intranet-list-links-v1') },
            { id: 'dashboardAccess', label: 'Dashboard-Zugriff', ok: has('ms365-dashboard-tool-access-v1') },
            {
                id: 'schoolApps',
                label: 'Schul-Apps (mind. eine Site/Einstellung)',
                ok: schoolApps.some(function (a) {
                    return a.present;
                })
            },
            {
                id: 'attachments',
                label: 'Anhänge (IndexedDB)',
                ok: indexedPresent.length > 0,
                optional: true
            }
        ];
        const missing = checks.filter(function (c) {
            return !c.ok && !c.optional;
        });
        return {
            checks: checks,
            schoolApps: schoolApps,
            indexedDbNames: indexedPresent,
            missingIds: missing.map(function (c) {
                return c.id;
            }),
            readyForDemoTransfer: checks[0].ok
        };
    }

    function criticalCoverageSummaryText(coverage) {
        if (!coverage || !coverage.checks) return '';
        const bits = coverage.checks
            .filter(function (c) {
                return !c.optional || c.ok;
            })
            .map(function (c) {
                return c.label + (c.ok ? ' ✓' : ' –');
            });
        const apps = (coverage.schoolApps || [])
            .filter(function (a) {
                return a.present;
            })
            .map(function (a) {
                return a.label;
            });
        if (apps.length) bits.push('Schul-Apps: ' + apps.join(', '));
        return bits.join(' · ');
    }

    function syncStorageBeforeBackup() {
        try {
            try {
                window.dispatchEvent(new CustomEvent('ms365-before-browser-backup', { detail: { at: new Date().toISOString() } }));
            } catch {
                /* ignore */
            }
            if (typeof window.ms365FlushSchulOrganigrammBeforeBackup === 'function') {
                window.ms365FlushSchulOrganigrammBeforeBackup();
            }
            if (typeof window.ms365TenantSettingsFlushPendingSave === 'function') {
                window.ms365TenantSettingsFlushPendingSave();
            }
            if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.getContainer !== 'function') return;
            const c = window.ms365AppDataV2.getContainer();
            if (typeof window.ms365AppDataV2.setContainer === 'function') {
                window.ms365AppDataV2.setContainer(c);
            }
            const core = c && c.core ? c.core : {};
            if (typeof window.ms365SetSchoolDomainNoAt === 'function' && core.domain) {
                window.ms365SetSchoolDomainNoAt(core.domain);
            }
        } catch {
            /* Backup soll trotzdem laufen */
        }
    }

    function buildInventory(local, session, indexed) {
        const localKeys = Object.keys(local || {});
        const sessionKeys = Object.keys(session || {});
        const indexedObj = indexed && typeof indexed === 'object' ? indexed : {};
        const indexedNames = Object.keys(indexedObj);
        const known = new Set();
        STORAGE_CATALOG.forEach(function (cat) {
            cat.keys.forEach(function (k) {
                known.add(k);
            });
        });

        const categories = STORAGE_CATALOG.map(function (cat) {
            const present = cat.keys.filter(function (k) {
                return Object.prototype.hasOwnProperty.call(local, k);
            });
            return {
                label: cat.label,
                keys: present,
                count: present.length
            };
        }).filter(function (c) {
            return c.count > 0;
        });

        const otherLocal = localKeys.filter(function (k) {
            return !known.has(k);
        });
        const otherSession = sessionKeys.filter(function () {
            return true;
        });

        if (otherLocal.length) {
            categories.push({ label: 'Weitere lokale Einträge', keys: otherLocal.sort(), count: otherLocal.length });
        }
        if (otherSession.length) {
            categories.push({
                label: 'Werkzeug-Sitzung (sessionStorage)',
                keys: otherSession.sort(),
                count: otherSession.length
            });
        }
        if (indexedNames.length) {
            categories.push({
                label: 'Anhänge (IndexedDB)',
                keys: indexedNames.sort(),
                count: indexedNames.length
            });
        }

        return {
            categories: categories,
            localKeyCount: localKeys.length,
            sessionKeyCount: sessionKeys.length,
            indexedDbCount: indexedNames.length
        };
    }

    function inventorySummaryText(inventory) {
        if (!inventory || !inventory.categories || !inventory.categories.length) return '';
        return inventory.categories
            .map(function (c) {
                return c.label + ': ' + c.count;
            })
            .join(' · ');
    }

    function postImportNormalize(_storage) {
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.invalidateCache === 'function') {
                window.ms365AppDataV2.invalidateCache();
            }
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function') {
                window.ms365AppDataV2.getContainer();
            }
            if (typeof window.ms365SetSchoolDomainNoAt === 'function' && typeof window.ms365TenantSettingsLoad === 'function') {
                const t = window.ms365TenantSettingsLoad();
                if (t && t.domain) window.ms365SetSchoolDomainNoAt(t.domain);
            }
            window.dispatchEvent(
                new CustomEvent('ms365-tenant-settings-changed', { detail: { source: 'browser-backup-import' } })
            );
        } catch {
            /* ignore */
        }
    }

    function backupFilename(d, storage) {
        const { schoolName, domain } = readSchoolMeta(storage);
        const school = sanitizeForFilename(schoolName || domain);
        const suffix = school ? '-' + school : '';
        return 'ms365-browser-backup-' + localDateTimeStamp(d) + suffix + '.json';
    }

    /**
     * Kurzer Fingerprint aller App-localStorage-Werte (ohne Sync-Meta).
     * Dient dem Auto-Sync: Push nur wenn sich Schul-/Werkzeugdaten geändert haben.
     */
    function contentFingerprint(storage) {
        const local = collectLocalStorage(storage);
        const parts = [];
        Object.keys(local)
            .sort()
            .forEach(function (k) {
                if (FINGERPRINT_SKIP_KEYS[k]) return;
                parts.push(k + '\0' + decodeValue(local[k]));
            });
        const s = parts.join('\n');
        let h = 5381;
        for (let i = 0; i < s.length; i++) {
            h = (Math.imul(h, 33) ^ s.charCodeAt(i)) | 0;
        }
        return (h >>> 0).toString(16) + ':' + parts.length;
    }

    /**
     * Signalisiert, dass sich lokale App-Daten geändert haben (z. B. Power-Automate-Felder).
     * Auto-Sync und UI können darauf reagieren.
     * @param {string} [source]
     * @param {object} [detail]
     */
    function notifyLocalDataChanged(source, detail) {
        try {
            const payload = Object.assign({}, detail || {}, {
                source: String(source || 'local-data'),
                at: new Date().toISOString()
            });
            window.dispatchEvent(new CustomEvent('ms365-local-data-changed', { detail: payload }));
        } catch {
            /* ignore */
        }
    }

    function buildLegacyMirrorMeta(local) {
        try {
            if (
                window.ms365TenantStorageMirror &&
                typeof window.ms365TenantStorageMirror.buildLegacyMirrorExportMeta === 'function'
            ) {
                return window.ms365TenantStorageMirror.buildLegacyMirrorExportMeta(local);
            }
        } catch {
            /* ignore */
        }
        return {
            key: 'ms365-tenant-settings-v1',
            role: 'legacy-read-only',
            writeEnabled: false,
            canonicalKey: 'ms365-schooltool-data-v2',
            sunset: '2026-H2',
            present: !!(local && Object.prototype.hasOwnProperty.call(local, 'ms365-tenant-settings-v1')),
            canonicalPresent: !!(local && Object.prototype.hasOwnProperty.call(local, 'ms365-schooltool-data-v2')),
            note: 'Kanonisch ist ms365-schooltool-data-v2; v1 wird nicht mehr automatisch geschrieben.'
        };
    }

    function finalizeBackupPayload(local, session, indexed, storage, when) {
        const { schoolName, domain } = readSchoolMeta(storage);
        const inventory = buildInventory(local, session, indexed);
        const legacyMirror = buildLegacyMirrorMeta(local);
        const schooltoolSummary = summarizeSchooltoolForBackup(local);
        const criticalCoverage = buildCriticalCoverage(local, indexed);
        const indexedKeyCount = Object.keys(indexed || {}).length;
        const payload = {
            kind: KIND,
            version: VERSION,
            exportedAt: when.toISOString(),
            schoolName: schoolName,
            domain: domain,
            keyCount: Object.keys(local).length + Object.keys(session).length + indexedKeyCount,
            localKeyCount: Object.keys(local).length,
            sessionKeyCount: Object.keys(session).length,
            indexedDbCount: indexedKeyCount,
            inventory: inventory,
            schooltoolSummary: schooltoolSummary,
            criticalCoverage: criticalCoverage,
            contentFingerprint: contentFingerprint(storage),
            includesPrefixes: ['ms365-', 'webuntis-'],
            includesIndexedDb: INDEXED_DB_CATALOG.map(function (e) {
                return e.name;
            }),
            includesNote:
                'Vollständiges App-Backup: Stammdaten (kanonisch v2 inkl. Schul-Organigramm, Schulprofil/Schulkennzahl, Verwaltungs-Zielgruppen, Intranet-Listen-Titel, Einrichtungsstand), E-Mail-Muster, optionaler v1-Spiegel (legacyMirror), Dashboard-Zugriff, alle Schul-Apps-Einstellungen (Sites/Rollen/Listen-Titel), Werkzeugstände, Planer, Power-Automate-/Freistellungs-Konfiguration, Anhänge in IndexedDB sowie alle weiteren ms365-/webuntis-Schlüssel im Browser.',
            excludesNote:
                'Nicht enthalten: Microsoft-Anmeldung (MSAL), PIN-/Admin-Freischaltung in dieser Sitzung, kurzlebige Login-Weiterleitungen sowie die eigentlichen Einträge in SharePoint-Listen (die liegen im Tenant, nicht im Browser).',
            legacyMirror: legacyMirror,
            localStorage: local,
            sessionStorage: session,
            indexedDB: indexed || {}
        };
        let summary = inventorySummaryText(inventory);
        if (schooltoolSummary) {
            const stammdatenBits = [];
            if (schooltoolSummary.schoolCode) stammdatenBits.push('Kennzahl ' + schooltoolSummary.schoolCode);
            if (schooltoolSummary.hasSchoolProfileLogo) stammdatenBits.push('Schullogo');
            if (schooltoolSummary.teacherCount) stammdatenBits.push(schooltoolSummary.teacherCount + ' Lehrkräfte');
            if (schooltoolSummary.classCount) stammdatenBits.push(schooltoolSummary.classCount + ' Klassen');
            if (schooltoolSummary.studentCount) stammdatenBits.push(schooltoolSummary.studentCount + ' Schüler');
            if (schooltoolSummary.subjectCount) stammdatenBits.push(schooltoolSummary.subjectCount + ' Fächer');
            if (stammdatenBits.length) {
                summary = summary
                    ? summary + ' · Stammdaten: ' + stammdatenBits.join(', ')
                    : 'Stammdaten: ' + stammdatenBits.join(', ');
            }
            if (schooltoolSummary.schulOrganigrammNodeCount > 0) {
                const orgBit =
                    'Schul-Organigramm: ' + schooltoolSummary.schulOrganigrammNodeCount + ' Knoten';
                summary = summary ? summary + ' · ' + orgBit : orgBit;
            }
        }
        const idbText = indexedDbSummary(indexed);
        if (idbText) summary = summary ? summary + ' · ' + idbText : idbText;
        const covText = criticalCoverageSummaryText(criticalCoverage);
        if (covText) summary = summary ? summary + ' · ' + covText : covText;
        payload.inventorySummary = summary;
        return payload;
    }

    /**
     * Synchrones Backup (localStorage + sessionStorage). IndexedDB bleibt leer –
     * für den vollständigen Export bitte buildBackupAsync / downloadBackup nutzen.
     */
    function buildBackup(storage, now, sessionStorageArg) {
        syncStorageBeforeBackup();
        const local = collectLocalStorage(storage);
        const session = collectSessionStorage(sessionStorageArg);
        return finalizeBackupPayload(local, session, {}, storage, asDate(now));
    }

    /**
     * Vollständiges Backup inkl. bekannter IndexedDB-Anhänge der Schul-Apps.
     */
    async function buildBackupAsync(storage, now, sessionStorageArg) {
        syncStorageBeforeBackup();
        const local = collectLocalStorage(storage);
        const session = collectSessionStorage(sessionStorageArg);
        let indexed = {};
        try {
            indexed = await collectIndexedDb();
        } catch {
            indexed = {};
        }
        return finalizeBackupPayload(local, session, indexed, storage, asDate(now));
    }

    function isBackupPayload(obj) {
        return !!(
            obj &&
            typeof obj === 'object' &&
            obj.kind === KIND &&
            (obj.version === undefined || obj.version >= 2) &&
            obj.localStorage &&
            typeof obj.localStorage === 'object' &&
            !Array.isArray(obj.localStorage)
        );
    }

    function isLegacyAppDataPayload(obj) {
        return !!(obj && typeof obj === 'object' && obj.version >= 2 && obj.core && obj.structure && obj.match);
    }

    function applyBackup(payload, storage, opts) {
        if (!isBackupPayload(payload)) throw new Error('Kein gültiges Browser-Backup.');
        const store = getStore(storage);
        if (!store) throw new Error('localStorage ist nicht verfügbar.');
        const sessionStore =
            opts && Object.prototype.hasOwnProperty.call(opts, 'sessionStorage')
                ? getSessionStore(opts.sessionStorage)
                : getSessionStore();
        const replace = !opts || opts.replace !== false;
        const incoming = payload.localStorage;
        const incomingSession =
            payload.sessionStorage && typeof payload.sessionStorage === 'object' && !Array.isArray(payload.sessionStorage)
                ? payload.sessionStorage
                : {};
        const written = [];
        const removed = [];
        const sessionWritten = [];
        const sessionRemoved = [];
        const skipped = [];
        const errors = [];

        // Zuerst schreiben, dann verwaiste Keys löschen – bei Teilfehler bleiben lokale Daten erhalten.
        Object.keys(incoming).forEach(function (k) {
            if (!isAppStorageKey(k)) {
                skipped.push(k);
                return;
            }
            try {
                store.setItem(k, decodeValue(incoming[k]));
                written.push(k);
            } catch (e) {
                const msg = e && e.message ? e.message : String(e);
                const quota =
                    (e && e.name === 'QuotaExceededError') || /quota|speicher|exceeded/i.test(msg);
                errors.push(k + ': ' + (quota ? 'Speicherplatz im Browser voll (QuotaExceeded)' : msg));
            }
        });

        if (sessionStore) {
            Object.keys(incomingSession).forEach(function (k) {
                if (!isRestorableSessionKey(k)) {
                    skipped.push(k);
                    return;
                }
                try {
                    sessionStore.setItem(k, decodeValue(incomingSession[k]));
                    sessionWritten.push(k);
                } catch (e) {
                    const msg = e && e.message ? e.message : String(e);
                    const quota =
                        (e && e.name === 'QuotaExceededError') || /quota|speicher|exceeded/i.test(msg);
                    errors.push(k + ': ' + (quota ? 'Speicherplatz im Browser voll (QuotaExceeded)' : msg));
                }
            });
        }

        if (errors.length) {
            const err = new Error(
                'Backup nur teilweise geschrieben: ' +
                    errors.slice(0, 3).join('; ') +
                    '. Tipp: Schul-Logo verkleinern oder in einem leeren Browser-Profil importieren.'
            );
            err.details = {
                written: written,
                removed: removed,
                sessionWritten: sessionWritten,
                sessionRemoved: sessionRemoved,
                skipped: skipped,
                errors: errors
            };
            throw err;
        }

        if (replace) {
            listAppKeys(store).forEach(function (k) {
                if (!Object.prototype.hasOwnProperty.call(incoming, k)) {
                    try {
                        store.removeItem(k);
                        removed.push(k);
                    } catch (e) {
                        errors.push(k + ': ' + (e && e.message ? e.message : String(e)));
                    }
                }
            });
        }

        if (sessionStore && replace) {
            const n = sessionStore.length || 0;
            const sessionKeys = [];
            for (let i = 0; i < n; i++) {
                const k = sessionStore.key(i);
                if (k && isRestorableSessionKey(k)) sessionKeys.push(k);
            }
            sessionKeys.forEach(function (k) {
                if (!Object.prototype.hasOwnProperty.call(incomingSession, k)) {
                    try {
                        sessionStore.removeItem(k);
                        sessionRemoved.push(k);
                    } catch (e) {
                        errors.push(k + ': ' + (e && e.message ? e.message : String(e)));
                    }
                }
            });
        }

        if (errors.length) {
            const err = new Error(
                'Backup nur teilweise geschrieben: ' +
                    errors.slice(0, 3).join('; ') +
                    '. Tipp: Schul-Logo verkleinern oder in einem leeren Browser-Profil importieren.'
            );
            err.details = {
                written: written,
                removed: removed,
                sessionWritten: sessionWritten,
                sessionRemoved: sessionRemoved,
                skipped: skipped,
                errors: errors
            };
            throw err;
        }
        postImportNormalize(storage);
        return {
            written: written,
            removed: removed,
            sessionWritten: sessionWritten,
            sessionRemoved: sessionRemoved,
            skipped: skipped,
            errors: errors,
            indexedDbWritten: [],
            indexedDbCleared: []
        };
    }

    /**
     * Wie applyBackup, zusätzlich IndexedDB-Anhänge (Version ≥ 7).
     */
    async function applyBackupAsync(payload, storage, opts) {
        const result = applyBackup(payload, storage, opts);
        const replace = !opts || opts.replace !== false;
        try {
            const idbResult = await applyIndexedDb(payload && payload.indexedDB, { replace: replace });
            result.indexedDbWritten = idbResult.written || [];
            result.indexedDbCleared = idbResult.cleared || [];
            if (idbResult.errors && idbResult.errors.length) {
                result.errors = (result.errors || []).concat(idbResult.errors);
            }
        } catch (e) {
            result.errors = (result.errors || []).concat([
                'indexedDB: ' + (e && e.message ? e.message : String(e))
            ]);
        }
        if (result.errors && result.errors.length) {
            const err = new Error(
                'Backup nur teilweise geschrieben: ' +
                    result.errors.slice(0, 3).join('; ') +
                    '. Tipp: Schul-Logo verkleinern oder in einem leeren Browser-Profil importieren.'
            );
            err.details = result;
            throw err;
        }
        return result;
    }

    async function importPayload(obj, storage) {
        if (!obj || typeof obj !== 'object') throw new Error('Keine gültige JSON-Datei.');
        if (isBackupPayload(obj)) {
            return { mode: 'backup', result: await applyBackupAsync(obj, storage) };
        }
        if (isLegacyAppDataPayload(obj) && window.ms365AppDataV2 && typeof window.ms365AppDataV2.importJson === 'function') {
            window.ms365AppDataV2.importJson(obj);
            return { mode: 'app-data-v2' };
        }
        const looksLikeTenant =
            Object.prototype.hasOwnProperty.call(obj, 'domain') ||
            Array.isArray(obj.subjects) ||
            Array.isArray(obj.teachers);
        if (looksLikeTenant) {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.importJson === 'function') {
                window.ms365AppDataV2.importJson(obj);
            }
            if (typeof window.ms365TenantSettingsSave === 'function') {
                window.ms365TenantSettingsSave(obj);
            }
            return { mode: 'tenant-v1' };
        }
        throw new Error('Unbekanntes JSON-Format. Erwartet wird ein Browser-Backup oder ein Schuldaten-Export.');
    }

    function downloadJson(filename, obj) {
        const blob = new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        // Programmatischer a.click() bubbelt bis document – sonst schließt der
        // Datensicherungs-Dismiss-Handler das Panel und der Export wirkt „tot“.
        a.addEventListener('click', function (ev) {
            ev.stopPropagation();
        });
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(function () {
            URL.revokeObjectURL(url);
        }, 250);
    }

    async function downloadBackup(storage, now) {
        const payload = await buildBackupAsync(storage, now);
        downloadJson(backupFilename(now, storage), payload);
        return payload;
    }

    function dlgConfirm(msg, opts) {
        if (typeof window.ms365AppDialogConfirm === 'function') {
            return window.ms365AppDialogConfirm(msg, opts);
        }
        return Promise.resolve(window.confirm(msg));
    }

    function dlgAlert(msg, opts) {
        if (typeof window.ms365AppDialogAlert === 'function') {
            return window.ms365AppDialogAlert(msg, opts);
        }
        window.alert(msg);
        return Promise.resolve();
    }

    function setStatus(text, kind) {
        const els = [
            document.getElementById('browserBackupStatus'),
            document.getElementById('ms365BackupPanelStatus')
        ].filter(Boolean);
        els.forEach(function (el) {
            el.textContent = text || '';
            if (el.id === 'browserBackupStatus') {
                el.style.display = text ? 'block' : 'none';
            }
            el.classList.toggle('ok', kind === 'ok');
            el.classList.toggle('warn', kind === 'warn');
            if (el.id === 'ms365BackupPanelStatus') {
                el.classList.toggle('ms365-backup-panel__status--ok', kind === 'ok');
                el.classList.toggle('ms365-backup-panel__status--warn', kind === 'warn');
            }
            if (kind) el.setAttribute('data-kind', kind);
            else el.removeAttribute('data-kind');
        });
        try {
            window.dispatchEvent(
                new CustomEvent('ms365-browser-backup-status', {
                    detail: { text: text || '', kind: kind || '' }
                })
            );
        } catch {
            /* ignore */
        }
    }

    function confirmImportMessage(obj) {
        if (isBackupPayload(obj)) {
            const n = Object.keys(obj.localStorage || {}).length;
            const s = Object.keys(obj.sessionStorage || {}).length;
            const when = obj.exportedAt ? String(obj.exportedAt).replace('T', ' ').replace(/\.\d+Z$/, ' UTC') : '';
            const school = String(obj.schoolName || obj.domain || '').trim();
            const summary =
                obj.inventorySummary ||
                inventorySummaryText(
                    obj.inventory ||
                        buildInventory(obj.localStorage || {}, obj.sessionStorage || {})
                );
            const idb = Object.keys(obj.indexedDB || {}).length;
            let msg =
                'Dieses Browser-Backup enthält ' +
                n +
                ' lokale und ' +
                s +
                ' Sitzungs-Einträge' +
                (idb ? ' sowie ' + idb + ' Anhang-Datenbanken' : '') +
                (when ? ' (Stand: ' + when + ')' : '') +
                (school ? ' für „' + school + '“' : '') +
                '.\n\n';
            if (summary) {
                msg += 'Enthalten u. a.: ' + summary + '.\n\n';
            }
            msg +=
                'Alle lokalen Schuldaten, Werkzeug-Zwischenstände und Anhänge in diesem Browser werden ersetzt. ' +
                'Gruppen und Benutzer in Microsoft 365 bleiben unverändert. ' +
                'Microsoft-Anmeldung und PIN-Freischaltung bleiben unberührt. Fortfahren?';
            return subjectCatalogShrinkAppend(obj, msg);
        }
        return subjectCatalogShrinkAppend(
            obj,
            'Diese JSON-Datei enthält nur Schul-/Stammdaten – kein vollständiges Browser-Backup.\n\n' +
                'Nicht übernommen werden dabei u. a.: E-Mail-Muster, Schul-Apps-Einstellungen (Schularbeiten, Projektwochen, Freistellung, Einsammler, Elternbrief, …), Dashboard-Zugriff, Power-Automate-Konfiguration, Anhänge und Werkzeug-Zwischenstände.\n\n' +
                'Für die Übergabe an einen anderen PC bitte oben „Browser-Backup“ / Datensicherung exportieren (kind: ms365-browser-backup-v1).\n\n' +
                'Vorhandene Stammdaten in diesem Browser werden überschrieben. Fortfahren?'
        );
    }

    function subjectCatalogShrinkAppend(obj, baseMsg) {
        try {
            if (
                typeof window.ms365SubjectCatalogShrinkReport !== 'function' ||
                typeof window.ms365SubjectsFromImportPayload !== 'function' ||
                typeof window.ms365FormatSubjectCatalogShrinkConfirm !== 'function' ||
                typeof window.ms365TenantSettingsLoad !== 'function'
            ) {
                return baseMsg;
            }
            const next = window.ms365SubjectsFromImportPayload(obj);
            if (!next.length) return baseMsg;
            const cur = window.ms365TenantSettingsLoad().subjects || [];
            const rep = window.ms365SubjectCatalogShrinkReport(cur, next);
            if (!rep.wouldShrink) return baseMsg;
            const extra = window.ms365FormatSubjectCatalogShrinkConfirm(rep, '');
            return extra ? baseMsg + '\n\n' + extra : baseMsg;
        } catch {
            return baseMsg;
        }
    }

    async function importFile(file, opts) {
        const reload = !opts || opts.reload !== false;
        const text = await file.text();
        let obj;
        try {
            obj = JSON.parse(text);
        } catch {
            throw new Error('Keine gültige JSON-Datei.');
        }
        const ok = await dlgConfirm(confirmImportMessage(obj), {
            title: 'Backup importieren',
            okText: 'Importieren',
            cancelText: 'Abbrechen',
            danger: true
        });
        if (!ok) return { cancelled: true };
        const imported = await importPayload(obj);
        if (reload) {
            location.reload();
        }
        return imported;
    }

    function clearRestorableSessionKeys(sessionStorageArg) {
        const sessionStore = getSessionStore(sessionStorageArg);
        const removed = [];
        if (!sessionStore || typeof sessionStore.key !== 'function') return removed;
        const n = sessionStore.length || 0;
        const sessionKeys = [];
        for (let i = 0; i < n; i++) {
            const k = sessionStore.key(i);
            if (k && isRestorableSessionKey(k)) sessionKeys.push(k);
        }
        sessionKeys.forEach(function (k) {
            try {
                sessionStore.removeItem(k);
                removed.push(k);
            } catch {
                /* ignore */
            }
        });
        return removed;
    }

    /**
     * Löscht alle lokalen App-Daten (ms365-* / webuntis-* + bekannte IndexedDBs), nicht MSAL.
     */
    async function clearAllAppData(opts) {
        const store = getStore(opts && opts.storage);
        if (!store) throw new Error('localStorage ist nicht verfügbar.');
        const removed = [];
        const errors = [];
        listAppKeys(store).forEach(function (k) {
            try {
                store.removeItem(k);
                removed.push(k);
            } catch (e) {
                errors.push(k + ': ' + (e && e.message ? e.message : String(e)));
            }
        });
        const sessionRemoved = clearRestorableSessionKeys(opts && opts.sessionStorage);
        let indexedDbCleared = [];
        try {
            indexedDbCleared = await clearIndexedDbCatalog();
        } catch (e) {
            errors.push('indexedDB: ' + (e && e.message ? e.message : String(e)));
        }
        if (errors.length) {
            throw new Error(errors.slice(0, 3).join('; '));
        }
        return {
            removed: removed.length,
            keys: removed,
            sessionRemoved: sessionRemoved.length,
            indexedDbCleared: indexedDbCleared
        };
    }

    async function resetAllAppData(reload) {
        const ok = await dlgConfirm(
            'Alle lokalen App-Daten in diesem Browser löschen?\n\n' +
                'Entfernt werden: Stammdaten, Einrichtungsstand, Werkzeug-Zwischenstände, Anhänge und Favoriten. ' +
                'Nicht gelöscht: Ihre Microsoft-Anmeldung (oben rechts) und der PIN-Zugang dieser Seite.\n\n' +
                'Tipp: Vorher „Browser-Backup“ exportieren, falls Sie Daten behalten möchten.',
            {
                title: 'Alles zurücksetzen',
                okText: 'Alles löschen',
                cancelText: 'Abbrechen',
                danger: true
            }
        );
        if (!ok) return { cancelled: true };
        const result = await clearAllAppData();
        if (reload !== false) {
            location.reload();
        }
        return result;
    }

    function onResetClick() {
        resetAllAppData(true).catch(function (e) {
            setStatus('Zurücksetzen fehlgeschlagen: ' + (e && e.message ? e.message : String(e)), 'warn');
            return dlgAlert('Zurücksetzen fehlgeschlagen: ' + (e && e.message ? e.message : String(e)), {
                title: 'Alles zurücksetzen'
            });
        });
    }

    function onExportClick(ev) {
        if (ev && typeof ev.preventDefault === 'function') ev.preventDefault();
        if (ev && typeof ev.stopPropagation === 'function') ev.stopPropagation();
        Promise.resolve()
            .then(function () {
                if (typeof syncStorageBeforeBackup === 'function') {
                    try {
                        syncStorageBeforeBackup();
                    } catch {
                        /* weiter mit aktuellem Storage */
                    }
                }
                return downloadBackup();
            })
            .then(function (payload) {
                let meaningful = false;
                try {
                    if (
                        window.ms365EmptyStateUi &&
                        typeof window.ms365EmptyStateUi.hasMeaningfulTenantData === 'function' &&
                        typeof window.ms365TenantSettingsLoad === 'function'
                    ) {
                        meaningful = window.ms365EmptyStateUi.hasMeaningfulTenantData(
                            window.ms365TenantSettingsLoad()
                        );
                    } else if (payload.schoolName || payload.domain) {
                        meaningful = true;
                    } else if (payload.inventorySummary) {
                        meaningful = /Stammdaten|Lehrer|Schüler|Klassen/i.test(String(payload.inventorySummary));
                    }
                } catch {
                    meaningful = !!(payload.schoolName || payload.domain);
                }
                const parts = [payload.keyCount + ' Einträge'];
                if (payload.localKeyCount != null) {
                    parts[0] = payload.localKeyCount + ' lokal';
                    if (payload.sessionKeyCount) parts.push(payload.sessionKeyCount + ' Sitzung');
                    if (payload.indexedDbCount) parts.push(payload.indexedDbCount + ' IndexedDB');
                }
                try {
                    localStorage.setItem('ms365-last-backup-export-at', new Date().toISOString());
                } catch {
                    /* ignore */
                }
                let status = 'Vollständiges Browser-Backup gespeichert: ' + parts.join(', ') + '.';
                if (payload.inventorySummary) status += ' ' + payload.inventorySummary + '.';
                if (payload.criticalCoverage && !payload.criticalCoverage.readyForDemoTransfer) {
                    status +=
                        ' Achtung: Stammdaten (ms365-schooltool-data-v2) fehlen – Demo-Übergabe wäre unvollständig.';
                    setStatus(status, 'warn');
                } else if (!meaningful) {
                    status +=
                        ' Hinweis: Kaum Stammdaten enthalten – für die Übergabe an einen anderen Browser zuerst Einrichtung oder Import.';
                    setStatus(status, 'warn');
                } else {
                    status +=
                        ' Datei sicher ablegen. Morgen: dieselbe Datei importieren, dann neu bei Microsoft anmelden.';
                    setStatus(status, 'ok');
                }
                try {
                    window.dispatchEvent(
                        new CustomEvent('ms365-browser-backup-status', {
                            detail: { text: status, kind: meaningful ? 'ok' : 'warn' }
                        })
                    );
                } catch {
                    /* ignore */
                }
            })
            .catch(function (e) {
                setStatus('Export fehlgeschlagen: ' + (e && e.message ? e.message : String(e)), 'warn');
            });
    }

    function refreshBackupReminder() {
        const el = document.getElementById('browserBackupStatus');
        if (!el || el.textContent) return;
        let hasData = false;
        try {
            if (window.ms365EmptyStateUi && typeof window.ms365EmptyStateUi.hasMeaningfulTenantData === 'function') {
                const load =
                    typeof window.ms365TenantSettingsLoad === 'function'
                        ? window.ms365TenantSettingsLoad()
                        : null;
                hasData = window.ms365EmptyStateUi.hasMeaningfulTenantData(load);
            }
        } catch {
            /* ignore */
        }
        if (!hasData) return;
        let last = '';
        try {
            last = localStorage.getItem('ms365-last-backup-export-at') || '';
        } catch {
            /* ignore */
        }
        if (!last) {
            setStatus(
                'Erinnerung: Stammdaten liegen nur in diesem Browser. Nach der Einrichtung ein Browser-Backup exportieren und sicher ablegen.',
                'warn'
            );
            return;
        }
        const ts = Date.parse(last);
        if (!isNaN(ts) && Date.now() - ts > 14 * 24 * 60 * 60 * 1000) {
            setStatus(
                'Letztes Browser-Backup ist älter als 14 Tage. Vor größeren Änderungen erneut exportieren.',
                'warn'
            );
        }
    }

    function bindImportInput(fileImport) {
        if (!fileImport || fileImport.dataset.ms365BackupBound === '1') return;
        fileImport.dataset.ms365BackupBound = '1';
        fileImport.addEventListener('change', function (e) {
            const f = e.target.files && e.target.files[0];
            if (!f) return;
            importFile(f)
                .then(function (res) {
                    if (res && res.cancelled) setStatus('Import abgebrochen.', 'warn');
                })
                .catch(function (err) {
                    setStatus('Import fehlgeschlagen: ' + (err && err.message ? err.message : String(err)), 'warn');
                    return dlgAlert('Import fehlgeschlagen: ' + (err && err.message ? err.message : String(err)), {
                        title: 'Backup importieren'
                    });
                })
                .finally(function () {
                    fileImport.value = '';
                });
        });
    }

    function resolveImportFileInput(fromEl) {
        if (!fromEl) return document.getElementById('browserBackupImportFile');
        const panel = fromEl.closest('#ms365BackupPanel, #ms365BackupHeader');
        if (panel) {
            const panelInput = panel.querySelector('[data-ms365-backup="import-file"], #ms365BackupPanelImportFile');
            if (panelInput) return panelInput;
        }
        return (
            document.getElementById('ms365BackupPanelImportFile') ||
            document.getElementById('browserBackupImportFile')
        );
    }

    function onImportClick(ev) {
        if (ev) {
            ev.preventDefault();
            ev.stopPropagation();
        }
        const input = resolveImportFileInput(ev && ev.currentTarget);
        if (!input) {
            setStatus('Import nicht verfügbar (Dateiauswahl fehlt).', 'warn');
            return;
        }
        bindImportInput(input);
        try {
            input.click();
        } catch (e) {
            setStatus('Import fehlgeschlagen: ' + (e && e.message ? e.message : String(e)), 'warn');
        }
    }

    function bindUi() {
        // Direkte Buttons (Stammdaten-Seite u. a.)
        const exportBtns = document.querySelectorAll('#browserBackupExport, [data-ms365-backup="export"]');
        exportBtns.forEach(function (btn) {
            if (!btn || btn.dataset.ms365BackupBound === '1') return;
            btn.dataset.ms365BackupBound = '1';
            btn.addEventListener('click', onExportClick);
        });
        document.querySelectorAll('[data-ms365-backup="import"]').forEach(function (btn) {
            if (!btn || btn.dataset.ms365BackupBound === '1') return;
            btn.dataset.ms365BackupBound = '1';
            btn.addEventListener('click', onImportClick);
        });
        const fileImport = document.getElementById('browserBackupImportFile');
        bindImportInput(fileImport);
        document.querySelectorAll('[data-ms365-backup="import-file"]').forEach(bindImportInput);
        document.querySelectorAll('[data-ms365-backup="reset"]').forEach(function (btn) {
            if (!btn || btn.dataset.ms365BackupBound === '1') return;
            btn.dataset.ms365BackupBound = '1';
            btn.addEventListener('click', onResetClick);
        });
        // Event-Delegation: dynamisch erzeugte Kopfzeilen-Buttons (nach MSAL) immer erreichbar
        if (!bindUi._delegated) {
            bindUi._delegated = true;
            document.addEventListener(
                'click',
                function (ev) {
                    const t = ev.target;
                    if (!t || typeof t.closest !== 'function') return;
                    const exportBtn = t.closest('[data-ms365-backup="export"], #browserBackupExport');
                    if (exportBtn) {
                        // Schon direkt gebunden → nicht doppelt
                        if (exportBtn.dataset.ms365BackupBound === '1') return;
                        ev.preventDefault();
                        ev.stopPropagation();
                        onExportClick(ev);
                        return;
                    }
                    const importBtn = t.closest('[data-ms365-backup="import"]');
                    if (importBtn) {
                        if (importBtn.dataset.ms365BackupBound === '1') return;
                        onImportClick(ev);
                        return;
                    }
                    const resetBtn = t.closest('[data-ms365-backup="reset"]');
                    if (resetBtn) {
                        if (resetBtn.dataset.ms365BackupBound === '1') return;
                        onResetClick(ev);
                    }
                },
                true
            );
        }
        refreshBackupReminder();
    }

    function mergeLocalStoragePatch(keysObj, storage) {
        const patch = keysObj && typeof keysObj === 'object' ? keysObj : {};
        return applyBackup(
            {
                kind: KIND,
                version: VERSION,
                localStorage: patch,
                sessionStorage: {}
            },
            storage,
            { replace: false }
        );
    }

    window.ms365BrowserBackup = {
        KIND: KIND,
        VERSION: VERSION,
        STORAGE_CATALOG: STORAGE_CATALOG,
        INDEXED_DB_CATALOG: INDEXED_DB_CATALOG,
        isAuthKey: isAuthKey,
        isAppStorageKey: isAppStorageKey,
        isBackupPayload: isBackupPayload,
        isLegacyAppDataPayload: isLegacyAppDataPayload,
        listAppKeys: listAppKeys,
        collectLocalStorage: collectLocalStorage,
        collectSessionStorage: collectSessionStorage,
        collectIndexedDb: collectIndexedDb,
        contentFingerprint: contentFingerprint,
        notifyLocalDataChanged: notifyLocalDataChanged,
        buildBackup: buildBackup,
        buildBackupAsync: buildBackupAsync,
        syncStorageBeforeBackup: syncStorageBeforeBackup,
        buildInventory: buildInventory,
        buildCriticalCoverage: buildCriticalCoverage,
        postImportNormalize: postImportNormalize,
        applyBackup: applyBackup,
        applyBackupAsync: applyBackupAsync,
        applyIndexedDb: applyIndexedDb,
        importPayload: importPayload,
        mergeLocalStoragePatch: mergeLocalStoragePatch,
        backupFilename: backupFilename,
        downloadBackup: downloadBackup,
        importFile: importFile,
        clearAllAppData: clearAllAppData,
        resetAllAppData: resetAllAppData,
        bindUi: bindUi
    };

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', bindUi);
    } else {
        bindUi();
    }
})();
