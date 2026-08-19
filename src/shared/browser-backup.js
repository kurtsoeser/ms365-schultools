/**
 * Vollständiges lokales Browser-Backup (localStorage + wiederherstellbare
 * sessionStorage-Daten der App).
 * Zum Übertragen zwischen Browsern/PCs. Microsoft-Anmeldung (MSAL) bleibt
 * bewusst außen vor – im Zielbrowser erneut anmelden.
 */
(function () {
    'use strict';

    const KIND = 'ms365-browser-backup-v1';
    const VERSION = 3;
    const SESSION_SKIP_KEYS = {
        'ms365-access-granted-v1': true,
        'ms365-admin-access-granted-v1': true,
        'ms365-post-login-url': true
    };

    /** Bekannte Schlüssel – dient der Inventar-Anzeige (Export umfasst alle ms365-* / webuntis-*). */
    const STORAGE_CATALOG = [
        { label: 'Zentrale Schuldaten', keys: ['ms365-schooltool-data-v2', 'ms365-tenant-settings-v1', 'ms365-school-email-domain-v1'] },
        { label: 'Einrichtung & Demo', keys: ['ms365-demo-mode-v1', 'ms365-onboarding-welcome-v1'] },
        { label: 'Dashboard', keys: ['ms365-dashboard-favorites-v1', 'ms365-dashboard-order-catalog-v1', 'ms365-dashboard-category-tab-v1'] },
        { label: 'Schulstruktur (Legacy-Spiegel)', keys: ['ms365-schulstruktur-sync-v1', 'ms365-schulstruktur-match-v1', 'ms365-schulstruktur-tenant-cache-v1'] },
        { label: 'Kursteams / WebUntis', keys: ['webuntis-teams-creator-state-v1'] },
        { label: 'Klassen & Jahrgang', keys: ['ms365-jahrgang-state-v1'] },
        { label: 'Gruppen & SLG', keys: ['ms365-schueler-lehrer-gruppen-v1', 'ms365-schueler-lehrer-gruppen-v2', 'ms365-verwaltung-gruppe-v1'] },
        { label: 'Fächer / ARGE / Weitere Teams', keys: ['ms365-arge-state-v1', 'ms365-arge-state-v2', 'ms365-wtg-state-v1'] },
        { label: 'Personen, Gäste, Hygiene', keys: ['ms365-hygiene-scan-v2', 'ms365-gast-einlader-policy-v1', 'ms365-gast-zugaenge-snapshot-v1', 'ms365-gruppenerstellung-policy-v1'] },
        { label: 'Verteiler & UI', keys: ['ms365-verteilerlisten-cache-v1', 'ms365-theme-v1', 'ms365-ss-graph-collapsed-v1'] },
        { label: 'Admin & Hinweise', keys: ['ms365-schooltool-access-override-v1', 'ms365-schooltool-release-notes-v1', 'ms365-schooltool-release-notes-last-seen-at-v1'] }
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
        return !SESSION_SKIP_KEYS[k];
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

    function readDemoModeFlag(storage) {
        try {
            const store = getStore(storage);
            return !!(store && store.getItem('ms365-demo-mode-v1') === '1');
        } catch {
            return false;
        }
    }

    /**
     * Vor dem Export: offene Änderungen aus app-data-v2 / Stammdaten in localStorage spiegeln,
     * damit das Backup möglichst vollständig und konsistent ist.
     */
    function syncStorageBeforeBackup() {
        try {
            if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.getContainer !== 'function') return;
            const c = window.ms365AppDataV2.getContainer();
            if (typeof window.ms365AppDataV2.setContainer === 'function') {
                window.ms365AppDataV2.setContainer(c);
            }
            const core = c && c.core ? c.core : {};
            const curYear = c && c.years ? String(c.years.current || '').trim() : '';
            const bucket =
                curYear && c.years.byLabel && c.years.byLabel[curYear] ? c.years.byLabel[curYear] : null;
            if (typeof window.ms365TenantSettingsSave === 'function') {
                window.ms365TenantSettingsSave({
                    schoolName: core.schoolName,
                    domain: core.domain,
                    subjects: core.subjects,
                    arges: core.arges,
                    teachers: core.teachers,
                    administration: core.administration,
                    admin: core.admin,
                    adminRoles: core.adminRoles,
                    sgaMode: core.sgaMode,
                    sga: core.sga,
                    students: bucket && bucket.students ? bucket.students : [],
                    studentCouncil: bucket && bucket.studentCouncil ? bucket.studentCouncil : [],
                    classes: bucket && bucket.classes ? bucket.classes : []
                });
            }
            if (typeof window.ms365SetSchoolDomainNoAt === 'function' && core.domain) {
                window.ms365SetSchoolDomainNoAt(core.domain);
            }
        } catch {
            /* Backup soll trotzdem laufen */
        }
    }

    function buildInventory(local, session) {
        const localKeys = Object.keys(local || {});
        const sessionKeys = Object.keys(session || {});
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

        return {
            categories: categories,
            localKeyCount: localKeys.length,
            sessionKeyCount: sessionKeys.length
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

    function postImportNormalize(storage) {
        try {
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
            window.dispatchEvent(
                new CustomEvent('ms365-demo-mode-changed', {
                    detail: { active: readDemoModeFlag(storage) }
                })
            );
        } catch {
            /* ignore */
        }
    }

    function backupFilename(d, storage) {
        const { schoolName, domain } = readSchoolMeta(storage);
        const school = sanitizeForFilename(schoolName || domain);
        const suffix = school ? '-' + school : '';
        return 'ms365-browser-backup-' + localDateStamp(d) + suffix + '.json';
    }

    function buildBackup(storage, now, sessionStorageArg) {
        syncStorageBeforeBackup();
        const local = collectLocalStorage(storage);
        const session = collectSessionStorage(sessionStorageArg);
        const when = asDate(now);
        const { schoolName, domain } = readSchoolMeta(storage);
        const inventory = buildInventory(local, session);
        const payload = {
            kind: KIND,
            version: VERSION,
            exportedAt: when.toISOString(),
            schoolName: schoolName,
            domain: domain,
            demoModeActive: readDemoModeFlag(storage),
            keyCount: Object.keys(local).length + Object.keys(session).length,
            localKeyCount: Object.keys(local).length,
            sessionKeyCount: Object.keys(session).length,
            inventory: inventory,
            includesPrefixes: ['ms365-', 'webuntis-'],
            excludesNote:
                'Nicht enthalten: Microsoft-Anmeldung (MSAL), PIN-/Admin-Freischaltung in dieser Sitzung und kurzlebige Login-Weiterleitungen.',
            localStorage: local,
            sessionStorage: session
        };
        payload.inventorySummary = inventorySummaryText(inventory);
        return payload;
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

        Object.keys(incoming).forEach(function (k) {
            if (!isAppStorageKey(k)) {
                skipped.push(k);
                return;
            }
            try {
                store.setItem(k, decodeValue(incoming[k]));
                written.push(k);
            } catch (e) {
                errors.push(k + ': ' + (e && e.message ? e.message : String(e)));
            }
        });

        if (sessionStore) {
            if (replace) {
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
            Object.keys(incomingSession).forEach(function (k) {
                if (!isRestorableSessionKey(k)) {
                    skipped.push(k);
                    return;
                }
                try {
                    sessionStore.setItem(k, decodeValue(incomingSession[k]));
                    sessionWritten.push(k);
                } catch (e) {
                    errors.push(k + ': ' + (e && e.message ? e.message : String(e)));
                }
            });
        }

        if (errors.length) {
            const err = new Error('Backup nur teilweise geschrieben: ' + errors.slice(0, 3).join('; '));
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
            errors: errors
        };
    }

    function importPayload(obj, storage) {
        if (!obj || typeof obj !== 'object') throw new Error('Keine gültige JSON-Datei.');
        if (isBackupPayload(obj)) {
            return { mode: 'backup', result: applyBackup(obj, storage) };
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
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(function () {
            URL.revokeObjectURL(url);
        }, 250);
    }

    function downloadBackup(storage, now) {
        const payload = buildBackup(storage, now);
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
        const el = document.getElementById('browserBackupStatus');
        if (!el) return;
        el.textContent = text || '';
        el.style.display = text ? 'block' : 'none';
        el.classList.toggle('ok', kind === 'ok');
        el.classList.toggle('warn', kind === 'warn');
        if (kind) el.setAttribute('data-kind', kind);
        else el.removeAttribute('data-kind');
    }

    function confirmImportMessage(obj) {
        if (isBackupPayload(obj)) {
            const n = Object.keys(obj.localStorage || {}).length;
            const s = Object.keys(obj.sessionStorage || {}).length;
            const when = obj.exportedAt ? String(obj.exportedAt).replace('T', ' ').replace(/\.\d+Z$/, ' UTC') : '';
            const school = String(obj.schoolName || obj.domain || '').trim();
            const demo = obj.demoModeActive ? ' (Demo-Modus war aktiv)' : '';
            const summary =
                obj.inventorySummary ||
                inventorySummaryText(
                    obj.inventory ||
                        buildInventory(obj.localStorage || {}, obj.sessionStorage || {})
                );
            let msg =
                'Dieses Browser-Backup enthält ' +
                n +
                ' lokale und ' +
                s +
                ' Sitzungs-Einträge' +
                (when ? ' (Stand: ' + when + ')' : '') +
                (school ? ' für „' + school + '“' : '') +
                demo +
                '.\n\n';
            if (summary) {
                msg += 'Enthalten u. a.: ' + summary + '.\n\n';
            }
            msg +=
                'Alle lokalen Schuldaten und Werkzeug-Zwischenstände in diesem Browser werden ersetzt. ' +
                'Microsoft-Anmeldung und PIN-Freischaltung bleiben unberührt. Fortfahren?';
            return msg;
        }
        return 'Diese JSON-Datei enthält Schuldaten (kein vollständiges Browser-Backup). Vorhandene Stammdaten in diesem Browser werden überschrieben. Fortfahren?';
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
        const imported = importPayload(obj);
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
     * Löscht alle lokalen App-Daten (ms365-* / webuntis-*), nicht die Microsoft-Anmeldung (MSAL).
     */
    function clearAllAppData(opts) {
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
        if (errors.length) {
            throw new Error(errors.slice(0, 3).join('; '));
        }
        try {
            window.dispatchEvent(new CustomEvent('ms365-demo-mode-changed', { detail: { active: false } }));
        } catch {
            /* ignore */
        }
        return {
            removed: removed.length,
            keys: removed,
            sessionRemoved: sessionRemoved.length
        };
    }

    async function resetAllAppData(reload) {
        const ok = await dlgConfirm(
            'Alle lokalen App-Daten in diesem Browser löschen?\n\n' +
                'Entfernt werden: Stammdaten, Einrichtungsstand, Werkzeug-Zwischenstände, Demo-Modus und Favoriten. ' +
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
        const result = clearAllAppData();
        if (reload !== false) {
            location.reload();
        }
        return result;
    }

    async function resetAndLoadDemo(reload) {
        const ok = await dlgConfirm(
            'Demo-Daten der MS365 Musterschule laden?\n\n' +
                'Zuerst werden alle lokalen App-Daten in diesem Browser gelöscht, danach die umfangreiche Muster-Schule ' +
                '(Stammdaten, Verknüpfungen, Beispiel-Schüler:innen, Eltern, Kursteams-Zwischenstand).\n\n' +
                'Ihre Microsoft-Anmeldung bleibt erhalten.',
            {
                title: 'Demo-Daten laden',
                okText: 'Demo laden',
                cancelText: 'Abbrechen'
            }
        );
        if (!ok) return { cancelled: true };
        clearAllAppData({ reload: false });
        if (!window.ms365DemoMode || typeof window.ms365DemoMode.activate !== 'function') {
            throw new Error('Demo-Modul nicht geladen.');
        }
        if (!window.ms365DemoMode.activate()) {
            throw new Error('Demo-Daten konnten nicht geladen werden.');
        }
        if (reload !== false) {
            location.reload();
        }
        return { demo: true };
    }

    function onResetClick() {
        resetAllAppData(true).catch(function (e) {
            setStatus('Zurücksetzen fehlgeschlagen: ' + (e && e.message ? e.message : String(e)), 'warn');
            return dlgAlert('Zurücksetzen fehlgeschlagen: ' + (e && e.message ? e.message : String(e)), {
                title: 'Alles zurücksetzen'
            });
        });
    }

    function onDemoLoadClick() {
        resetAndLoadDemo(true)
            .then(function (res) {
                if (res && res.cancelled) setStatus('Demo-Laden abgebrochen.', 'warn');
            })
            .catch(function (e) {
                setStatus('Demo-Laden fehlgeschlagen: ' + (e && e.message ? e.message : String(e)), 'warn');
                return dlgAlert('Demo-Laden fehlgeschlagen: ' + (e && e.message ? e.message : String(e)), {
                    title: 'Demo-Daten laden'
                });
            });
    }

    function onExportClick() {
        try {
            const payload = downloadBackup();
            const parts = [payload.keyCount + ' Einträge'];
            if (payload.localKeyCount != null) {
                parts[0] = payload.localKeyCount + ' lokal';
                if (payload.sessionKeyCount) parts.push(payload.sessionKeyCount + ' Sitzung');
            }
            let status = 'Backup gespeichert: ' + parts.join(', ') + '.';
            if (payload.inventorySummary) status += ' ' + payload.inventorySummary + '.';
            setStatus(status, 'ok');
        } catch (e) {
            setStatus('Export fehlgeschlagen: ' + (e && e.message ? e.message : String(e)), 'warn');
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

    function bindUi() {
        const exportBtns = document.querySelectorAll('#browserBackupExport, [data-ms365-backup="export"]');
        exportBtns.forEach(function (btn) {
            if (!btn || btn.dataset.ms365BackupBound === '1') return;
            btn.dataset.ms365BackupBound = '1';
            btn.addEventListener('click', onExportClick);
        });
        const fileImport = document.getElementById('browserBackupImportFile');
        bindImportInput(fileImport);
        document.querySelectorAll('[data-ms365-backup="import-file"]').forEach(bindImportInput);
        document.querySelectorAll('[data-ms365-backup="reset"]').forEach(function (btn) {
            if (!btn || btn.dataset.ms365BackupBound === '1') return;
            btn.dataset.ms365BackupBound = '1';
            btn.addEventListener('click', onResetClick);
        });
        document.querySelectorAll('[data-ms365-backup="demo"]').forEach(function (btn) {
            if (!btn || btn.dataset.ms365BackupBound === '1') return;
            btn.dataset.ms365BackupBound = '1';
            btn.addEventListener('click', onDemoLoadClick);
        });
    }

    window.ms365BrowserBackup = {
        KIND: KIND,
        VERSION: VERSION,
        isAuthKey: isAuthKey,
        isAppStorageKey: isAppStorageKey,
        isBackupPayload: isBackupPayload,
        isLegacyAppDataPayload: isLegacyAppDataPayload,
        listAppKeys: listAppKeys,
        collectLocalStorage: collectLocalStorage,
        collectSessionStorage: collectSessionStorage,
        buildBackup: buildBackup,
        syncStorageBeforeBackup: syncStorageBeforeBackup,
        buildInventory: buildInventory,
        postImportNormalize: postImportNormalize,
        applyBackup: applyBackup,
        importPayload: importPayload,
        backupFilename: backupFilename,
        downloadBackup: downloadBackup,
        importFile: importFile,
        clearAllAppData: clearAllAppData,
        resetAllAppData: resetAllAppData,
        resetAndLoadDemo: resetAndLoadDemo,
        bindUi: bindUi
    };

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', bindUi);
    } else {
        bindUi();
    }
})();
