/**
 * Vollständiges lokales Browser-Backup (localStorage der App).
 * Zum Übertragen zwischen Browsern/PCs. Microsoft-Anmeldung (MSAL) bleibt
 * bewusst außen vor – im Zielbrowser erneut anmelden.
 */
(function () {
    'use strict';

    const KIND = 'ms365-browser-backup-v1';
    const VERSION = 1;

    function getStore(storage) {
        if (storage) return storage;
        try {
            if (typeof localStorage !== 'undefined') return localStorage;
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

    function backupFilename(d) {
        return 'ms365-browser-backup-' + localDateStamp(d) + '.json';
    }

    function buildBackup(storage, now) {
        const local = collectLocalStorage(storage);
        const when = asDate(now);
        return {
            kind: KIND,
            version: VERSION,
            exportedAt: when.toISOString(),
            keyCount: Object.keys(local).length,
            localStorage: local
        };
    }

    function isBackupPayload(obj) {
        return !!(
            obj &&
            typeof obj === 'object' &&
            obj.kind === KIND &&
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
        const replace = !opts || opts.replace !== false;
        const incoming = payload.localStorage;
        const written = [];
        const removed = [];
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

        if (errors.length) {
            const err = new Error('Backup nur teilweise geschrieben: ' + errors.slice(0, 3).join('; '));
            err.details = { written: written, removed: removed, skipped: skipped, errors: errors };
            throw err;
        }
        return { written: written, removed: removed, skipped: skipped, errors: errors };
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
        downloadJson(backupFilename(now), payload);
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
            const when = obj.exportedAt ? String(obj.exportedAt).replace('T', ' ').replace(/\.\d+Z$/, ' UTC') : '';
            return (
                'Dieses Browser-Backup enthält ' +
                n +
                ' gespeicherte Einträge' +
                (when ? ' (Stand: ' + when + ')' : '') +
                '. Alle lokalen Schuldaten und Werkzeug-Zwischenstände in diesem Browser werden ersetzt. Die Microsoft-Anmeldung bleibt unberührt. Fortfahren?'
            );
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

    function onExportClick() {
        try {
            const payload = downloadBackup();
            setStatus('Backup gespeichert: ' + payload.keyCount + ' Einträge.', 'ok');
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
        buildBackup: buildBackup,
        applyBackup: applyBackup,
        importPayload: importPayload,
        backupFilename: backupFilename,
        downloadBackup: downloadBackup,
        importFile: importFile,
        bindUi: bindUi
    };

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', bindUi);
    } else {
        bindUi();
    }
})();
