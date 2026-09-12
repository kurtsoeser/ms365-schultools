(function () {
    'use strict';

    const G = window.ms365SpoGraph;
    if (!G) return;

    const SCOPES = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Sites.ReadWrite.All'
    ];

    function $(id) {
        return document.getElementById(id);
    }

    function log(msg) {
        const el = $('sptLog');
        if (!el) return;
        el.textContent += (el.textContent ? '\n' : '') + msg;
        el.scrollTop = el.scrollHeight;
    }

    function toast(m) {
        if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
        else window.alert(m);
    }

    async function ensureToken() {
        return await G.getGraphToken(SCOPES);
    }

    /** Spalten für Power-Automate / Kalender-Sync (keine Listenelemente). */
    function columnDefsSchultermine() {
        const kategorien = [
            'Schulferien',
            'Feiertag',
            'Unterricht',
            'Prüfung',
            'Veranstaltung',
            'Elternabend',
            'Tag der offenen Tür',
            'sonstiges'
        ];
        return [
            {
                name: 'Beginn',
                displayName: 'Beginn',
                dateTime: { displayAs: 'default', format: 'dateTime' }
            },
            {
                name: 'Ende',
                displayName: 'Ende',
                dateTime: { displayAs: 'default', format: 'dateTime' }
            },
            {
                name: 'Kategorie',
                displayName: 'Kategorie',
                choice: { allowTextEntry: true, choices: kategorien }
            },
            {
                name: 'OutlookEventID',
                displayName: 'OutlookEventID',
                text: { allowMultipleLines: false, maxLength: 512 }
            },
            {
                name: 'Info',
                displayName: 'Info',
                text: { allowMultipleLines: true, maxLength: 8000 }
            },
            {
                name: 'ZeitraumText',
                displayName: 'ZeitraumText',
                text: { allowMultipleLines: false, maxLength: 255 }
            },
            {
                name: 'AllDay',
                displayName: 'AllDay',
                boolean: {}
            },
            {
                name: 'SyncStatus',
                displayName: 'SyncStatus',
                choice: {
                    allowTextEntry: false,
                    choices: ['ok', 'pending', 'error', 'manual']
                }
            },
            {
                name: 'SyncError',
                displayName: 'SyncError',
                text: { allowMultipleLines: true, maxLength: 4000 }
            },
            {
                name: 'LastSync',
                displayName: 'LastSync',
                dateTime: { displayAs: 'default', format: 'dateTime' }
            }
        ];
    }

    async function addColumns(siteId, listId, token) {
        const base = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns';
        const defs = columnDefsSchultermine();
        for (let i = 0; i < defs.length; i++) {
            await G.graphJson('POST', base, token, defs[i], 'v1.0');
            await G.sleep(120);
        }
    }

    async function createSchultermineList(webUrl, listTitle, logFn) {
        const write = typeof logFn === 'function' ? logFn : log;
        const url = String(webUrl || '').trim();
        const title = String(listTitle || '').trim() || 'Schultermine';
        if (!url) throw new Error('Bitte die Adresse der SharePoint-Website eintragen.');

        const token = await ensureToken();
        write('Löse Website auf …');
        const site = await G.resolveSiteFromWebUrl(token, url);
        const siteId = site && site.id ? String(site.id) : '';
        const siteTitle = site && site.displayName ? String(site.displayName) : '';
        if (!siteId) throw new Error('Site-ID fehlt in der Graph-Antwort.');
        write('Site: ' + (siteTitle || siteId));

        write('Erstelle leere Liste „' + title + '" …');
        const created = await G.graphJson(
            'POST',
            G.graphPathSite(siteId) + '/lists',
            token,
            {
                displayName: title,
                list: { template: 'genericList' }
            },
            'v1.0'
        );
        const listId = created && created.id ? String(created.id) : '';
        if (!listId) throw new Error('Listen-ID fehlt in der Antwort.');
        write('Liste angelegt, ID: ' + listId);

        write('Füge Spalten hinzu (Beginn, Ende, Kategorie, OutlookEventID, Info, ZeitraumText, AllDay, SyncStatus, SyncError, LastSync) …');
        await addColumns(siteId, listId, token);
        write('Fertig – keine Zeilen angelegt (Sync z. B. per Power Automate).');
        write('Tipp: Spalten SyncStatus / SyncError / LastSync für Flow-Überwachung nutzen.');
        const listWeb = created && created.webUrl ? String(created.webUrl) : '';
        if (listWeb) write('Liste: ' + listWeb);
        if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
            window.ms365ActionLog.append({
                tool: 'sharepoint',
                action: 'create-termine-list',
                target: url,
                summary: 'Schultermine-Liste „' + title + '“'
            });
        }
        return { listId: listId, webUrl: listWeb };
    }

    async function findListByTitle(token, siteId, listTitle) {
        const title = String(listTitle || '').trim() || 'Schultermine';
        const path =
            G.graphPathSite(siteId) +
            '/lists?$filter=' +
            encodeURIComponent("displayName eq '" + title.replace(/'/g, "''") + "'") +
            '&$select=id,displayName,webUrl';
        const data = await G.graphJson('GET', path, token, undefined, 'v1.0');
        const list = (data && data.value) || [];
        return list[0] || null;
    }

    async function probeSyncHealth(webUrl, listTitle, logFn) {
        const write = typeof logFn === 'function' ? logFn : log;
        const url = String(webUrl || '').trim();
        const title = String(listTitle || '').trim() || 'Schultermine';
        if (!url) throw new Error('Bitte die Adresse der SharePoint-Website eintragen.');

        const token = await ensureToken();
        write('Löse Website auf …');
        const site = await G.resolveSiteFromWebUrl(token, url);
        const siteId = site && site.id ? String(site.id) : '';
        if (!siteId) throw new Error('Site-ID fehlt.');
        write('Site: ' + (site.displayName || siteId));

        write('Suche Liste „' + title + '" …');
        const list = await findListByTitle(token, siteId, title);
        if (!list || !list.id) {
            write('Liste nicht gefunden. Bitte zuerst „Liste anlegen“ ausführen.');
            return { ok: false, reason: 'missing_list' };
        }
        write('Liste gefunden: ' + (list.displayName || title) + (list.webUrl ? ' · ' + list.webUrl : ''));

        const colsPath = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(list.id) + '/columns?$top=200';
        const colsData = await G.graphJson('GET', colsPath, token, undefined, 'v1.0');
        const colNames = ((colsData && colsData.value) || [])
            .map(function (c) {
                return String((c && (c.name || c.displayName)) || '').trim();
            })
            .filter(Boolean);
        const required = [
            'Beginn',
            'Ende',
            'Kategorie',
            'OutlookEventID',
            'Info',
            'ZeitraumText',
            'AllDay',
            'SyncStatus',
            'SyncError',
            'LastSync'
        ];
        const missing = required.filter(function (n) {
            return colNames.indexOf(n) === -1;
        });
        if (missing.length) {
            write('Fehlende Spalten: ' + missing.join(', '));
            write('Hinweis: Bei älteren Listen Spalten manuell ergänzen oder Liste neu anlegen.');
        } else {
            write('Spalten komplett (inkl. SyncStatus / SyncError / LastSync).');
        }

        const itemsPath =
            G.graphPathSite(siteId) +
            '/lists/' +
            encodeURIComponent(list.id) +
            '/items?$expand=fields&$top=50';
        const itemsData = await G.graphJson('GET', itemsPath, token, undefined, 'v1.0');
        const items = (itemsData && itemsData.value) || [];
        let withEventId = 0;
        let withError = 0;
        let pending = 0;
        let okStatus = 0;
        items.forEach(function (it) {
            const f = (it && it.fields) || {};
            if (f.OutlookEventID) withEventId++;
            if (f.SyncError) withError++;
            const st = String(f.SyncStatus || '').toLowerCase();
            if (st === 'pending') pending++;
            if (st === 'ok') okStatus++;
        });
        write(
            'Stichprobe: ' +
                items.length +
                ' Zeilen · mit OutlookEventID: ' +
                withEventId +
                ' · SyncStatus ok: ' +
                okStatus +
                ' · pending: ' +
                pending +
                ' · mit SyncError: ' +
                withError
        );
        if (!items.length) {
            write('Noch keine Termine in der Liste – Power Automate / Import kann befüllen.');
        } else if (withError) {
            write('Achtung: Es gibt Zeilen mit SyncError – Flow oder Kalenderrechte prüfen.');
        } else if (items.length && withEventId === items.length) {
            write('Alle Stichproben-Zeilen haben eine OutlookEventID – Sync wirkt gesund.');
        } else if (items.length && withEventId < items.length) {
            write(
                'Einige Zeilen ohne OutlookEventID (' +
                    (items.length - withEventId) +
                    ') – ggf. manuell angelegt oder Flow noch nicht gelaufen.'
            );
        }

        const summary = {
            ok: missing.length === 0 && withError === 0,
            listId: list.id,
            webUrl: list.webUrl || '',
            missingColumns: missing,
            itemSample: items.length,
            withEventId: withEventId,
            withError: withError,
            pending: pending,
            okStatus: okStatus
        };
        const panel = $('sptSyncHealth');
        if (panel) {
            panel.hidden = false;
            panel.textContent =
                (summary.ok ? 'Status: OK' : 'Status: prüfen') +
                ' · Liste „' +
                title +
                '" · Stichprobe ' +
                items.length +
                ' · EventIDs ' +
                withEventId +
                ' · Fehler ' +
                withError +
                (missing.length ? ' · fehlende Spalten: ' + missing.join(', ') : '');
            panel.classList.toggle('ok', !!summary.ok);
            panel.classList.toggle('warn', !summary.ok);
        }
        return summary;
    }

    async function runCreate() {
        const logEl = $('sptLog');
        if (logEl) logEl.textContent = '';
        const webUrl = String($('sptSiteUrl') && $('sptSiteUrl').value || '').trim();
        const listTitle = String($('sptListName') && $('sptListName').value || '').trim() || 'Schultermine';
        const created = await createSchultermineList(webUrl, listTitle);
        toast('Schultermine-Liste mit Spalten erstellt.');
        return created;
    }

    /**
     * Bereits normalisierte Felder (Title, Beginn, …) als List Items anlegen.
     * @param {string} webUrl
     * @param {string} listTitle
     * @param {Record<string, unknown>[]} fieldsList
     * @param {(msg: string) => void} [logFn]
     */
    async function importTermine(webUrl, listTitle, fieldsList, logFn) {
        const write = typeof logFn === 'function' ? logFn : log;
        const url = String(webUrl || '').trim();
        const title = String(listTitle || '').trim() || 'Schultermine';
        const rows = Array.isArray(fieldsList) ? fieldsList : [];
        if (!url) throw new Error('Bitte die Adresse der SharePoint-Website eintragen.');
        if (!rows.length) throw new Error('Keine Termine zum Import.');

        const token = await ensureToken();
        write('Löse Website auf …');
        const site = await G.resolveSiteFromWebUrl(token, url);
        const siteId = site && site.id ? String(site.id) : '';
        if (!siteId) throw new Error('Site-ID fehlt.');

        write('Suche Liste „' + title + '" …');
        const list = await findListByTitle(token, siteId, title);
        if (!list || !list.id) {
            throw new Error('Liste „' + title + '“ nicht gefunden. Bitte zuerst anlegen.');
        }

        const itemsPath = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(list.id) + '/items';
        let ok = 0;
        for (let i = 0; i < rows.length; i++) {
            const fields = rows[i] || {};
            await G.graphJson('POST', itemsPath, token, { fields: fields }, 'v1.0');
            ok++;
            if (ok % 10 === 0) write('… ' + ok + ' Termine geschrieben');
            await G.sleep(80);
        }
        write('Import fertig: ' + ok + ' Zeilen.');
        if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
            window.ms365ActionLog.append({
                tool: 'sharepoint',
                action: 'import-termine',
                target: url,
                summary: ok + ' Termine → „' + title + '“'
            });
        }
        return { ok: ok, listId: list.id, webUrl: list.webUrl || '' };
    }

    window.ms365SpoSchultermine = {
        createList: createSchultermineList,
        probeSyncHealth: probeSyncHealth,
        importTermine: importTermine
    };

    const runBtn = $('sptBtnRun');
    if (runBtn) {
        runBtn.addEventListener('click', function () {
            if (!window.confirm('Neue leere Liste auf der Website anlegen (nur Struktur, keine Termine)?')) return;
            runCreate().catch(function (e) {
                log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                toast('Fehler: ' + (e && e.message ? e.message : e));
            });
        });
    }

    const probeBtn = $('sptBtnProbe');
    if (probeBtn) {
        probeBtn.addEventListener('click', function () {
            if ($('sptLog')) $('sptLog').textContent = '';
            const webUrl = String($('sptSiteUrl') && $('sptSiteUrl').value || '').trim();
            if (!webUrl) {
                toast('Website-URL fehlt.');
                return;
            }
            ensureToken()
                .then(function (token) {
                    return G.resolveSiteFromWebUrl(token, webUrl);
                })
                .then(function (site) {
                    log('Site gefunden: ' + (site.displayName || '') + '\nid: ' + (site.id || ''));
                    if (site.webUrl) log('webUrl: ' + site.webUrl);
                    toast('Website erkannt.');
                })
                .catch(function (e) {
                    log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                    toast('Fehler: ' + (e && e.message ? e.message : e));
                });
        });
    }

    const healthBtn = $('sptBtnSyncHealth');
    if (healthBtn) {
        healthBtn.addEventListener('click', function () {
            if ($('sptLog')) $('sptLog').textContent = '';
            const webUrl = String($('sptSiteUrl') && $('sptSiteUrl').value || '').trim();
            const listTitle = String($('sptListName') && $('sptListName').value || '').trim() || 'Schultermine';
            probeSyncHealth(webUrl, listTitle)
                .then(function (summary) {
                    toast(summary && summary.ok ? 'Sync-Check OK' : 'Sync-Check: bitte Protokoll prüfen');
                })
                .catch(function (e) {
                    log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                    toast('Fehler: ' + (e && e.message ? e.message : e));
                });
        });
    }

    try {
        const setup = window.ms365AppDataV2 && window.ms365AppDataV2.getSetup ? window.ms365AppDataV2.getSetup() : null;
        const saved = setup && setup.intranetSiteUrl ? String(setup.intranetSiteUrl).trim() : '';
        if (saved && $('sptSiteUrl') && !$('sptSiteUrl').value) $('sptSiteUrl').value = saved;
    } catch {
        /* ignore */
    }
})();
