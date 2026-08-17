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

        write('Füge Spalten hinzu (Beginn, Ende, Kategorie, OutlookEventID, Info, ZeitraumText, AllDay) …');
        await addColumns(siteId, listId, token);
        write('Fertig – keine Zeilen angelegt (Sync z. B. per Power Automate).');
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

    async function runCreate() {
        const logEl = $('sptLog');
        if (logEl) logEl.textContent = '';
        const webUrl = String($('sptSiteUrl') && $('sptSiteUrl').value || '').trim();
        const listTitle = String($('sptListName') && $('sptListName').value || '').trim() || 'Schultermine';
        const created = await createSchultermineList(webUrl, listTitle);
        toast('Schultermine-Liste mit Spalten erstellt.');
        return created;
    }

    window.ms365SpoSchultermine = { createList: createSchultermineList };

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

    try {
        const setup = window.ms365AppDataV2 && window.ms365AppDataV2.getSetup ? window.ms365AppDataV2.getSetup() : null;
        const saved = setup && setup.intranetSiteUrl ? String(setup.intranetSiteUrl).trim() : '';
        if (saved && $('sptSiteUrl') && !$('sptSiteUrl').value) $('sptSiteUrl').value = saved;
    } catch {
        /* ignore */
    }
})();
