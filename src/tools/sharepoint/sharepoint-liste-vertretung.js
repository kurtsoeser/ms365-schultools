/**
 * SharePoint Vertretungsplan-Liste (Schema + Anlegen).
 */
(function () {
    'use strict';

    const G = window.ms365SpoGraph;
    if (!G) return;

    const SCOPES = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Sites.ReadWrite.All'
    ];

    const COLUMNS = [
        { name: 'Datum', displayName: 'Datum', dateTime: { format: 'dateOnly' } },
        { name: 'Stunde', displayName: 'Stunde', text: { allowMultipleLines: false, maxLength: 40 } },
        { name: 'Klasse', displayName: 'Klasse', text: { allowMultipleLines: false, maxLength: 80 } },
        { name: 'Fach', displayName: 'Fach', text: { allowMultipleLines: false, maxLength: 80 } },
        { name: 'Abwesend', displayName: 'Abwesend', text: { allowMultipleLines: false, maxLength: 120 } },
        { name: 'Vertretung', displayName: 'Vertretung', text: { allowMultipleLines: false, maxLength: 120 } },
        { name: 'Raum', displayName: 'Raum', text: { allowMultipleLines: false, maxLength: 80 } },
        { name: 'Hinweis', displayName: 'Hinweis', text: { allowMultipleLines: true } }
    ];

    function $(id) {
        return document.getElementById(id);
    }

    function log(msg) {
        const el = $('vpLog');
        if (!el) return;
        el.textContent += (el.textContent ? '\n' : '') + msg;
        el.scrollTop = el.scrollHeight;
    }

    async function ensureToken() {
        return await G.getGraphToken(SCOPES);
    }

    async function addColumns(siteId, listId, token) {
        const base = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns';
        for (let i = 0; i < COLUMNS.length; i++) {
            await G.graphJson('POST', base, token, COLUMNS[i], 'v1.0');
            await G.sleep(120);
        }
    }

    async function createVertretungList(webUrl, listTitle, logFn) {
        const write = typeof logFn === 'function' ? logFn : log;
        const url = String(webUrl || '').trim();
        const title = String(listTitle || '').trim() || 'Vertretungsplan';
        if (!url) throw new Error('Bitte die Adresse der SharePoint-Website eintragen.');

        const token = await ensureToken();
        write('Website auflösen …');
        const site = await G.resolveSite(token, url);
        const siteId = site.id;
        write('Site: ' + (site.displayName || siteId));

        write('Liste anlegen: ' + title);
        const list = await G.graphJson(
            'POST',
            G.graphPathSite(siteId) + '/lists',
            token,
            {
                displayName: title,
                list: { template: 'genericList' }
            },
            'v1.0'
        );
        write('Spalten …');
        await addColumns(siteId, list.id, token);
        write('Fertig. Liste-ID: ' + list.id);
        return list;
    }

    function wire() {
        const run = $('vpBtnRun');
        if (!run) return;
        run.addEventListener('click', async function () {
            try {
                $('vpLog').textContent = '';
                await createVertretungList($('vpSiteUrl').value, $('vpListName').value, log);
                if (typeof window.ms365ToastOrAlert === 'function') {
                    window.ms365ToastOrAlert('Vertretungsplan-Liste angelegt.');
                }
            } catch (e) {
                log('Fehler: ' + (e && e.message ? e.message : String(e)));
            }
        });
        const probe = $('vpBtnProbe');
        if (probe) {
            probe.addEventListener('click', async function () {
                try {
                    $('vpLog').textContent = '';
                    const token = await ensureToken();
                    const site = await G.resolveSite(token, $('vpSiteUrl').value);
                    log('OK: ' + (site.webUrl || site.displayName || site.id));
                } catch (e) {
                    log('Fehler: ' + (e && e.message ? e.message : String(e)));
                }
            });
        }
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
                const setup = window.ms365AppDataV2.getSetup();
                const saved = setup && setup.intranetSiteUrl ? String(setup.intranetSiteUrl).trim() : '';
                if (saved && $('vpSiteUrl') && !$('vpSiteUrl').value) $('vpSiteUrl').value = saved;
            }
        } catch {
            /* ignore */
        }
    }

    window.ms365SpoVertretungListe = { createList: createVertretungList, COLUMNS: COLUMNS };

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', wire);
    else wire();
})();
