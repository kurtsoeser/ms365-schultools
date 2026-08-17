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
        const el = $('splLog');
        if (!el) return;
        el.textContent += (el.textContent ? '\n' : '') + msg;
        el.scrollTop = el.scrollHeight;
    }

    function toast(m) {
        if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
        else window.alert(m);
    }

    function loadTeachers() {
        if (typeof window.ms365TenantSettingsLoad !== 'function') {
            throw new Error('Stammdaten nicht geladen (tenant-settings-core.js fehlt?).');
        }
        const s = window.ms365TenantSettingsLoad();
        const teachers = (s && Array.isArray(s.teachers) ? s.teachers : []).filter(function (t) {
            return t && String(t.code || '').trim();
        });
        return teachers;
    }

    async function ensureToken() {
        return await G.getGraphToken(SCOPES);
    }

    async function addColumnsLehrer(siteId, listId, token) {
        const base = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns';
        const defs = [
            {
                name: 'LehrerCode',
                displayName: 'Kürzel',
                text: { allowMultipleLines: false, maxLength: 40 }
            },
            {
                name: 'EMail',
                displayName: 'E-Mail',
                text: { allowMultipleLines: false, maxLength: 255 }
            },
            {
                name: 'UPN',
                displayName: 'UPN',
                text: { allowMultipleLines: false, maxLength: 255 }
            }
        ];
        for (let i = 0; i < defs.length; i++) {
            await G.graphJson('POST', base, token, defs[i], 'v1.0');
            await G.sleep(120);
        }
    }

    async function createLehrerList(webUrl, listTitle, logFn) {
        const write = typeof logFn === 'function' ? logFn : log;
        const url = String(webUrl || '').trim();
        const title = String(listTitle || '').trim() || 'Lehrerinnen';
        if (!url) throw new Error('Bitte die Adresse der SharePoint-Website eintragen.');

        const teachers = loadTeachers();
        if (!teachers.length) {
            throw new Error('Keine Lehrkräfte in den Schul-Grundeinstellungen – zuerst unter Stammdaten pflegen.');
        }

        write('Lehrkräfte aus lokalem Speicher: ' + teachers.length);
        const token = await ensureToken();
        write('Löse Website auf …');
        const site = await G.resolveSiteFromWebUrl(token, url);
        const siteId = site && site.id ? String(site.id) : '';
        const siteTitle = site && site.displayName ? String(site.displayName) : '';
        if (!siteId) throw new Error('Site-ID fehlt in der Graph-Antwort.');
        write('Site: ' + (siteTitle || siteId));

        write('Erstelle Liste „' + title + '" …');
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

        write('Füge Spalten hinzu (Kürzel, E-Mail, UPN) …');
        await addColumnsLehrer(siteId, listId, token);
        write('Spalten fertig.');

        const itemsPath = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/items';
        let ok = 0;
        for (let i = 0; i < teachers.length; i++) {
            const t = teachers[i];
            const email = String(t.email || '').trim();
            const name = String(t.name || '').trim() || String(t.code || '').trim();
            const code = String(t.code || '').trim();
            await G.graphJson(
                'POST',
                itemsPath,
                token,
                {
                    fields: {
                        Title: name,
                        LehrerCode: code,
                        EMail: email,
                        UPN: email
                    }
                },
                'v1.0'
            );
            ok++;
            if (ok % 10 === 0) write('… ' + ok + ' Zeilen geschrieben');
            await G.sleep(80);
        }
        write('Fertig: ' + ok + ' Lehrkräfte als Listenelemente.');
        const listWeb = created && created.webUrl ? String(created.webUrl) : '';
        if (listWeb) write('Liste im Browser: ' + listWeb);
        if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
            window.ms365ActionLog.append({
                tool: 'sharepoint',
                action: 'create-lehrer-list',
                target: url,
                summary: 'Lehrerliste „' + title + '“ mit ' + ok + ' Einträgen'
            });
        }
        return { listId: listId, webUrl: listWeb, count: ok };
    }

    async function runCreate() {
        const logEl = $('splLog');
        if (logEl) logEl.textContent = '';
        const webUrl = String($('splSiteUrl') && $('splSiteUrl').value || '').trim();
        const listTitle = String($('splListName') && $('splListName').value || '').trim() || 'Lehrerinnen';
        const created = await createLehrerList(webUrl, listTitle);
        toast('Lehrerliste erstellt und befüllt.');
        return created;
    }

    window.ms365SpoLehrerListe = { createList: createLehrerList };

    const runBtn = $('splBtnRun');
    if (runBtn) {
        runBtn.addEventListener('click', function () {
            if (!window.confirm('Neue Liste auf der angegebenen Website anlegen und alle Lehrkräfte aus den Grundeinstellungen eintragen?')) return;
            runCreate().catch(function (e) {
                log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                toast('Fehler: ' + (e && e.message ? e.message : e));
            });
        });
    }

    const probeBtn = $('splBtnProbe');
    if (probeBtn) {
        probeBtn.addEventListener('click', function () {
            if ($('splLog')) $('splLog').textContent = '';
            const webUrl = String($('splSiteUrl') && $('splSiteUrl').value || '').trim();
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
        if (saved && $('splSiteUrl') && !$('splSiteUrl').value) $('splSiteUrl').value = saved;
    } catch {
        /* ignore */
    }
})();
