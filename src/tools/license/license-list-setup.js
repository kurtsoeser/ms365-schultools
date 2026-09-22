/**
 * Phase 1: Lizenz-Liste auf der Betreiber-SharePoint-Site anlegen (idempotent).
 */
(function () {
    'use strict';

    const G = window.ms365SpoGraph;
    if (!G) return;

    const SCOPES = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Sites.ReadWrite.All'
    ];

    function cfg() {
        return window.MS365_LICENSE_BACKEND || {};
    }

    function $(id) {
        return document.getElementById(id);
    }

    function log(msg) {
        const el = $('licSetupLog');
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

    async function findListByDisplayName(siteId, displayName, token) {
        const want = String(displayName || '').trim().toLowerCase();
        let url = G.graphPathSite(siteId) + '/lists?$select=id,displayName,webUrl&$top=200';
        while (url) {
            const page = await G.graphJson('GET', url, token, undefined, 'v1.0');
            const rows = (page && page.value) || [];
            for (let i = 0; i < rows.length; i++) {
                const n = String((rows[i] && rows[i].displayName) || '').trim().toLowerCase();
                if (n === want) return rows[i];
            }
            url = page && page['@odata.nextLink'] ? String(page['@odata.nextLink']) : '';
        }
        return null;
    }

    async function ensureColumns(siteId, listId, token, write) {
        const base = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns';
        const existing = await G.graphJson('GET', base + '?$select=name,displayName&$top=200', token, undefined, 'v1.0');
        const have = {};
        ((existing && existing.value) || []).forEach(function (c) {
            if (c && c.name) have[String(c.name).toLowerCase()] = true;
        });

        const choices = (cfg().statusChoices || ['trial', 'active', 'expired', 'blocked']).slice();
        const defs = [
            {
                name: 'TenantId',
                displayName: 'Tenant-ID',
                text: { allowMultipleLines: false, maxLength: 64 }
            },
            {
                name: 'PrimaryDomain',
                displayName: 'Primäre Domain',
                text: { allowMultipleLines: false, maxLength: 255 }
            },
            {
                name: 'AdditionalDomains',
                displayName: 'Weitere Domains / URLs',
                text: { allowMultipleLines: true }
            },
            {
                name: 'Status',
                displayName: 'Status',
                choice: {
                    allowTextEntry: false,
                    choices: choices
                }
            },
            {
                name: 'ValidUntil',
                displayName: 'Gültig bis',
                dateTime: { format: 'dateOnly' }
            },
            {
                name: 'ContactEmail',
                displayName: 'Kontakt-E-Mail',
                text: { allowMultipleLines: false, maxLength: 255 }
            },
            {
                name: 'Notes',
                displayName: 'Notizen',
                text: { allowMultipleLines: true }
            }
        ];

        for (let i = 0; i < defs.length; i++) {
            const d = defs[i];
            const key = String(d.name).toLowerCase();
            if (have[key]) {
                write('Spalte vorhanden: ' + d.name);
                continue;
            }
            write('Lege Spalte an: ' + d.name + ' …');
            await G.graphJson('POST', base, token, d, 'v1.0');
            have[key] = true;
            await G.sleep(150);
        }
    }

    async function runSetup() {
        const siteUrl = String(($('licSiteUrl') && $('licSiteUrl').value) || cfg().siteWebUrl || '').trim();
        const listTitle = String(($('licListName') && $('licListName').value) || cfg().listDisplayName || '').trim();
        if (!siteUrl) throw new Error('Site-URL fehlt.');
        if (!listTitle) throw new Error('Listenname fehlt.');

        log('Anmeldung / Token …');
        const token = await ensureToken();

        log('Löse Website auf: ' + siteUrl);
        const site = await G.resolveSiteFromWebUrl(token, siteUrl);
        const siteId = site && site.id ? String(site.id) : '';
        if (!siteId) throw new Error('Site-ID fehlt in der Graph-Antwort.');
        log('Site: ' + (site.displayName || siteId));

        let list = await findListByDisplayName(siteId, listTitle, token);
        if (list && list.id) {
            log('Liste existiert bereits: ' + listTitle + ' (' + list.id + ')');
        } else {
            log('Erstelle Liste „' + listTitle + '“ …');
            list = await G.graphJson(
                'POST',
                G.graphPathSite(siteId) + '/lists',
                token,
                {
                    displayName: listTitle,
                    description: 'Schul-Lizenzen / Tenant-Freischaltung für MS365-Schultools (Betreiber).',
                    list: { template: 'genericList' }
                },
                'v1.0'
            );
            if (!list || !list.id) throw new Error('Listen-ID fehlt in der Antwort.');
            log('Liste angelegt, ID: ' + list.id);
        }

        log('Prüfe / ergänze Spalten …');
        await ensureColumns(siteId, list.id, token, log);
        log('Fertig.');

        const webUrl = (list && list.webUrl) || siteUrl + '/Lists/' + encodeURIComponent(listTitle);
        const openEl = $('licOpenLink');
        if (openEl) {
            openEl.href = webUrl;
            openEl.hidden = false;
            openEl.style.display = '';
        }
        return { siteId: siteId, listId: String(list.id), webUrl: webUrl };
    }

    async function probeSite() {
        const siteUrl = String(($('licSiteUrl') && $('licSiteUrl').value) || cfg().siteWebUrl || '').trim();
        if (!siteUrl) throw new Error('Site-URL fehlt.');
        log('Anmeldung / Token …');
        const token = await ensureToken();
        log('Löse Website auf …');
        const site = await G.resolveSiteFromWebUrl(token, siteUrl);
        log('OK: ' + (site.displayName || '') + ' → ' + (site.id || ''));
        const listTitle = String(($('licListName') && $('licListName').value) || cfg().listDisplayName || '').trim();
        if (listTitle && site.id) {
            const existing = await findListByDisplayName(site.id, listTitle, token);
            if (existing) log('Liste „' + listTitle + '“ ist bereits vorhanden (' + existing.id + ').');
            else log('Liste „' + listTitle + '“ existiert noch nicht.');
        }
        return site;
    }

    function wire() {
        const siteInput = $('licSiteUrl');
        const listInput = $('licListName');
        const c = cfg();
        if (siteInput && !siteInput.value) siteInput.value = c.siteWebUrl || '';
        if (listInput && !listInput.value) listInput.value = c.listDisplayName || '';

        const btnProbe = $('licBtnProbe');
        const btnRun = $('licBtnRun');
        if (btnProbe) {
            btnProbe.addEventListener('click', function () {
                const logEl = $('licSetupLog');
                if (logEl) logEl.textContent = '';
                probeSite()
                    .then(function () {
                        toast('Website erreichbar.');
                    })
                    .catch(function (e) {
                        log('Fehler: ' + ((e && e.message) || e));
                        toast('Website-Prüfung fehlgeschlagen.');
                    });
            });
        }
        if (btnRun) {
            btnRun.addEventListener('click', function () {
                const logEl = $('licSetupLog');
                if (logEl) logEl.textContent = '';
                runSetup()
                    .then(function (result) {
                        toast('Lizenz-Liste ist bereit.');
                        try {
                            document.dispatchEvent(
                                new CustomEvent('ms365-license-list-setup-done', { detail: result || {} })
                            );
                        } catch {
                            /* ignore */
                        }
                    })
                    .catch(function (e) {
                        log('Fehler: ' + ((e && e.message) || e));
                        toast('Anlage fehlgeschlagen – siehe Protokoll.');
                    });
            });
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', wire);
    } else {
        wire();
    }

    /**
     * Prüft, welche Schema-Spalten fehlen (ohne etwas anzulegen).
     * @returns {Promise<{ siteId: string, listId: string|null, present: string[], missing: string[], listExists: boolean }>}
     */
    async function checkSchemaStatus() {
        const siteUrl = String(($('licSiteUrl') && $('licSiteUrl').value) || cfg().siteWebUrl || '').trim();
        const listTitle = String(($('licListName') && $('licListName').value) || cfg().listDisplayName || '').trim();
        if (!siteUrl) throw new Error('Site-URL fehlt.');
        if (!listTitle) throw new Error('Listenname fehlt.');

        const token = await ensureToken();
        const site = await G.resolveSiteFromWebUrl(token, siteUrl);
        const siteId = site && site.id ? String(site.id) : '';
        if (!siteId) throw new Error('Site-ID fehlt in der Graph-Antwort.');

        const required = [
            'TenantId',
            'PrimaryDomain',
            'AdditionalDomains',
            'Status',
            'ValidUntil',
            'ContactEmail',
            'Notes'
        ];
        const list = await findListByDisplayName(siteId, listTitle, token);
        if (!list || !list.id) {
            return {
                siteId: siteId,
                listId: null,
                listExists: false,
                present: [],
                missing: required.slice()
            };
        }

        const base = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(list.id) + '/columns';
        const existing = await G.graphJson(
            'GET',
            base + '?$select=name,displayName&$top=200',
            token,
            undefined,
            'v1.0'
        );
        const have = {};
        ((existing && existing.value) || []).forEach(function (c) {
            if (c && c.name) have[String(c.name).toLowerCase()] = true;
        });

        const present = [];
        const missing = [];
        required.forEach(function (name) {
            if (have[String(name).toLowerCase()]) present.push(name);
            else missing.push(name);
        });

        return {
            siteId: siteId,
            listId: String(list.id),
            listExists: true,
            present: present,
            missing: missing,
            webUrl: list.webUrl || ''
        };
    }

    window.ms365LicenseListSetup = {
        runSetup: runSetup,
        probeSite: probeSite,
        checkSchemaStatus: checkSchemaStatus
    };
})();
