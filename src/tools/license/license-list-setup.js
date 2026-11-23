/**
 * Phase 1: Lizenz-Liste auf der Betreiber-SharePoint-Site anlegen (idempotent).
 */
(function () {
    'use strict';

    const G = window.ms365SpoGraph;
    if (!G) return;

    // Spalten anlegen: Sites.Manage.All (Sites.ReadWrite.All reicht laut Graph-Doku nicht).
    const SCOPES = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Sites.Manage.All'
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

    function withTimeout(promise, ms, label) {
        const wait = Number(ms) > 0 ? Number(ms) : 45000;
        return Promise.race([
            promise,
            new Promise(function (_, reject) {
                setTimeout(function () {
                    reject(
                        new Error(
                            (label || 'Vorgang') +
                                ' abgebrochen (Timeout ' +
                                Math.round(wait / 1000) +
                                's). Bitte Popup erlauben oder Spalten manuell anlegen.'
                        )
                    );
                }, wait);
            })
        ]);
    }

    async function ensureToken() {
        // Explizit Popup – Silent/iframe hängt oft bei Sites.Manage.All.
        if (typeof window.ms365AuthAcquireTokenPopup === 'function') {
            return withTimeout(
                window.ms365AuthAcquireTokenPopup(SCOPES),
                45000,
                'Anmeldung/Token'
            );
        }
        return withTimeout(G.getGraphToken(SCOPES), 45000, 'Anmeldung/Token');
    }

    /** Stammdaten-Spalten inkl. Rollen-E-Mails – per License-API ohne Sites.Manage.All. */
    var STAMMDATEN_COLUMNS = [
        { name: 'Standort', displayName: 'Standort', type: 'text' },
        { name: 'Strasse', displayName: 'Straße', type: 'text' },
        { name: 'PLZ', displayName: 'PLZ', type: 'text' },
        { name: 'Ort', displayName: 'Ort', type: 'text' },
        { name: 'Direktion', displayName: 'Direktion', type: 'text' },
        { name: 'DirektionEmail', displayName: 'Direktion E-Mail', type: 'text' },
        { name: 'Schuladministration', displayName: 'Schuladministration', type: 'text' },
        { name: 'SchuladministrationEmail', displayName: 'Schuladministration E-Mail', type: 'text' },
        { name: 'ItM365Admin', displayName: 'IT-/M365-Administration', type: 'text' },
        { name: 'ItM365AdminEmail', displayName: 'IT-/M365-Administration E-Mail', type: 'text' },
        { name: 'WeitereKontakte', displayName: 'Weitere Kontakte / Rollen', type: 'multiline' },
        { name: 'Telefon', displayName: 'Telefon', type: 'text' },
        { name: 'Schulnummer', displayName: 'Schulnummer', type: 'text' }
    ];

    /**
     * Bevorzugter Weg: License-API (App-only Graph auf dem Server) + adminCreateColumn.
     * @returns {Promise<object|null>}
     */
    async function runSetupViaLicenseApi() {
        var api = window.ms365LicenseApi;
        if (!api || typeof api.acquireLicenseToken !== 'function') return null;
        if (typeof api.adminEnsureSchema !== 'function' && typeof api.adminCreateColumn !== 'function') {
            return null;
        }

        log('License-API: Token holen (ohne Sites.Manage.All) …');
        var token = await withTimeout(
            api.acquireLicenseToken({ popup: true }),
            45000,
            'License-API Token'
        );

        if (typeof api.adminEnsureSchema === 'function') {
            log('License-API: Schema sicherstellen …');
            try {
                var schema = await api.adminEnsureSchema(token);
                var created =
                    schema && schema.stammdaten && Array.isArray(schema.stammdaten.created)
                        ? schema.stammdaten.created
                        : [];
                if (created.length) log('API neu: ' + created.join(', '));
                else log('API-Schema geprüft.');
            } catch (e) {
                log('API ensure-schema: ' + ((e && e.message) || e));
            }
        }

        var have = {};
        if (typeof api.adminListColumns === 'function') {
            try {
                var colRes = await api.adminListColumns(token);
                var cols =
                    colRes && Array.isArray(colRes.columns)
                        ? colRes.columns
                        : Array.isArray(colRes)
                          ? colRes
                          : [];
                cols.forEach(function (c) {
                    var n = String((c && c.name) || '')
                        .trim()
                        .toLowerCase();
                    if (n) have[n] = true;
                });
            } catch (e) {
                log('Spaltenliste: ' + ((e && e.message) || e));
            }
        }

        var createdLocal = [];
        if (typeof api.adminCreateColumn === 'function') {
            for (var i = 0; i < STAMMDATEN_COLUMNS.length; i++) {
                var f = STAMMDATEN_COLUMNS[i];
                var key = String(f.name).toLowerCase();
                if (have[key]) {
                    log('Spalte vorhanden: ' + f.name);
                    continue;
                }
                log('Lege Spalte an (API): ' + f.name + ' …');
                try {
                    await api.adminCreateColumn(token, {
                        name: f.name,
                        displayName: f.displayName,
                        type: f.type || 'text'
                    });
                    have[key] = true;
                    createdLocal.push(f.name);
                } catch (ce) {
                    var msg = String((ce && ce.message) || ce || '');
                    if ((ce && ce.status === 409) || /existiert bereits|already exists/i.test(msg)) {
                        have[key] = true;
                        log('Spalte vorhanden: ' + f.name);
                    } else {
                        log('Fehler ' + f.name + ': ' + msg);
                    }
                }
            }
        }

        log(
            createdLocal.length
                ? 'Fertig über License-API – neu: ' + createdLocal.join(', ')
                : 'Fertig über License-API – Spalten aktuell.'
        );
        return {
            via: 'license-api',
            created: createdLocal,
            webUrl: String((cfg().siteWebUrl || '')).trim()
        };
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
            },
            {
                name: 'ItContacts',
                displayName: 'IT-Kontakte',
                text: { allowMultipleLines: true }
            },
            {
                name: 'LicenseEntitlements',
                displayName: 'Lizenz-Module (JSON)',
                text: { allowMultipleLines: true }
            },
            { name: 'Standort', displayName: 'Standort', text: { allowMultipleLines: false, maxLength: 255 } },
            { name: 'Strasse', displayName: 'Straße', text: { allowMultipleLines: false, maxLength: 255 } },
            { name: 'PLZ', displayName: 'PLZ', text: { allowMultipleLines: false, maxLength: 32 } },
            { name: 'Ort', displayName: 'Ort', text: { allowMultipleLines: false, maxLength: 255 } },
            { name: 'Direktion', displayName: 'Direktion', text: { allowMultipleLines: false, maxLength: 255 } },
            {
                name: 'DirektionEmail',
                displayName: 'Direktion E-Mail',
                text: { allowMultipleLines: false, maxLength: 255 }
            },
            {
                name: 'Schuladministration',
                displayName: 'Schuladministration',
                text: { allowMultipleLines: false, maxLength: 255 }
            },
            {
                name: 'SchuladministrationEmail',
                displayName: 'Schuladministration E-Mail',
                text: { allowMultipleLines: false, maxLength: 255 }
            },
            {
                name: 'ItM365Admin',
                displayName: 'IT-/M365-Administration',
                text: { allowMultipleLines: false, maxLength: 255 }
            },
            {
                name: 'ItM365AdminEmail',
                displayName: 'IT-/M365-Administration E-Mail',
                text: { allowMultipleLines: false, maxLength: 255 }
            },
            { name: 'WeitereKontakte', displayName: 'Weitere Kontakte / Rollen', text: { allowMultipleLines: true } },
            { name: 'Telefon', displayName: 'Telefon', text: { allowMultipleLines: false, maxLength: 64 } },
            { name: 'Schulnummer', displayName: 'Schulnummer', text: { allowMultipleLines: false, maxLength: 255 } },
            { name: 'DomainHinweis', displayName: 'Domain-Hinweis', text: { allowMultipleLines: true } },
            { name: 'Datenstatus', displayName: 'Datenstatus', text: { allowMultipleLines: false, maxLength: 255 } },
            { name: 'WorkshopTeilnehmende', displayName: 'Workshop-Teilnehmende', text: { allowMultipleLines: true } },
            { name: 'WorkshopQuelle', displayName: 'Workshop-Quelle', text: { allowMultipleLines: true } },
            { name: 'OeffentlicheQuellen', displayName: 'Öffentliche Quellen', text: { allowMultipleLines: true } },
            { name: 'OffenePunkte', displayName: 'Offene Punkte', text: { allowMultipleLines: true } }
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

        // 1) Bevorzugt License-API – hängt nicht an Sites.Manage.All / iframe
        try {
            var viaApi = await runSetupViaLicenseApi();
            if (viaApi) {
                const openElApi = $('licOpenLink');
                if (openElApi && viaApi.webUrl) {
                    openElApi.href = viaApi.webUrl;
                    openElApi.hidden = false;
                    const openLabelApi = $('licOpenLinkLabel');
                    if (openLabelApi) openLabelApi.textContent = 'Site öffnen';
                }
                return viaApi;
            }
        } catch (apiErr) {
            log('License-API-Weg: ' + ((apiErr && apiErr.message) || apiErr));
            log('Fallback: Browser-Graph mit Sites.Manage.All …');
        }

        log('Anmeldung / Token (Sites.Manage.All, Popup) …');
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
            const openLabel = $('licOpenLinkLabel');
            if (openLabel) openLabel.textContent = 'Liste öffnen';
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
        const openEl = $('licOpenLink');
        if (openEl) {
            openEl.href = (site && site.webUrl) || siteUrl;
            openEl.hidden = false;
            const openLabel = $('licOpenLinkLabel');
            if (openLabel) openLabel.textContent = 'Site öffnen';
        }
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
            'Notes',
            'Direktion',
            'DirektionEmail',
            'Schuladministration',
            'SchuladministrationEmail',
            'ItM365Admin',
            'ItM365AdminEmail'
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
