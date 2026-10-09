(function () {
    'use strict';

    const G = window.ms365SpoGraph;
    if (!G) return;

    const SCOPES = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Sites.ReadWrite.All'
    ];

    const STORAGE_KEY = 'ms365-lfr-setup-v1';
    const STEP_STORAGE_KEY = 'ms365-lfr-setup-step-v1';
    const SETUP_STEP_COUNT = 5;
    const FLOW_DONE_KEY = 'ms365-pa-done-lehrer-freistellung';
    const PERMS_STORAGE_KEY = 'ms365-lfr-perms-v1';
    const TEMPLATE_BASE = '../assets/power-automate/lehrer-freistellung';
    const FLOW_ASSET_ID = 'b8d4e2f1-6a3c-4d5e-8f9a-1b2c3d4e5f6a';

    /** Platzhalter in der Flow-Vorlage – werden beim Paketbau ersetzt. */
    const SOURCE = {
        siteUrl: 'https://kurtrocks.sharepoint.com/sites/MS365-Schultools',
        listId: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee',
        emailDirektion: 'direktor@ms365.schule',
        emailDirektor: 'Direktor@ms365.schule',
        emailMailbox: 'automate@ms365.schule',
        connectionOwner: 'kurt@kurtsoeser.at'
    };

    function $(id) {
        return document.getElementById(id);
    }

    function log(msg) {
        const el = $('frLog');
        if (!el) return;
        el.textContent += (el.textContent ? '\n' : '') + msg;
        el.scrollTop = el.scrollHeight;
        const details = el.closest('details');
        if (details) details.open = true;
    }

    function toast(m) {
        if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
        else window.alert(m);
    }

    function loadCfg() {
        try {
            return JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}') || {};
        } catch (e) {
            return {};
        }
    }

    function saveCfg(cfg) {
        try {
            localStorage.setItem(STORAGE_KEY, JSON.stringify(cfg));
        } catch (e) {}
        if (window.ms365BrowserBackup && typeof window.ms365BrowserBackup.notifyLocalDataChanged === 'function') {
            window.ms365BrowserBackup.notifyLocalDataChanged('lfr-setup');
        }
    }

    function readForm() {
        return {
            siteUrl: String(($('frSiteUrl') && $('frSiteUrl').value) || '').trim().replace(/\/$/, ''),
            listName: String(($('frListName') && $('frListName').value) || '').trim() || 'Lehrer-Freistellungen',
            outlookCalendarUser: String(($('lfrOutlookCalUser') && $('lfrOutlookCalUser').value) || '')
                .trim()
                .toLowerCase(),
            outlookCalendarId: String(($('lfrOutlookCalId') && $('lfrOutlookCalId').value) || '').trim(),
            listId: String(($('frListId') && $('frListId').value) || '').trim(),
            emailDirektion: String(($('frEmailDirektion') && $('frEmailDirektion').value) || '')
                .trim()
                .toLowerCase(),
            emailMailbox: String(($('frEmailMailbox') && $('frEmailMailbox').value) || '')
                .trim()
                .toLowerCase(),
            flowServiceAccount: String(($('frFlowServiceAccount') && $('frFlowServiceAccount').value) || '')
                .trim()
                .toLowerCase(),
            flowDisplayName:
                String(($('frFlowName') && $('frFlowName').value) || '').trim() ||
                'Lehrer-Freistellungen - Genehmigung Direktion',
            mailAsTechnikUser: !($('frMailAsTechnikUser') && !$('frMailAsTechnikUser').checked)
        };
    }

    function writeForm(cfg) {
        if (!cfg) return;
        if ($('frSiteUrl') && cfg.siteUrl) $('frSiteUrl').value = cfg.siteUrl;
        if ($('frListName') && cfg.listName) $('frListName').value = cfg.listName;
        if ($('frListId') && cfg.listId) $('frListId').value = cfg.listId;
        if ($('frEmailDirektion') && cfg.emailDirektion) $('frEmailDirektion').value = cfg.emailDirektion;
        if ($('frEmailMailbox') && cfg.emailMailbox) $('frEmailMailbox').value = cfg.emailMailbox;
        if ($('frFlowServiceAccount') && cfg.flowServiceAccount) {
            $('frFlowServiceAccount').value = cfg.flowServiceAccount;
        }
        if ($('frFlowName') && cfg.flowDisplayName) $('frFlowName').value = cfg.flowDisplayName;
        if ($('frMailAsTechnikUser')) {
            $('frMailAsTechnikUser').checked = cfg.mailAsTechnikUser !== false;
        }
        if ($('lfrOutlookCalUser') && cfg.outlookCalendarUser) {
            $('lfrOutlookCalUser').value = cfg.outlookCalendarUser;
        }
        if ($('lfrOutlookCalId') && cfg.outlookCalendarId) $('lfrOutlookCalId').value = cfg.outlookCalendarId;
    }

    function effectiveFlowAccount(cfg) {
        const st = window.ms365LfrSetupStatus;
        if (st && typeof st.effectiveLfrFlowAccount === 'function') {
            return st.effectiveLfrFlowAccount(cfg);
        }
        const explicit = String((cfg && cfg.flowServiceAccount) || '').trim().toLowerCase();
        if (explicit) return explicit;
        return String((cfg && cfg.emailMailbox) || '').trim().toLowerCase();
    }

    function refreshImportAccountHint() {
        const el = $('frImportAccountHint');
        if (!el) return;
        const account = effectiveFlowAccount(readForm());
        el.textContent = account || 'Technik-Konto aus Schritt 2';
    }

    function persistFromForm() {
        const cfg = readForm();
        saveCfg(cfg);
        refreshGlance();
        return cfg;
    }

    async function ensureToken() {
        return await G.getGraphToken(SCOPES);
    }

    function columnDefsLfr() {
        const sch = window.ms365LfrSchema;
        if (sch && sch.columns && typeof sch.toGraphColumnBody === 'function') {
            return sch.columns.map(function (def) {
                return sch.toGraphColumnBody(def);
            });
        }
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
                name: 'Status',
                displayName: 'Status',
                choice: {
                    allowTextEntry: false,
                    choices: ['Ausstehend', 'Genehmigt', 'Abgelehnt']
                }
            },
            {
                name: 'Klasse',
                displayName: 'Klasse',
                choice: {
                    allowTextEntry: true,
                    choices: ['1AHW', '2AHW', '3AHW', '4AHW', '5AHW']
                }
            },
            {
                name: 'Klassenvorstand',
                displayName: 'Klassenvorstand',
                personOrGroup: {
                    allowMultipleSelection: false,
                    chooseFromType: 'peopleOnly'
                }
            },
            {
                name: 'Kategorie',
                displayName: 'Kategorie',
                choice: {
                    allowTextEntry: true,
                    choices: [
                        'Ärztlicher Termin',
                        'Familiäre Angelegenheit',
                        'Bewerbung / Schnuppertag',
                        'Sonstiges'
                    ]
                }
            },
            {
                name: 'Beschreibung',
                displayName: 'Beschreibung',
                text: { allowMultipleLines: true, maxLength: 8000 }
            },
            {
                name: 'Bemerkungen',
                displayName: 'Bemerkungen',
                text: { allowMultipleLines: true, maxLength: 8000 }
            }
        ];
    }

    async function ensureColumns(siteId, listId, token, write) {
        const logFn = typeof write === 'function' ? write : log;
        const base = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns';
        const existing = await G.graphJson(
            'GET',
            base + '?$select=id,name,displayName,dateTime&$top=200',
            token,
            undefined,
            'v1.0'
        );
        const have = {};
        const byName = {};
        ((existing && existing.value) || []).forEach(function (c) {
            if (c && c.name) {
                const key = String(c.name).toLowerCase();
                have[key] = true;
                byName[key] = c;
            }
        });

        const defs = columnDefsLfr();
        const added = [];
        const skipped = [];
        const upgraded = [];
        for (let i = 0; i < defs.length; i++) {
            const d = defs[i];
            const key = String(d.name).toLowerCase();
            if (have[key]) {
                skipped.push(d.name);
                const wantDt = d.dateTime && d.dateTime.format === 'dateTime';
                const col = byName[key];
                const curFmt =
                    col && col.dateTime && col.dateTime.format
                        ? String(col.dateTime.format)
                        : '';
                if (wantDt && col && col.id && curFmt === 'dateOnly') {
                    try {
                        logFn('Aktualisiere Spalte „' + d.name + '“ auf Datum+Uhrzeit …');
                        await G.graphJson(
                            'PATCH',
                            base + '/' + encodeURIComponent(col.id),
                            token,
                            { dateTime: { displayAs: 'default', format: 'dateTime' } },
                            'v1.0'
                        );
                        upgraded.push(d.name);
                        await G.sleep(140);
                    } catch (e) {
                        logFn(
                            '  ! Spalte „' +
                                d.name +
                                '“ konnte nicht auf Uhrzeit umgestellt werden: ' +
                                (e && e.message ? e.message : e) +
                                ' – in SharePoint manuell: Spalte → Datum und Uhrzeit.'
                        );
                    }
                }
                continue;
            }
            logFn('Lege Spalte an: ' + d.name + ' …');
            await G.graphJson('POST', base, token, d, 'v1.0');
            have[key] = true;
            added.push(d.name);
            await G.sleep(140);
        }
        if (added.length) {
            logFn('Neu angelegt: ' + added.join(', '));
        }
        if (upgraded.length) {
            logFn('Auf Datum+Uhrzeit umgestellt: ' + upgraded.join(', '));
        }
        if (skipped.length === defs.length && !upgraded.length) {
            logFn('Alle Spalten vorhanden (' + skipped.join(', ') + ').');
        } else if (skipped.length) {
            logFn('Bereits vorhanden: ' + skipped.join(', '));
        }
        return { added: added, skipped: skipped, upgraded: upgraded };
    }

    async function findListByTitle(token, siteId, listTitle) {
        const title = String(listTitle || '').trim();
        const path =
            G.graphPathSite(siteId) +
            '/lists?$filter=' +
            encodeURIComponent("displayName eq '" + title.replace(/'/g, "''") + "'") +
            '&$select=id,displayName,webUrl';
        const data = await G.graphJson('GET', path, token, undefined, 'v1.0');
        const list = (data && data.value) || [];
        return list[0] || null;
    }

    async function createOrResolveList(cfg, logFn) {
        const write = typeof logFn === 'function' ? logFn : log;
        if (!cfg.siteUrl) throw new Error('Bitte die SharePoint-Website eintragen.');
        if (!cfg.emailDirektion || !cfg.emailMailbox) {
            throw new Error('Bitte Direktion und Absender-Postfach ausfüllen (Schritt 2).');
        }
        if (!effectiveFlowAccount(cfg)) {
            throw new Error('Bitte Technik-Konto für Power Automate eintragen (Schritt 2).');
        }

        const token = await ensureToken();
        write('Löse Website auf …');
        const site = await G.resolveSiteFromWebUrl(token, cfg.siteUrl);
        const siteId = site && site.id ? String(site.id) : '';
        if (!siteId) throw new Error('Site-ID fehlt in der Graph-Antwort.');
        write('Site: ' + (site.displayName || siteId));

        let listId = cfg.listId;
        let listWeb = '';
        let listDisplayName = String(cfg.listName || '').trim();

        if (listId) {
            write('Verwende vorhandene Listen-ID: ' + listId);
            try {
                const existing = await G.graphJson(
                    'GET',
                    G.graphPathSite(siteId) +
                        '/lists/' +
                        encodeURIComponent(listId) +
                        '?$select=id,displayName,webUrl',
                    token,
                    undefined,
                    'v1.0'
                );
                listWeb = existing && existing.webUrl ? String(existing.webUrl) : '';
                listDisplayName = String((existing && existing.displayName) || listDisplayName || listId).trim();
                write('Liste gefunden: ' + listDisplayName);
                if (
                    cfg.listName &&
                    listDisplayName &&
                    String(cfg.listName).trim() !== listDisplayName
                ) {
                    write(
                        'Hinweis: SharePoint-Titel ist „' +
                            listDisplayName +
                            '“, nicht „' +
                            cfg.listName +
                            '“ im Formular – Berechtigungen nutzen den echten Titel.'
                    );
                }
            } catch (e) {
                throw new Error('Listen-ID ungültig oder keine Rechte: ' + (e && e.message ? e.message : e));
            }
        } else {
            write('Suche Liste „' + cfg.listName + '" …');
            const found = await findListByTitle(token, siteId, cfg.listName);
            if (found && found.id) {
                listId = String(found.id);
                listWeb = found.webUrl ? String(found.webUrl) : '';
                listDisplayName = String(found.displayName || cfg.listName).trim();
                write('Bereits vorhanden – ID: ' + listId);
            } else {
                write('Erstelle Liste „' + cfg.listName + '" …');
                const created = await G.graphJson(
                    'POST',
                    G.graphPathSite(siteId) + '/lists',
                    token,
                    {
                        displayName: cfg.listName,
                        description: 'Freistellungen Lehrkräfte (Genehmigung nur Direktion)',
                        list: { template: 'genericList' }
                    },
                    'v1.0'
                );
                listId = created && created.id ? String(created.id) : '';
                listWeb = created && created.webUrl ? String(created.webUrl) : '';
                listDisplayName = String((created && created.displayName) || cfg.listName).trim();
                if (!listId) throw new Error('Listen-ID fehlt in der Antwort.');
                write('Liste angelegt, ID: ' + listId);
            }
        }

        write('Prüfe / ergänze Spalten …');
        await ensureColumns(siteId, listId, token, write);

        const flowAccount = effectiveFlowAccount(cfg);

        if (window.ms365LfrListPerms) {
            const lp = window.ms365LfrListPerms;
            if (typeof lp.apply === 'function') {
                try {
                    await lp.apply(cfg.siteUrl, listDisplayName, {
                        listId: listId,
                        flowServiceAccount: flowAccount
                    }, write);
                } catch (e) {
                    write('! Berechtigungen: ' + (e && e.message ? e.message : e));
                }
            } else if (flowAccount && typeof lp.grantFlowServiceAccount === 'function') {
                try {
                    await lp.grantFlowServiceAccount(
                        cfg.siteUrl,
                        listDisplayName,
                        flowAccount,
                        { listId: listId },
                        write
                    );
                } catch (e) {
                    write('! Flow-Technik Berechtigung: ' + (e && e.message ? e.message : e));
                }
            }
        }

        if ($('frListId')) $('frListId').value = listId;
        if ($('frListName') && listDisplayName) $('frListName').value = listDisplayName;
        const next = Object.assign({}, cfg, { listId: listId, listName: listDisplayName });
        saveCfg(next);
        refreshGlance();

        if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
            window.ms365ActionLog.append({
                tool: 'lehrer-freistellung-setup',
                action: 'ensure-list',
                target: cfg.siteUrl,
                summary: 'Freistellungsliste „' + listDisplayName + '“ (' + listId + ')'
            });
        }

        return { listId: listId, webUrl: listWeb, siteId: siteId, listDisplayName: listDisplayName };
    }

    async function resolveListDisplayNameForCfg(cfg) {
        const token = await ensureToken();
        const site = await G.resolveSiteFromWebUrl(token, cfg.siteUrl);
        const siteId = site && site.id ? String(site.id) : '';
        if (!siteId) throw new Error('Site-ID fehlt.');
        const listId = String(cfg.listId || '').trim();
        if (listId) {
            const existing = await G.graphJson(
                'GET',
                G.graphPathSite(siteId) +
                    '/lists/' +
                    encodeURIComponent(listId) +
                    '?$select=displayName',
                token,
                undefined,
                'v1.0'
            );
            const name = existing && existing.displayName ? String(existing.displayName).trim() : '';
            if (name) return name;
        }
        const found = await findListByTitle(token, siteId, cfg.listName);
        if (found && found.displayName) return String(found.displayName).trim();
        return String(cfg.listName || '').trim();
    }

    function replaceAll(haystack, needle, replacement) {
        if (!needle) return haystack;
        return String(haystack).split(needle).join(replacement);
    }

    /** Shared Mailbox → normales Senden als angemeldetes Technik-Konto (weniger Rechte-Probleme). */
    function convertSharedMailboxToUserSend(definitionRoot, useUserMailboxSend) {
        if (!useUserMailboxSend || !definitionRoot) return;

        function patchAction(action) {
            if (
                !action ||
                !action.inputs ||
                !action.inputs.host ||
                action.inputs.host.operationId !== 'SharedMailboxSendEmailV2'
            ) {
                return;
            }
            const params = action.inputs.parameters || {};
            delete params['emailMessage/MailboxAddress'];
            action.inputs.parameters = params;
            action.inputs.host.operationId = 'SendEmailV2';
        }

        function walkActions(actions) {
            if (!actions || typeof actions !== 'object') return;
            Object.keys(actions).forEach(function (key) {
                const action = actions[key];
                patchAction(action);
                if (action && action.actions) walkActions(action.actions);
                if (action && action.else && action.else.actions) walkActions(action.else.actions);
            });
        }

        walkActions(definitionRoot.actions);
    }

    function applyParamsToDefinition(rawDef, cfg) {
        let s = typeof rawDef === 'string' ? rawDef : JSON.stringify(rawDef);
        s = replaceAll(s, SOURCE.siteUrl, cfg.siteUrl);
        s = replaceAll(s, SOURCE.listId, cfg.listId);
        s = replaceAll(s, SOURCE.emailDirektion, cfg.emailDirektion);
        if (SOURCE.emailDirektor) {
            s = replaceAll(s, SOURCE.emailDirektor, cfg.emailDirektion);
        }
        s = replaceAll(s, SOURCE.emailMailbox, cfg.emailMailbox);
        // Ältere Flow-Exporte: frühere Sondergenehmigung → heute Direktion (v3)
        s = replaceAll(s, 'kurt@kurtsoeser.at', cfg.emailDirektion);
        const flowAccount = effectiveFlowAccount(cfg);
        if (flowAccount) {
            s = replaceAll(s, SOURCE.connectionOwner, flowAccount);
        }

        const obj = JSON.parse(s);
        if (obj.properties) {
            obj.properties.displayName = cfg.flowDisplayName;
            // Neue IDs, damit Import als „neu“ nicht mit dem Quell-Tenant kollidiert
            const newId = cryptoRandomGuid();
            obj.name = newId;
            obj.id = '/providers/Microsoft.Flow/flows/' + newId;
            if (obj.properties.definition && obj.properties.definition.metadata) {
                obj.properties.definition.metadata.creator = null;
                obj.properties.definition.metadata.lastModifiedBy = null;
            }
            if (obj.properties.definition) {
                convertSharedMailboxToUserSend(obj.properties.definition, cfg.mailAsTechnikUser !== false);
            }
        }
        return obj;
    }

    function cryptoRandomGuid() {
        if (window.crypto && typeof window.crypto.randomUUID === 'function') {
            return window.crypto.randomUUID();
        }
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
            const r = (Math.random() * 16) | 0;
            const v = c === 'x' ? r : (r & 0x3) | 0x8;
            return v.toString(16);
        });
    }

    async function fetchText(path) {
        const res = await fetch(path, { cache: 'no-store' });
        if (!res.ok) throw new Error('Vorlage nicht geladen: ' + path + ' (' + res.status + ')');
        return await res.text();
    }

    async function buildPackageZip(cfg) {
        if (typeof JSZip === 'undefined') {
            throw new Error('JSZip fehlt – Seite neu laden.');
        }
        if (!cfg.listId) throw new Error('Listen-ID fehlt – zuerst Liste anlegen oder ID eintragen.');
        if (!cfg.siteUrl) throw new Error('SharePoint-Website fehlt.');
        const flowAccount = effectiveFlowAccount(cfg);
        if (!flowAccount) {
            throw new Error('Technik-Konto fehlt – in Schritt 2 „Technik-Konto für Power Automate“ eintragen.');
        }

        const defPath = TEMPLATE_BASE + '/Microsoft.Flow/flows/' + FLOW_ASSET_ID + '/definition.json';
        const rootManifestPath = TEMPLATE_BASE + '/manifest.json';
        const flowsManifestPath = TEMPLATE_BASE + '/Microsoft.Flow/flows/manifest.json';
        const apisMapPath =
            TEMPLATE_BASE + '/Microsoft.Flow/flows/' + FLOW_ASSET_ID + '/apisMap.json';
        const connectionsMapPath =
            TEMPLATE_BASE + '/Microsoft.Flow/flows/' + FLOW_ASSET_ID + '/connectionsMap.json';

        log('Lade Flow-Vorlage …');
        const [defRaw, rootManifestRaw, flowsManifestRaw, apisMapRaw, connectionsMapRaw] =
            await Promise.all([
                fetchText(defPath),
                fetchText(rootManifestPath),
                fetchText(flowsManifestPath),
                fetchText(apisMapPath),
                fetchText(connectionsMapPath)
            ]);

        const definition = applyParamsToDefinition(defRaw, cfg);

        let rootManifest = JSON.parse(rootManifestRaw);
        rootManifest.details = rootManifest.details || {};
        rootManifest.details.displayName = cfg.flowDisplayName;
        rootManifest.details.description =
            'Lehrer-Freistellungen: Trigger Liste, eine Genehmigung Direktion, Audit-Felder – parametriert für Ziel-Tenant.';
        rootManifest.details.createdTime = new Date().toISOString();
        rootManifest.details.sourceEnvironment = '';

        // Connection-Anzeigenamen auf Technik-Konto (Ziel beim Import)
        Object.keys(rootManifest.resources || {}).forEach(function (key) {
            const r = rootManifest.resources[key];
            if (!r || !r.details) return;
            if (r.type === 'Microsoft.PowerApps/apis/connections') {
                if (String(r.details.displayName || '').indexOf('@') !== -1) {
                    r.details.displayName = flowAccount;
                }
            }
            if (r.type === 'Microsoft.Flow/flows') {
                r.details.displayName = cfg.flowDisplayName;
                r.suggestedCreationType = 'New';
            }
        });

        let rootStr = JSON.stringify(rootManifest);
        rootStr = replaceAll(rootStr, SOURCE.connectionOwner, flowAccount);
        rootManifest = JSON.parse(rootStr);

        const zip = new JSZip();
        zip.file('manifest.json', JSON.stringify(rootManifest, null, 2));
        zip.file('Microsoft.Flow/flows/manifest.json', flowsManifestRaw);
        const folder = zip.folder('Microsoft.Flow/flows/' + FLOW_ASSET_ID);
        folder.file('definition.json', JSON.stringify(definition, null, 2));
        folder.file('apisMap.json', apisMapRaw);
        folder.file('connectionsMap.json', connectionsMapRaw);

        log('Erzeuge ZIP …');
        const blob = await zip.generateAsync({ type: 'blob' });
        const stamp = new Date().toISOString().replace(/[:.]/g, '').slice(0, 15);
        const filename = 'Lehrer-Freistellung_' + stamp + '.zip';
        downloadBlob(blob, filename);
        log('Paket heruntergeladen: ' + filename);
        log('Nächster Schritt: make.powerautomate.com → Meine Flows → Importieren → Package (Legacy).');
        log(
            'Wichtig: In Power Automate mit ' +
                flowAccount +
                ' anmelden (Technik-Konto), dann alle Connections (SharePoint, Approvals, Outlook) mit diesem Konto verbinden.'
        );
        if (cfg.mailAsTechnikUser !== false) {
            log(
                'E-Mail im Paket: Senden als angemeldetes Technik-Konto (' +
                    flowAccount +
                    '), nicht als separates freigegebenes Postfach.'
            );
        } else {
            log(
                'E-Mail im Paket: freigegebenes Postfach „' +
                    cfg.emailMailbox +
                    '“ – Outlook-Connection von ' +
                    flowAccount +
                    ' braucht „Senden als“ auf dieses Postfach.'
            );
        }
        log('Flow danach einschalten und mit Testantrag prüfen.');
        return filename;
    }

    function downloadBlob(blob, filename) {
        const a = document.createElement('a');
        const url = URL.createObjectURL(blob);
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(function () {
            URL.revokeObjectURL(url);
        }, 2000);
    }

    async function onEnsureList() {
        try {
            const cfg = persistFromForm();
            log('— Liste —');
            const res = await createOrResolveList(cfg, log);
            toast('Liste bereit: ' + res.listId);
            if (res.webUrl) log('URL: ' + res.webUrl);
        } catch (e) {
            log('Fehler: ' + (e && e.message ? e.message : e));
            toast(e && e.message ? e.message : String(e));
        }
    }

    async function onBuildPackage() {
        try {
            const cfg = persistFromForm();
            if (!cfg.listId) {
                log('Keine Listen-ID – lege Liste zuerst an …');
                await createOrResolveList(cfg, log);
            }
            const fresh = readForm();
            log('— Flow-Paket —');
            await buildPackageZip(fresh);
            toast('Flow-Paket heruntergeladen.');
        } catch (e) {
            log('Fehler: ' + (e && e.message ? e.message : e));
            toast(e && e.message ? e.message : String(e));
        }
    }

    function flowImportedFlag() {
        try {
            return localStorage.getItem(FLOW_DONE_KEY) === '1';
        } catch (e) {
            return false;
        }
    }

    function onboardingProgress() {
        try {
            if (window.ms365PaOnboarding && typeof window.ms365PaOnboarding.progress === 'function') {
                return window.ms365PaOnboarding.progress();
            }
        } catch (e) {
            /* ignore */
        }
        return { done: 0, total: 0 };
    }

    function computeGlance() {
        const st = window.ms365LfrSetupStatus;
        const fn = st && typeof st.computeSetupGlance === 'function' ? st.computeSetupGlance : null;
        const cfg = readForm();
        const ob = onboardingProgress();
        if (fn) {
            return fn(cfg, {
                flowImported: flowImportedFlag(),
                onboardingDone: ob.done,
                onboardingTotal: ob.total
            });
        }
        return {
            emailsOk: !!(cfg.emailDirektion && cfg.emailMailbox && effectiveFlowAccount(cfg)),
            listOk: !!(cfg.siteUrl && cfg.listId),
            flowOk: flowImportedFlag(),
            prepOk: null
        };
    }

    function permsStepConfigured() {
        const st = window.ms365LfrSetupStatus;
        if (st && typeof st.permsStepConfigured === 'function') {
            return st.permsStepConfigured();
        }
        try {
            const p = JSON.parse(localStorage.getItem(PERMS_STORAGE_KEY) || '{}') || {};
            return !!(
                String(p.groupLehrerId || '').trim() ||
                String(p.groupDirektionId || '').trim()
            );
        } catch (e) {
            return false;
        }
    }

    function refreshGlance() {
        const g = computeGlance();
        const permsOk = permsStepConfigured();
        refreshImportAccountHint();
        const host = $('frSetupGlance');
        if (!host) return;
        host.querySelectorAll('[data-fr-glance]').forEach(function (card) {
            const key = card.getAttribute('data-fr-glance');
            let ok = false;
            let warn = false;
            if (key === 'prep') {
                if (g.prepOk === null) {
                    ok = false;
                    warn = false;
                } else {
                    ok = g.prepOk;
                    warn = !ok;
                }
            } else if (key === 'emails') {
                ok = g.emailsOk;
                warn = !ok;
            } else if (key === 'list') {
                ok = g.listOk;
                warn = !ok;
            } else if (key === 'perms' || key === 'planer') {
                ok = permsOk;
                warn = !ok;
            } else if (key === 'flow') {
                ok = g.flowOk;
                warn = !ok;
            }
            card.classList.toggle('is-ok', !!ok);
            card.classList.toggle('is-warn', !!warn && !ok);
            const val = card.querySelector('[data-fr-glance-value]');
            if (val) {
                if (key === 'prep' && g.prepOk === null) val.textContent = 'Optional';
                else val.textContent = ok ? 'Erledigt' : 'Offen';
            }
        });
    }

    function updatePhaseHint(step) {
        const el = $('frPhaseHint');
        if (!el) return;
        const g = computeGlance();
        if (step === 1) {
            const ob = onboardingProgress();
            if (ob.total && ob.done >= ob.total) {
                el.textContent = 'Vorbereitung abgeschlossen – weiter zu Konten (Schritt 2).';
            } else if (ob.total) {
                el.textContent =
                    'Schritt 1: Environment & Rechte (' +
                    ob.done +
                    '/' +
                    ob.total +
                    ' abgehakt). Oder direkt Schritt 2, wenn die IT schon fertig ist.';
            } else {
                el.textContent = 'Schritt 1: Schule vorbereiten (Checkliste).';
            }
        } else if (step === 2) {
            el.textContent = g.emailsOk
                ? 'Schritt 2: Konten sind gesetzt – weiter zur SharePoint-Liste.'
                : 'Schritt 2: Technik-Konto, Direktion und Absender eintragen (Suchen-Button).';
        } else if (step === 3) {
            el.textContent = g.listOk
                ? 'Schritt 3: Liste ist bereit – weiter zu Berechtigungen.'
                : 'Schritt 3: Website wählen und „Liste anlegen / prüfen“.';
        } else if (step === 4) {
            el.textContent = permsStepConfigured()
                ? 'Schritt 4: Berechtigungen gesetzt – optional Outlook-Kalender, dann Flow.'
                : 'Schritt 4: Entra-Gruppen Lehrkräfte / Direktion wählen und speichern.';
        } else {
            el.textContent = g.flowOk
                ? 'Schritt 5: Flow als importiert markiert – im Planer testen.'
                : 'Schritt 5: Flow-Paket laden und in Power Automate importieren.';
        }
    }

    function loadSetupStep() {
        try {
            const n = parseInt(localStorage.getItem(STEP_STORAGE_KEY) || '1', 10);
            if (n >= 1 && n <= SETUP_STEP_COUNT) return n;
        } catch (e) {
            /* ignore */
        }
        return 1;
    }

    function saveSetupStep(n) {
        try {
            localStorage.setItem(STEP_STORAGE_KEY, String(n));
        } catch (e) {
            /* ignore */
        }
    }

    function showSetupStep(n) {
        const step = Math.max(1, Math.min(SETUP_STEP_COUNT, parseInt(n, 10) || 1));
        saveSetupStep(step);
        for (let i = 1; i <= SETUP_STEP_COUNT; i++) {
            const panel = $('frSetupStep' + i);
            if (!panel) continue;
            const on = i === step;
            panel.hidden = !on;
            panel.setAttribute('aria-hidden', on ? 'false' : 'true');
        }
        document.querySelectorAll('#frSetupGlance [data-fr-setup-step]').forEach(function (btn) {
            const sn = parseInt(btn.getAttribute('data-fr-setup-step'), 10);
            const on = sn === step;
            btn.classList.toggle('is-active', on);
            btn.setAttribute('aria-selected', on ? 'true' : 'false');
            btn.setAttribute('tabindex', on ? '0' : '-1');
        });
        const back = $('frSetupBack');
        const next = $('frSetupNext');
        if (back) back.disabled = step <= 1;
        if (next) {
            next.textContent = step >= SETUP_STEP_COUNT ? 'Fertig' : 'Weiter';
            next.setAttribute(
                'aria-label',
                step >= SETUP_STEP_COUNT ? 'Setup abschließen' : 'Nächster Schritt'
            );
        }
        updatePhaseHint(step);
        refreshGlance();
        refreshImportAccountHint();
        if (step === 4 && typeof window.ms365LfrInitSetupPermissions === 'function') {
            window.ms365LfrInitSetupPermissions();
        }
    }

    function wireWizard() {
        document.querySelectorAll('[data-fr-setup-step]').forEach(function (btn) {
            btn.addEventListener('click', function () {
                showSetupStep(btn.getAttribute('data-fr-setup-step'));
            });
        });
        const back = $('frSetupBack');
        const next = $('frSetupNext');
        if (back) {
            back.addEventListener('click', function () {
                const cur = loadSetupStep();
                showSetupStep(Math.max(1, cur - 1));
            });
        }
        if (next) {
            next.addEventListener('click', function () {
                const cur = loadSetupStep();
                if (cur >= SETUP_STEP_COUNT) {
                    toast('Setup abgeschlossen – Lehrer-Freistellungs-Planer öffnen und testen.');
                    return;
                }
                if (cur === 1 || cur === 2) {
                    persistFromForm();
                }
                if (cur === 2) {
                    const cfg = readForm();
                    if (!cfg.emailDirektion || !effectiveFlowAccount(cfg)) {
                        toast('Tipp: Technik-Konto und Direktion eintragen, bevor Sie weitergehen.');
                    }
                }
                if (cur === 3) {
                    const cfg = readForm();
                    if (!cfg.listId) {
                        toast('Tipp: Zuerst „Liste anlegen / prüfen“, dann weiter.');
                    }
                }
                if (cur === 4 && !permsStepConfigured()) {
                    toast('Tipp: Lehrer- und Direktions-Gruppe wählen und speichern.');
                }
                showSetupStep(cur + 1);
            });
        }
        showSetupStep(loadSetupStep());
    }

    function wire() {
        const cfg = loadCfg();
        if (!cfg.flowServiceAccount && cfg.emailMailbox) {
            cfg.flowServiceAccount = String(cfg.emailMailbox).trim().toLowerCase();
        }
        writeForm(cfg);
        wireWizard();
        refreshGlance();
        const btnList = $('frBtnList');
        const btnPkg = $('frBtnPackage');
        const btnSave = $('frBtnSave');
        if (btnList) btnList.addEventListener('click', onEnsureList);
        const btnPerms = $('frBtnPerms');
        if (btnPerms) {
            btnPerms.addEventListener('click', async function () {
                const cfg = persistFromForm();
                if (!cfg.siteUrl || !cfg.listName) {
                    toast('Site-URL und Listenname fehlen.');
                    return;
                }
                try {
                    if (window.ms365LfrListPerms && window.ms365LfrListPerms.apply) {
                        const listTitle = await resolveListDisplayNameForCfg(cfg);
                        const listId = String(cfg.listId || ($('frListId') && $('frListId').value) || '').trim();
                        log('Berechtigungen für Liste „' + listTitle + '“ …');
                        await window.ms365LfrListPerms.apply(
                            cfg.siteUrl,
                            listTitle,
                            {
                                listId: listId,
                                flowServiceAccount: effectiveFlowAccount(cfg)
                            },
                            log
                        );
                        if ($('frListName') && listTitle) $('frListName').value = listTitle;
                        toast('Berechtigungen angewendet – Protokoll prüfen.');
                    }
                } catch (e) {
                    toast(String((e && e.message) || e));
                }
            });
        }
        if (btnPkg) btnPkg.addEventListener('click', onBuildPackage);
        if (btnSave) {
            btnSave.addEventListener('click', async function () {
                persistFromForm();
                refreshGlance();
                toast('Setup-Felder (Site, E-Mails, Liste, Kalender) in diesem Browser gespeichert.');
            });
        }
        const doneBox = $('frDone');
        if (doneBox) {
            try {
                doneBox.checked = flowImportedFlag();
            } catch (e) {}
            doneBox.addEventListener('change', function () {
                try {
                    if (doneBox.checked) localStorage.setItem(FLOW_DONE_KEY, '1');
                    else localStorage.removeItem(FLOW_DONE_KEY);
                } catch (e2) {}
                refreshGlance();
                updatePhaseHint(loadSetupStep());
                if (window.ms365BrowserBackup && typeof window.ms365BrowserBackup.notifyLocalDataChanged === 'function') {
                    window.ms365BrowserBackup.notifyLocalDataChanged('lfr-done');
                }
            });
        }
        [
            'frSiteUrl',
            'frListName',
            'frListId',
            'frEmailDirektion',
            'frEmailMailbox',
            'frFlowServiceAccount',
            'frFlowName',
            'frMailAsTechnikUser',
            'lfrOutlookCalUser',
            'lfrOutlookCalId'
        ].forEach(function (id) {
            const el = $(id);
            if (el) {
                el.addEventListener('change', function () {
                    persistFromForm();
                    refreshGlance();
                });
                el.addEventListener('input', refreshGlance);
            }
        });
        const syncMailboxBtn = $('frFlowAccountUseMailbox');
        if (syncMailboxBtn) {
            syncMailboxBtn.addEventListener('click', function () {
                const mb = String(($('frEmailMailbox') && $('frEmailMailbox').value) || '')
                    .trim()
                    .toLowerCase();
                if (!mb) {
                    toast('Zuerst Absender-Postfach eintragen.');
                    return;
                }
                if ($('frFlowServiceAccount')) $('frFlowServiceAccount').value = mb;
                persistFromForm();
                refreshGlance();
            });
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', wire);
    } else {
        wire();
    }

    window.ms365LfrSetup = {
        createOrResolveList: createOrResolveList,
        buildPackageZip: buildPackageZip,
        columnDefs: columnDefsLfr,
        ensureColumns: ensureColumns,
        loadSetupStep: loadSetupStep,
        showSetupStep: showSetupStep,
        refreshGlance: refreshGlance
    };
})();
