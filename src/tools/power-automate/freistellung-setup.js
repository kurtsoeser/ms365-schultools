(function () {
    'use strict';

    const G = window.ms365SpoGraph;
    if (!G) return;

    const SCOPES = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Sites.ReadWrite.All'
    ];

    const STORAGE_KEY = 'ms365-freistellung-setup-v1';
    const TEMPLATE_BASE = '../assets/power-automate/freistellung';
    const FLOW_ASSET_ID = 'c9164e06-4dbf-46f1-b99c-86d74bcdf8e4';

    /** Werte aus dem Original-Export (HAK Steyr) – werden ersetzt. */
    const SOURCE = {
        siteUrl: 'https://haksteyrat.sharepoint.com/sites/Administration',
        listId: '1f18f04b-c4e2-4845-92b6-190dae7e4411',
        emailDirektion: 'andreas.steininger@hak-steyr.at',
        emailSonder: 'ute.wiesmayr@hak-steyr.at',
        emailMailbox: 'automate@hak-steyr.at',
        connectionOwner: 'kurt.soeser@hak-steyr.at'
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
    }

    function readForm() {
        return {
            siteUrl: String(($('frSiteUrl') && $('frSiteUrl').value) || '').trim().replace(/\/$/, ''),
            listName: String(($('frListName') && $('frListName').value) || '').trim() || 'Freistellungen',
            listId: String(($('frListId') && $('frListId').value) || '').trim(),
            emailDirektion: String(($('frEmailDirektion') && $('frEmailDirektion').value) || '')
                .trim()
                .toLowerCase(),
            emailSonder: String(($('frEmailSonder') && $('frEmailSonder').value) || '')
                .trim()
                .toLowerCase(),
            emailMailbox: String(($('frEmailMailbox') && $('frEmailMailbox').value) || '')
                .trim()
                .toLowerCase(),
            flowDisplayName:
                String(($('frFlowName') && $('frFlowName').value) || '').trim() ||
                'Freistellungen - Genehmigungsprozess'
        };
    }

    function writeForm(cfg) {
        if (!cfg) return;
        if ($('frSiteUrl') && cfg.siteUrl) $('frSiteUrl').value = cfg.siteUrl;
        if ($('frListName') && cfg.listName) $('frListName').value = cfg.listName;
        if ($('frListId') && cfg.listId) $('frListId').value = cfg.listId;
        if ($('frEmailDirektion') && cfg.emailDirektion) $('frEmailDirektion').value = cfg.emailDirektion;
        if ($('frEmailSonder') && cfg.emailSonder) $('frEmailSonder').value = cfg.emailSonder;
        if ($('frEmailMailbox') && cfg.emailMailbox) $('frEmailMailbox').value = cfg.emailMailbox;
        if ($('frFlowName') && cfg.flowDisplayName) $('frFlowName').value = cfg.flowDisplayName;
    }

    function persistFromForm() {
        const cfg = readForm();
        saveCfg(cfg);
        return cfg;
    }

    async function ensureToken() {
        return await G.getGraphToken(SCOPES);
    }

    function columnDefsFreistellung() {
        return [
            {
                name: 'Beginn',
                displayName: 'Beginn',
                dateTime: { displayAs: 'default', format: 'dateOnly' }
            },
            {
                name: 'Ende',
                displayName: 'Ende',
                dateTime: { displayAs: 'default', format: 'dateOnly' }
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

    async function addColumns(siteId, listId, token) {
        const base = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns';
        const defs = columnDefsFreistellung();
        for (let i = 0; i < defs.length; i++) {
            await G.graphJson('POST', base, token, defs[i], 'v1.0');
            await G.sleep(140);
        }
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
        if (!cfg.emailDirektion || !cfg.emailSonder || !cfg.emailMailbox) {
            throw new Error('Bitte Direktion, Sondergenehmigung und freigegebenes Postfach ausfüllen.');
        }

        const token = await ensureToken();
        write('Löse Website auf …');
        const site = await G.resolveSiteFromWebUrl(token, cfg.siteUrl);
        const siteId = site && site.id ? String(site.id) : '';
        if (!siteId) throw new Error('Site-ID fehlt in der Graph-Antwort.');
        write('Site: ' + (site.displayName || siteId));

        let listId = cfg.listId;
        let listWeb = '';

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
                write('Liste gefunden: ' + (existing.displayName || listId));
            } catch (e) {
                throw new Error('Listen-ID ungültig oder keine Rechte: ' + (e && e.message ? e.message : e));
            }
        } else {
            write('Suche Liste „' + cfg.listName + '" …');
            const found = await findListByTitle(token, siteId, cfg.listName);
            if (found && found.id) {
                listId = String(found.id);
                listWeb = found.webUrl ? String(found.webUrl) : '';
                write('Bereits vorhanden – ID: ' + listId);
            } else {
                write('Erstelle Liste „' + cfg.listName + '" …');
                const created = await G.graphJson(
                    'POST',
                    G.graphPathSite(siteId) + '/lists',
                    token,
                    {
                        displayName: cfg.listName,
                        description: 'Anträge auf Freistellung (Genehmigung KV + Direktion)',
                        list: { template: 'genericList' }
                    },
                    'v1.0'
                );
                listId = created && created.id ? String(created.id) : '';
                listWeb = created && created.webUrl ? String(created.webUrl) : '';
                if (!listId) throw new Error('Listen-ID fehlt in der Antwort.');
                write('Liste angelegt, ID: ' + listId);
                write('Füge Spalten hinzu …');
                await addColumns(siteId, listId, token);
                write('Spalten fertig (Beginn, Ende, Status, Klasse, Klassenvorstand, Kategorie, Beschreibung, Bemerkungen).');
            }
        }

        if ($('frListId')) $('frListId').value = listId;
        const next = Object.assign({}, cfg, { listId: listId });
        saveCfg(next);

        if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
            window.ms365ActionLog.append({
                tool: 'freistellung-setup',
                action: 'ensure-list',
                target: cfg.siteUrl,
                summary: 'Freistellungsliste „' + cfg.listName + '“ (' + listId + ')'
            });
        }

        return { listId: listId, webUrl: listWeb, siteId: siteId };
    }

    function replaceAll(haystack, needle, replacement) {
        if (!needle) return haystack;
        return String(haystack).split(needle).join(replacement);
    }

    function applyParamsToDefinition(rawDef, cfg) {
        let s = typeof rawDef === 'string' ? rawDef : JSON.stringify(rawDef);
        s = replaceAll(s, SOURCE.siteUrl, cfg.siteUrl);
        s = replaceAll(s, SOURCE.listId, cfg.listId);
        s = replaceAll(s, SOURCE.emailDirektion, cfg.emailDirektion);
        s = replaceAll(s, SOURCE.emailSonder, cfg.emailSonder);
        s = replaceAll(s, SOURCE.emailMailbox, cfg.emailMailbox);

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
            'Freistellungen: SharePoint-Antrag → Genehmigung KV + Direktion (parametriert für Ziel-Tenant).';
        rootManifest.details.createdTime = new Date().toISOString();
        rootManifest.details.sourceEnvironment = '';

        // Connection-Anzeigenamen anonymisieren / auf aktuelle Schule hinweisen
        Object.keys(rootManifest.resources || {}).forEach(function (key) {
            const r = rootManifest.resources[key];
            if (!r || !r.details) return;
            if (r.type === 'Microsoft.PowerApps/apis/connections') {
                if (String(r.details.displayName || '').indexOf('@') !== -1) {
                    r.details.displayName = 'Ziel-Tenant Connection';
                }
            }
            if (r.type === 'Microsoft.Flow/flows') {
                r.details.displayName = cfg.flowDisplayName;
                r.suggestedCreationType = 'New';
            }
        });

        // Auch falls Owner-Mail noch im JSON steht
        let rootStr = JSON.stringify(rootManifest);
        rootStr = replaceAll(rootStr, SOURCE.connectionOwner, 'Ziel-Tenant Connection');
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
        const filename = 'Freistellung_' + stamp + '.zip';
        downloadBlob(blob, filename);
        log('Paket heruntergeladen: ' + filename);
        log('Nächster Schritt: make.powerautomate.com → Meine Flows → Importieren → Package (Legacy).');
        log('Dort Connections (SharePoint, Approvals, Outlook) dem Ziel-Tenant zuweisen und Flow einschalten.');
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

    function wire() {
        writeForm(loadCfg());
        const btnList = $('frBtnList');
        const btnPkg = $('frBtnPackage');
        const btnSave = $('frBtnSave');
        if (btnList) btnList.addEventListener('click', onEnsureList);
        if (btnPkg) btnPkg.addEventListener('click', onBuildPackage);
        if (btnSave) {
            btnSave.addEventListener('click', function () {
                persistFromForm();
                toast('Einstellungen gespeichert (dieser Browser).');
            });
        }
        const doneBox = $('frDone');
        if (doneBox) {
            try {
                doneBox.checked = localStorage.getItem('ms365-pa-done-freistellung') === '1';
            } catch (e) {}
            doneBox.addEventListener('change', function () {
                try {
                    if (doneBox.checked) localStorage.setItem('ms365-pa-done-freistellung', '1');
                    else localStorage.removeItem('ms365-pa-done-freistellung');
                } catch (e2) {}
            });
        }
        ['frSiteUrl', 'frListName', 'frListId', 'frEmailDirektion', 'frEmailSonder', 'frEmailMailbox', 'frFlowName'].forEach(
            function (id) {
                const el = $(id);
                if (el) el.addEventListener('change', persistFromForm);
            }
        );
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', wire);
    } else {
        wire();
    }

    window.ms365FreistellungSetup = {
        createOrResolveList: createOrResolveList,
        buildPackageZip: buildPackageZip,
        columnDefs: columnDefsFreistellung
    };
})();
