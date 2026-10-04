(function () {
    'use strict';

    const G = window.ms365SpoGraph;
    if (!G) return;

    const SCOPES = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Sites.ReadWrite.All'
    ];

    const STORAGE_KEY = 'ms365-freistellung-setup-v1';
    const STEP_STORAGE_KEY = 'ms365-freistellung-setup-step-v1';
    const FLOW_DONE_KEY = 'ms365-pa-done-freistellung';
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
        if (window.ms365BrowserBackup && typeof window.ms365BrowserBackup.notifyLocalDataChanged === 'function') {
            window.ms365BrowserBackup.notifyLocalDataChanged('freistellung-setup');
        }
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
        refreshGlance();
        return cfg;
    }

    async function ensureToken() {
        return await G.getGraphToken(SCOPES);
    }

    function columnDefsFreistellung() {
        const sch = window.ms365FreistellungSchema;
        if (sch && sch.columns && typeof sch.toGraphColumnBody === 'function') {
            return sch.columns.map(function (def) {
                return sch.toGraphColumnBody(def);
            });
        }
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

    async function ensureColumns(siteId, listId, token, write) {
        const logFn = typeof write === 'function' ? write : log;
        const base = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns';
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

        const defs = columnDefsFreistellung();
        const added = [];
        const skipped = [];
        for (let i = 0; i < defs.length; i++) {
            const d = defs[i];
            const key = String(d.name).toLowerCase();
            if (have[key]) {
                skipped.push(d.name);
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
        if (skipped.length === defs.length) {
            logFn('Alle Spalten vorhanden (' + skipped.join(', ') + ').');
        } else if (skipped.length) {
            logFn('Bereits vorhanden: ' + skipped.join(', '));
        }
        return { added: added, skipped: skipped };
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
                        description: 'Anträge auf Freistellung (Genehmigung KV + Direktion)',
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

        if (window.ms365FreistellungListPerms) {
            const lp = window.ms365FreistellungListPerms;
            if (typeof lp.apply === 'function') {
                try {
                    await lp.apply(cfg.siteUrl, listDisplayName, null, write);
                } catch (e) {
                    write('! Berechtigungen: ' + (e && e.message ? e.message : e));
                }
            }
            if (typeof lp.publishConfig === 'function') {
                try {
                    const pub = await lp.publishConfig(cfg.siteUrl, null, listId);
                    if (pub && pub.ok) {
                        write(
                            'Planer-Gruppen: ' +
                                (pub.listDescription ? 'Listen-Beschreibung + ' : '') +
                                (pub.path || 'ms365/freistellung-planer-groups.json') +
                                (pub.driveLabel ? ' (' + pub.driveLabel + ')' : '')
                        );
                    }
                } catch (e) {
                    write('! Planer-Gruppen-JSON: ' + (e && e.message ? e.message : e));
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
                tool: 'freistellung-setup',
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
        const st = window.ms365FreistellungSetupStatus;
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
            emailsOk: !!(cfg.emailDirektion && cfg.emailSonder && cfg.emailMailbox),
            listOk: !!(cfg.siteUrl && cfg.listId),
            flowOk: flowImportedFlag(),
            prepOk: null
        };
    }

    function refreshGlance() {
        const g = computeGlance();
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
                el.textContent = 'Vorbereitung abgeschlossen – weiter zu Liste und E-Mails.';
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
            el.textContent = g.listOk
                ? 'Schritt 2: Liste ist bereit – Einstellungen prüfen oder zu Schritt 3.'
                : 'Schritt 2: Website, Genehmiger und freigegebenes Postfach – dann „Liste anlegen / prüfen“.';
        } else {
            el.textContent = g.flowOk
                ? 'Schritt 3: Flow als importiert markiert – im Planer testen.'
                : 'Schritt 3: Paket laden und in Power Automate importieren.';
        }
    }

    function loadSetupStep() {
        try {
            const n = parseInt(localStorage.getItem(STEP_STORAGE_KEY) || '1', 10);
            if (n >= 1 && n <= 3) return n;
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
        const step = Math.max(1, Math.min(3, parseInt(n, 10) || 1));
        saveSetupStep(step);
        for (let i = 1; i <= 3; i++) {
            const panel = $('frSetupStep' + i);
            if (!panel) continue;
            const on = i === step;
            panel.hidden = !on;
            panel.setAttribute('aria-hidden', on ? 'false' : 'true');
        }
        document.querySelectorAll('[data-fr-setup-step]').forEach(function (btn) {
            const sn = parseInt(btn.getAttribute('data-fr-setup-step'), 10);
            const on = sn === step;
            btn.classList.toggle('active', on);
            btn.setAttribute('aria-selected', on ? 'true' : 'false');
            btn.setAttribute('tabindex', on ? '0' : '-1');
        });
        const back = $('frSetupBack');
        const next = $('frSetupNext');
        if (back) back.disabled = step <= 1;
        if (next) {
            next.textContent = step >= 3 ? 'Fertig' : 'Weiter';
            next.setAttribute('aria-label', step >= 3 ? 'Setup abschließen' : 'Nächster Schritt');
        }
        updatePhaseHint(step);
        refreshGlance();
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
                if (cur >= 3) {
                    toast('Setup abgeschlossen – Freistellungen-Planer öffnen und testen.');
                    return;
                }
                if (cur === 1) {
                    persistFromForm();
                }
                if (cur === 2) {
                    const cfg = readForm();
                    if (!cfg.listId) {
                        toast('Tipp: Zuerst „Liste anlegen / prüfen“, dann zu Schritt 3.');
                    }
                }
                showSetupStep(cur + 1);
            });
        }
        showSetupStep(loadSetupStep());
    }

    function wire() {
        writeForm(loadCfg());
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
                    if (window.ms365FreistellungListPerms && window.ms365FreistellungListPerms.apply) {
                        const listTitle = await resolveListDisplayNameForCfg(cfg);
                        log('Berechtigungen für Liste „' + listTitle + '“ …');
                        await window.ms365FreistellungListPerms.apply(cfg.siteUrl, listTitle, null, log);
                        if ($('frListName') && listTitle) $('frListName').value = listTitle;
                        const listId = String(cfg.listId || ($('frListId') && $('frListId').value) || '').trim();
                        if (
                            listId &&
                            window.ms365FreistellungListPerms.publishConfig &&
                            typeof window.ms365FreistellungListPerms.publishConfig === 'function'
                        ) {
                            try {
                                const pub = await window.ms365FreistellungListPerms.publishConfig(
                                    cfg.siteUrl,
                                    null,
                                    listId
                                );
                                if (pub && pub.ok) {
                                    log(
                                        'Planer-Gruppen: ' +
                                            (pub.listDescription ? 'Listen-Beschreibung + ' : '') +
                                            (pub.path || 'ms365/freistellung-planer-groups.json') +
                                            (pub.driveLabel ? ' (' + pub.driveLabel + ')' : '')
                                    );
                                }
                            } catch (e) {
                                log('! Planer-Gruppen-JSON: ' + (e && e.message ? e.message : e));
                            }
                        } else if (!listId) {
                            log(
                                '! Planer-Gruppen nicht veröffentlicht: Listen-ID fehlt – zuerst „Liste anlegen / prüfen“ oder „Alles speichern“.'
                            );
                        }
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
                let msg = 'Setup-Felder (Site, E-Mails, Liste) in diesem Browser gespeichert.';
                if (window.ms365FreistellungPlannerSave && window.ms365FreistellungPlannerSave.saveAndPublish) {
                    try {
                        const pub = await window.ms365FreistellungPlannerSave.saveAndPublish();
                        if (window.ms365FreistellungPlannerSave.formatToast) {
                            msg = window.ms365FreistellungPlannerSave.formatToast(pub);
                        } else if (pub && pub.ok) {
                            msg += ' Planer-Gruppen auf SharePoint veröffentlicht.';
                        }
                    } catch (e) {
                        msg +=
                            ' Planer-Gruppen: SharePoint-Fehler – ' +
                            String((e && e.message) || e);
                    }
                }
                toast(msg);
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
                    window.ms365BrowserBackup.notifyLocalDataChanged('freistellung-done');
                }
            });
        }
        ['frSiteUrl', 'frListName', 'frListId', 'frEmailDirektion', 'frEmailSonder', 'frEmailMailbox', 'frFlowName'].forEach(
            function (id) {
                const el = $(id);
                if (el) {
                    el.addEventListener('change', function () {
                        persistFromForm();
                        refreshGlance();
                    });
                    el.addEventListener('input', refreshGlance);
                }
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
        columnDefs: columnDefsFreistellung,
        ensureColumns: ensureColumns,
        showSetupStep: showSetupStep,
        refreshGlance: refreshGlance
    };
})();
