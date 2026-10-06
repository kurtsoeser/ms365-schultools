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
        const el = $('spsLog');
        if (!el) return;
        el.textContent += (el.textContent ? '\n' : '') + msg;
        el.scrollTop = el.scrollHeight;
    }

    function toast(m) {
        if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
        else window.alert(m);
    }

    function loadSettings() {
        if (typeof window.ms365TenantSettingsLoad !== 'function') {
            throw new Error('Stammdaten nicht geladen (tenant-settings-core.js fehlt?).');
        }
        return window.ms365TenantSettingsLoad() || {};
    }

    function readRunOpts() {
        const syncMode = !$('spsAlwaysNew') || !$('spsAlwaysNew').checked;
        return {
            syncMode: syncMode,
            removeOrphans: !$('spsRemoveOrphans') || $('spsRemoveOrphans').checked,
            klassenPersonen: !$('spsKlassenPersonen') || $('spsKlassenPersonen').checked
        };
    }

    async function ensureToken() {
        return await G.getGraphToken(SCOPES);
    }

    async function findListByDisplayName(token, siteId, listTitle) {
        const title = String(listTitle || '').trim();
        const path =
            G.graphPathSite(siteId) +
            '/lists?$filter=' +
            encodeURIComponent("displayName eq '" + title.replace(/'/g, "''") + "'") +
            '&$select=id,displayName,webUrl';
        const data = await G.graphJson('GET', path, token, undefined, 'v1.0');
        return ((data && data.value) || [])[0] || null;
    }

    async function createGenericList(siteId, listTitle, token, write) {
        const title = String(listTitle || '').trim();
        if (!title) throw new Error('Listenname fehlt.');
        write('Erstelle neue Liste „' + title + '" …');
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
        const listWeb = created && created.webUrl ? String(created.webUrl) : '';
        return { listId: listId, webUrl: listWeb, created: true };
    }

    async function ensureOrCreateList(siteId, listTitle, token, write, syncMode) {
        const title = String(listTitle || '').trim();
        if (!syncMode) {
            return await createGenericList(siteId, title, token, write);
        }
        const existing = await findListByDisplayName(token, siteId, title);
        if (existing && existing.id) {
            write('Liste „' + title + '" vorhanden – Abgleich der Zeilen.');
            return {
                listId: String(existing.id),
                webUrl: existing.webUrl ? String(existing.webUrl) : '',
                created: false
            };
        }
        return await createGenericList(siteId, title, token, write);
    }

    async function addMissingColumns(siteId, listId, token, defs) {
        const colsPath = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns?$top=200';
        const colsData = await G.graphJson('GET', colsPath, token, undefined, 'v1.0');
        const existing = new Set(
            ((colsData && colsData.value) || [])
                .map(function (c) {
                    return String((c && c.name) || '').trim();
                })
                .filter(Boolean)
        );
        const base = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns';
        for (let i = 0; i < defs.length; i++) {
            const def = defs[i];
            if (existing.has(def.name)) continue;
            await G.graphJson('POST', base, token, defs[i], 'v1.0');
            await G.sleep(120);
        }
    }

    async function fetchAllListItems(token, siteId, listId) {
        let path =
            G.graphPathSite(siteId) +
            '/lists/' +
            encodeURIComponent(listId) +
            '/items?$expand=fields&$top=200';
        const out = [];
        while (path) {
            const data = await G.graphJson('GET', path, token, undefined, 'v1.0');
            const rows = (data && data.value) || [];
            for (let i = 0; i < rows.length; i++) out.push(rows[i]);
            path = data && data['@odata.nextLink'] ? data['@odata.nextLink'] : '';
        }
        return out;
    }

    function fieldsEqual(a, b, keys) {
        for (let i = 0; i < keys.length; i++) {
            const k = keys[i];
            if (String(a[k] || '') !== String(b[k] || '')) return false;
        }
        return true;
    }

    /**
     * @param {object} params
     * @param {string} params.siteId
     * @param {string} params.listId
     * @param {string} params.token
     * @param {Array} params.rows
     * @param {function} params.keyFromRow
     * @param {function} params.keyFromItem
     * @param {function} params.fieldsFromRow
     * @param {string[]} params.compareKeys
     * @param {boolean} params.removeOrphans
     * @param {function} params.write
     * @param {string} params.label
     */
    async function syncRowsToList(params) {
        const write = params.write;
        const items = await fetchAllListItems(params.token, params.siteId, params.listId);
        const byKey = new Map();
        items.forEach(function (it) {
            const fields = (it && it.fields) || {};
            const key = params.keyFromItem(fields);
            if (!key) return;
            if (!byKey.has(key)) byKey.set(key, it);
        });

        const desiredKeys = new Set();
        let added = 0;
        let updated = 0;
        let unchanged = 0;
        const itemsPath =
            G.graphPathSite(params.siteId) + '/lists/' + encodeURIComponent(params.listId) + '/items';

        for (let i = 0; i < params.rows.length; i++) {
            const row = params.rows[i];
            const key = params.keyFromRow(row);
            if (!key) continue;
            desiredKeys.add(key);
            const nextFields = await Promise.resolve(params.fieldsFromRow(row));
            const existing = byKey.get(key);
            if (!existing) {
                await G.graphJson('POST', itemsPath, params.token, { fields: nextFields }, 'v1.0');
                added++;
                if (added % 25 === 0) write('… ' + added + ' neu');
                await G.sleep(80);
                continue;
            }
            const prev = existing.fields || {};
            const isSame =
                typeof params.fieldsEqual === 'function'
                    ? params.fieldsEqual(prev, nextFields)
                    : fieldsEqual(prev, nextFields, params.compareKeys);
            if (isSame) {
                unchanged++;
                continue;
            }
            const itemId = existing.id ? String(existing.id) : '';
            if (!itemId) continue;
            const patchPath = itemsPath + '/' + encodeURIComponent(itemId) + '/fields';
            await G.graphJson('PATCH', patchPath, params.token, nextFields, 'v1.0');
            updated++;
            if (updated % 25 === 0) write('… ' + updated + ' aktualisiert');
            await G.sleep(80);
        }

        let deleted = 0;
        if (params.removeOrphans) {
            const keys = Array.from(byKey.keys());
            for (let k = 0; k < keys.length; k++) {
                const key = keys[k];
                if (desiredKeys.has(key)) continue;
                const it = byKey.get(key);
                const itemId = it && it.id ? String(it.id) : '';
                if (!itemId) continue;
                await G.graphJson(
                    'DELETE',
                    itemsPath + '/' + encodeURIComponent(itemId),
                    params.token,
                    undefined,
                    'v1.0'
                );
                deleted++;
                await G.sleep(60);
            }
        }

        write(
            params.label +
                ': ' +
                added +
                ' neu, ' +
                updated +
                ' geändert, ' +
                unchanged +
                ' unverändert' +
                (params.removeOrphans ? ', ' + deleted + ' entfernt (nicht mehr in Stammdaten)' : '')
        );
        return { added: added, updated: updated, unchanged: unchanged, deleted: deleted };
    }

    async function postAllRows(siteId, listId, token, rows, fieldsFn, write, label) {
        const itemsPath = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/items';
        let ok = 0;
        for (let i = 0; i < rows.length; i++) {
            await G.graphJson('POST', itemsPath, token, { fields: fieldsFn(rows[i]) }, 'v1.0');
            ok++;
            if (ok % 25 === 0) write('… ' + ok + ' ' + label);
            await G.sleep(80);
        }
        write('Fertig: ' + ok + ' Zeilen geschrieben (' + label + ').');
        return ok;
    }

    async function resolveSite(webUrl, token, write) {
        write('Löse Website auf …');
        const site = await G.resolveSiteFromWebUrl(token, webUrl);
        const siteId = site && site.id ? String(site.id) : '';
        if (!siteId) throw new Error('Site-ID fehlt in der Graph-Antwort.');
        write('Site: ' + (site.displayName || siteId));
        return siteId;
    }

    function splitFirstLastName(displayName) {
        const s = String(displayName || '').trim();
        if (!s) return { vorname: '', nachname: '' };
        const parts = s.split(/\s+/).filter(Boolean);
        if (parts.length <= 1) return { vorname: parts[0] || '', nachname: '' };
        return { vorname: parts[0], nachname: parts.slice(1).join(' ') };
    }

    function studentKeyFromParts(klasse, name, email, externalId) {
        const ext = String(externalId || '').trim();
        if (ext) return 'id:' + ext.toLowerCase();
        const em = String(email || '').trim().toLowerCase();
        if (em) return 'mail:' + em;
        const k = String(klasse || '').trim().toLowerCase();
        const n = String(name || '').trim().toLowerCase();
        if (k || n) return 'kn:' + k + '\u0001' + n;
        return '';
    }

    function studentFields(st) {
        const name = String(st.name || '').trim() || String(st.email || '').trim() || '—';
        const parts = splitFirstLastName(st.name || name);
        const email = String(st.email || '').trim();
        const fields = {
            Title: name,
            Vorname: parts.vorname,
            Nachname: parts.nachname,
            Klasse: String(st.klasse || '').trim(),
            EMail: email,
            UPN: email
        };
        const ext = String(st.externalId || st.id || '').trim();
        if (ext) fields.ExternalId = ext;
        return fields;
    }

    const COL_SCHUELER = [
        { name: 'Vorname', displayName: 'Vorname', text: { allowMultipleLines: false, maxLength: 120 } },
        { name: 'Nachname', displayName: 'Nachname', text: { allowMultipleLines: false, maxLength: 120 } },
        { name: 'Klasse', displayName: 'Klasse', text: { allowMultipleLines: false, maxLength: 40 } },
        { name: 'EMail', displayName: 'E-Mail', text: { allowMultipleLines: false, maxLength: 255 } },
        { name: 'UPN', displayName: 'UPN', text: { allowMultipleLines: false, maxLength: 255 } },
        { name: 'ExternalId', displayName: 'Externe ID', text: { allowMultipleLines: false, maxLength: 80 } },
        {
            name: 'Schueler',
            displayName: 'Schüler (Person)',
            personOrGroup: { allowMultipleSelection: false, chooseFromType: 'peopleOnly' }
        }
    ];

    const COL_FACH = [
        { name: 'FachCode', displayName: 'Kürzel', text: { allowMultipleLines: false, maxLength: 40 } }
    ];

    const COL_KLASSE = [
        { name: 'KlassenCode', displayName: 'Kürzel', text: { allowMultipleLines: false, maxLength: 40 } },
        { name: 'Abschlussjahr', displayName: 'Abschlussjahr', text: { allowMultipleLines: false, maxLength: 4 } },
        { name: 'KVName', displayName: 'Klassenvorstand', text: { allowMultipleLines: false, maxLength: 120 } },
        { name: 'KVEmail', displayName: 'KV E-Mail', text: { allowMultipleLines: false, maxLength: 255 } }
    ];

    const COL_KLASSE_PERSON = COL_KLASSE.concat([
        {
            name: 'Schuelerinnen',
            displayName: 'Schülerinnen',
            personOrGroup: { allowMultipleSelection: true, chooseFromType: 'peopleOnly' }
        },
        {
            name: 'Klassenvorstand',
            displayName: 'KV (Person)',
            personOrGroup: { allowMultipleSelection: false, chooseFromType: 'peopleOnly' }
        }
    ]);

    function studentBelongsToClass(st, c) {
        const sk = String(st && st.klasse || '').trim().toLowerCase();
        if (!sk) return false;
        const code = String(c && c.code || '').trim().toLowerCase();
        const name = String(c && c.name || '').trim().toLowerCase();
        return (code && sk === code) || (name && sk === name);
    }

    function personLookupFingerprint(fields, columnName) {
        const name = String(columnName || '').trim();
        const lidKey = name + 'LookupId';
        const rawIds = fields && fields[lidKey];
        if (Array.isArray(rawIds)) {
            return rawIds
                .map(function (x) {
                    return String(x);
                })
                .sort()
                .join(',');
        }
        if (rawIds != null && rawIds !== '') return String(rawIds);
        const raw = fields && fields[name];
        if (Array.isArray(raw)) {
            return raw
                .map(function (x) {
                    return x && x.LookupId != null ? String(x.LookupId) : '';
                })
                .filter(Boolean)
                .sort()
                .join(',');
        }
        if (raw && typeof raw === 'object' && raw.LookupId != null) return String(raw.LookupId);
        return '';
    }

    const COL_ARGE = [
        { name: 'ArgeCode', displayName: 'Kürzel', text: { allowMultipleLines: false, maxLength: 40 } },
        {
            name: 'Faecher',
            displayName: 'Fächer (Kürzel)',
            text: { allowMultipleLines: true, maxLength: 2000 }
        }
    ];

    async function runListPipeline(webUrl, listTitle, defaultTitle, rows, columnDefs, syncConfig, logFn, runOpts, actionName, summaryLabel) {
        const write = typeof logFn === 'function' ? logFn : log;
        const url = String(webUrl || '').trim();
        const title = String(listTitle || '').trim() || defaultTitle;
        if (!url) throw new Error('Bitte die Adresse der SharePoint-Website eintragen.');
        if (!rows.length) throw new Error('Keine Daten in den Stammdaten für „' + summaryLabel + '“.');

        write(summaryLabel + ' aus lokalem Speicher: ' + rows.length);
        const token = await ensureToken();
        const siteId = await resolveSite(url, token, write);
        const syncMode = runOpts && runOpts.syncMode !== false;
        const removeOrphans = !(runOpts && runOpts.removeOrphans === false);

        const listMeta = await ensureOrCreateList(siteId, title, token, write, syncMode);
        write('Prüfe Spalten …');
        await addMissingColumns(siteId, listMeta.listId, token, columnDefs);

        if (syncMode) {
            await syncRowsToList({
                siteId: siteId,
                listId: listMeta.listId,
                token: token,
                rows: rows,
                keyFromRow: syncConfig.keyFromRow,
                keyFromItem: syncConfig.keyFromItem,
                fieldsFromRow: syncConfig.fieldsFromRow,
                compareKeys: syncConfig.compareKeys,
                removeOrphans: removeOrphans,
                write: write,
                label: summaryLabel
            });
        } else {
            await postAllRows(siteId, listMeta.listId, token, rows, syncConfig.fieldsFromRow, write, summaryLabel);
        }

        if (listMeta.webUrl) write('Liste im Browser: ' + listMeta.webUrl);
        if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
            window.ms365ActionLog.append({
                tool: 'sharepoint',
                action: actionName,
                target: url,
                summary: summaryLabel + ' „' + title + '“ (' + rows.length + ' Stammdaten-Zeilen)'
            });
        }
        return { listId: listMeta.listId, webUrl: listMeta.webUrl, count: rows.length };
    }

    async function syncSchuelerList(webUrl, listTitle, logFn, runOpts) {
        const write = typeof logFn === 'function' ? logFn : log;
        const url = String(webUrl || '').trim();
        const title = String(listTitle || '').trim() || 'Schülerinnen';
        if (!url) throw new Error('Bitte die Adresse der SharePoint-Website eintragen.');

        const ro = runOpts && typeof runOpts === 'object' ? runOpts : readRunOpts();
        const s = loadSettings();
        const students = (Array.isArray(s.students) ? s.students : []).filter(function (st) {
            return st && (String(st.name || '').trim() || String(st.email || '').trim() || String(st.klasse || '').trim());
        });
        if (!students.length) throw new Error('Keine Schüler:innen in den Stammdaten – zuerst dort pflegen.');

        write('Schülerinnen aus lokalem Speicher: ' + students.length + ' (mit Personenfeld Schüler)');
        const token = await ensureToken();
        const siteId = await resolveSite(url, token, write);
        const syncMode = ro.syncMode !== false;
        const removeOrphans = ro.removeOrphans !== false;

        const listMeta = await ensureOrCreateList(siteId, title, token, write, syncMode);
        write('Prüfe Spalten (Vorname, Nachname, Personenfeld Schüler, …) …');
        await addMissingColumns(siteId, listMeta.listId, token, COL_SCHUELER);

        let resolver = null;
        if (typeof G.createSitePersonResolver === 'function') {
            write('Personen: User Information List / ensureUser …');
            resolver = await G.createSitePersonResolver(url, token, siteId, { write: write, ensureMissing: true });
        } else {
            write('Hinweis: Personen-Hilfen fehlen – Textfelder werden gesetzt, Personenfeld bleibt leer.');
        }

        const compareKeys = ['Title', 'Vorname', 'Nachname', 'Klasse', 'EMail', 'UPN', 'ExternalId'];
        let personMiss = 0;

        async function studentFieldsWithPerson(st) {
            const fields = studentFields(st);
            const email = String(st.email || '').trim();
            if (!email || !resolver) return fields;
            const id = await resolver.lookupIdForEmail(email);
            if (id && typeof G.applySinglePersonLookupField === 'function') {
                G.applySinglePersonLookupField(fields, 'Schueler', id);
            } else if (email) {
                personMiss++;
            }
            return fields;
        }

        function studentFieldsEqual(prev, next) {
            for (let i = 0; i < compareKeys.length; i++) {
                const k = compareKeys[i];
                if (String(prev[k] || '') !== String(next[k] || '')) return false;
            }
            if (personLookupFingerprint(prev, 'Schueler') !== personLookupFingerprint(next, 'Schueler')) {
                return false;
            }
            return true;
        }

        const syncConfig = {
            keyFromRow: function (st) {
                return studentKeyFromParts(st.klasse, st.name, st.email, st.externalId || st.id);
            },
            keyFromItem: function (f) {
                return studentKeyFromParts(f.Klasse, f.Title, f.EMail, f.ExternalId);
            },
            fieldsFromRow: studentFieldsWithPerson,
            fieldsEqual: studentFieldsEqual,
            compareKeys: compareKeys
        };

        if (syncMode) {
            await syncRowsToList({
                siteId: siteId,
                listId: listMeta.listId,
                token: token,
                rows: students,
                keyFromRow: syncConfig.keyFromRow,
                keyFromItem: syncConfig.keyFromItem,
                fieldsFromRow: syncConfig.fieldsFromRow,
                fieldsEqual: syncConfig.fieldsEqual,
                compareKeys: syncConfig.compareKeys,
                removeOrphans: removeOrphans,
                write: write,
                label: 'Schülerinnen'
            });
        } else {
            const itemsPath = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listMeta.listId) + '/items';
            let ok = 0;
            for (let i = 0; i < students.length; i++) {
                const fields = await studentFieldsWithPerson(students[i]);
                await G.graphJson('POST', itemsPath, token, { fields: fields }, 'v1.0');
                ok++;
                if (ok % 25 === 0) write('… ' + ok + ' Zeilen geschrieben');
                await G.sleep(80);
            }
            write('Fertig: ' + ok + ' Schüler:innen als Listenelemente.');
        }

        if (personMiss) {
            write(
                'Hinweis: ' +
                    personMiss +
                    ' Schüler:innen ohne Personen-Verknüpfung (E-Mail fehlt oder Benutzer auf der Site nicht auflösbar).'
            );
        }

        if (listMeta.webUrl) write('Liste im Browser: ' + listMeta.webUrl);
        if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
            window.ms365ActionLog.append({
                tool: 'sharepoint',
                action: 'sync-schueler-list',
                target: url,
                summary: 'Schülerinnen „' + title + '“ (' + students.length + ' Stammdaten-Zeilen)'
            });
        }
        return { listId: listMeta.listId, webUrl: listMeta.webUrl, count: students.length };
    }

    async function syncFaecherList(webUrl, listTitle, logFn, runOpts) {
        const s = loadSettings();
        const subjects = (Array.isArray(s.subjects) ? s.subjects : []).filter(function (sub) {
            return sub && String(sub.code || '').trim();
        });
        const compareKeys = ['Title', 'FachCode'];
        return await runListPipeline(
            webUrl,
            listTitle,
            'Fächer',
            subjects,
            COL_FACH,
            {
                keyFromRow: function (sub) {
                    return String(sub.code || '').trim().toLowerCase();
                },
                keyFromItem: function (f) {
                    const c = String(f.FachCode || '').trim().toLowerCase();
                    return c || String(f.Title || '').trim().toLowerCase();
                },
                fieldsFromRow: function (sub) {
                    const code = String(sub.code || '').trim();
                    const name = String(sub.name || '').trim() || code;
                    return { Title: name, FachCode: code };
                },
                compareKeys: compareKeys
            },
            logFn,
            runOpts,
            'sync-faecher-list',
            'Fächer'
        );
    }

    async function syncFachgruppenList(webUrl, listTitle, logFn, runOpts) {
        const s = loadSettings();
        const subjects = (Array.isArray(s.subjects) ? s.subjects : []).filter(function (sub) {
            return sub && String(sub.code || '').trim();
        });
        const compareKeys = ['Title', 'FachgruppeCode'];
        return await runListPipeline(
            webUrl,
            listTitle,
            'Fachgruppen',
            subjects,
            [{ name: 'FachgruppeCode', displayName: 'Kürzel', text: { allowMultipleLines: false, maxLength: 40 } }],
            {
                keyFromRow: function (sub) {
                    return String(sub.code || '').trim().toLowerCase();
                },
                keyFromItem: function (f) {
                    const c = String(f.FachgruppeCode || f.FachCode || '').trim().toLowerCase();
                    return c || String(f.Title || '').trim().toLowerCase();
                },
                fieldsFromRow: function (sub) {
                    const code = String(sub.code || '').trim();
                    const name = String(sub.name || '').trim() || code;
                    return { Title: name, FachgruppeCode: code };
                },
                compareKeys: compareKeys
            },
            logFn,
            runOpts,
            'sync-fachgruppen-list',
            'Fachgruppen'
        );
    }

    async function syncKlassenList(webUrl, listTitle, logFn, runOpts) {
        const write = typeof logFn === 'function' ? logFn : log;
        const url = String(webUrl || '').trim();
        const title = String(listTitle || '').trim() || 'Klassen';
        if (!url) throw new Error('Bitte die Adresse der SharePoint-Website eintragen.');

        const ro = runOpts && typeof runOpts === 'object' ? runOpts : readRunOpts();
        const withPersons = ro.klassenPersonen !== false;

        const s = loadSettings();
        const classes = (Array.isArray(s.classes) ? s.classes : []).filter(function (c) {
            return c && (String(c.code || '').trim() || String(c.name || '').trim());
        });
        if (!classes.length) throw new Error('Keine Klassen in den Stammdaten – zuerst dort pflegen.');

        const students = (Array.isArray(s.students) ? s.students : []).filter(function (st) {
            return st && String(st.email || '').trim();
        });

        write(
            'Klassen: ' +
                classes.length +
                (withPersons ? ' (mit Personenfeld Schülerinnen, ' + students.length + ' Schüler mit E-Mail)' : '')
        );
        const token = await ensureToken();
        const siteId = await resolveSite(url, token, write);
        const syncMode = ro.syncMode !== false;
        const removeOrphans = ro.removeOrphans !== false;

        const listMeta = await ensureOrCreateList(siteId, title, token, write, syncMode);
        const columnDefs = withPersons ? COL_KLASSE_PERSON : COL_KLASSE;
        write('Prüfe Spalten …');
        await addMissingColumns(siteId, listMeta.listId, token, columnDefs);

        let resolver = null;
        if (withPersons && typeof G.createSitePersonResolver === 'function') {
            write('Personen: User Information List / ensureUser …');
            resolver = await G.createSitePersonResolver(url, token, siteId, { write: write, ensureMissing: true });
        } else if (withPersons) {
            write('Hinweis: Personen-Hilfen fehlen – nur Textfelder werden gesetzt.');
        }

        const compareKeys = ['Title', 'KlassenCode', 'Abschlussjahr', 'KVName', 'KVEmail'];

        async function fieldsForClass(c) {
            const code = String(c.code || '').trim();
            const name = String(c.name || '').trim() || code;
            const fields = {
                Title: name,
                KlassenCode: code,
                Abschlussjahr: String(c.year || '').trim(),
                KVName: String(c.headName || '').trim(),
                KVEmail: String(c.headEmail || '').trim()
            };
            if (!withPersons || !resolver) return fields;

            const lookupIds = [];
            let missingMail = 0;
            for (let i = 0; i < students.length; i++) {
                const st = students[i];
                if (!studentBelongsToClass(st, c)) continue;
                const em = String(st.email || '').trim();
                if (!em) continue;
                const id = await resolver.lookupIdForEmail(em);
                if (id) lookupIds.push(id);
                else missingMail++;
            }
            if (missingMail) {
                write(
                    'Klasse ' +
                        (code || name) +
                        ': ' +
                        missingMail +
                        ' Schüler mit E-Mail konnten auf der Site nicht verknüpft werden (ensureUser/Rechte?).'
                );
            }
            G.applyMultiPersonLookupFields(fields, 'Schuelerinnen', lookupIds);

            const kvEm = String(c.headEmail || '').trim();
            if (kvEm) {
                const kvId = await resolver.lookupIdForEmail(kvEm);
                if (kvId) G.applySinglePersonLookupField(fields, 'Klassenvorstand', kvId);
            }
            return fields;
        }

        function fieldsEqualKlassen(prev, next) {
            for (let i = 0; i < compareKeys.length; i++) {
                const k = compareKeys[i];
                if (String(prev[k] || '') !== String(next[k] || '')) return false;
            }
            if (!withPersons) return true;
            if (
                personLookupFingerprint(prev, 'Schuelerinnen') !==
                personLookupFingerprint(next, 'Schuelerinnen')
            )
                return false;
            if (
                personLookupFingerprint(prev, 'Klassenvorstand') !==
                personLookupFingerprint(next, 'Klassenvorstand')
            )
                return false;
            return true;
        }

        const syncConfig = {
            keyFromRow: function (c) {
                const code = String(c.code || '').trim().toLowerCase();
                if (code) return code;
                return String(c.name || '').trim().toLowerCase();
            },
            keyFromItem: function (f) {
                const code = String(f.KlassenCode || '').trim().toLowerCase();
                if (code) return code;
                return String(f.Title || '').trim().toLowerCase();
            },
            fieldsFromRow: fieldsForClass,
            fieldsEqual: fieldsEqualKlassen,
            compareKeys: compareKeys
        };

        if (syncMode) {
            await syncRowsToList({
                siteId: siteId,
                listId: listMeta.listId,
                token: token,
                rows: classes,
                keyFromRow: syncConfig.keyFromRow,
                keyFromItem: syncConfig.keyFromItem,
                fieldsFromRow: syncConfig.fieldsFromRow,
                fieldsEqual: syncConfig.fieldsEqual,
                compareKeys: syncConfig.compareKeys,
                removeOrphans: removeOrphans,
                write: write,
                label: 'Klassen'
            });
        } else {
            const itemsPath = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listMeta.listId) + '/items';
            let ok = 0;
            for (let i = 0; i < classes.length; i++) {
                const fields = await fieldsForClass(classes[i]);
                await G.graphJson('POST', itemsPath, token, { fields: fields }, 'v1.0');
                ok++;
                await G.sleep(80);
            }
            write('Fertig: ' + ok + ' Klassen als Listenelemente.');
        }

        if (listMeta.webUrl) write('Liste im Browser: ' + listMeta.webUrl);
        if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
            window.ms365ActionLog.append({
                tool: 'sharepoint',
                action: 'sync-klassen-list',
                target: url,
                summary: 'Klassenliste „' + title + '“' + (withPersons ? ' inkl. Personenfeld' : '')
            });
        }
        return { listId: listMeta.listId, webUrl: listMeta.webUrl, count: classes.length };
    }

    async function syncArgeList(webUrl, listTitle, logFn, runOpts) {
        const s = loadSettings();
        const arges = (Array.isArray(s.arges) ? s.arges : []).filter(function (a) {
            return a && String(a.code || '').trim();
        });
        const compareKeys = ['Title', 'ArgeCode', 'Faecher'];
        return await runListPipeline(
            webUrl,
            listTitle,
            'ARGEs',
            arges,
            COL_ARGE,
            {
                keyFromRow: function (a) {
                    return String(a.code || '').trim().toLowerCase();
                },
                keyFromItem: function (f) {
                    const c = String(f.ArgeCode || '').trim().toLowerCase();
                    return c || String(f.Title || '').trim().toLowerCase();
                },
                fieldsFromRow: function (a) {
                    const code = String(a.code || '').trim();
                    const name = String(a.name || '').trim() || code;
                    const faecher = (Array.isArray(a.subjects) ? a.subjects : [])
                        .map(function (x) {
                            return String(x || '').trim();
                        })
                        .filter(Boolean)
                        .join('; ');
                    return { Title: name, ArgeCode: code, Faecher: faecher };
                },
                compareKeys: compareKeys
            },
            logFn,
            runOpts,
            'sync-arge-list',
            'ARGEs'
        );
    }

    /** @deprecated Alias – nutzt Abgleich-Modus */
    async function createSchuelerList(webUrl, listTitle, logFn) {
        return await syncSchuelerList(webUrl, listTitle, logFn, readRunOpts());
    }
    async function createFaecherList(webUrl, listTitle, logFn) {
        return await syncFaecherList(webUrl, listTitle, logFn, readRunOpts());
    }
    async function createKlassenList(webUrl, listTitle, logFn) {
        return await syncKlassenList(webUrl, listTitle, logFn, readRunOpts());
    }

    /**
     * @param {string} webUrl
     * @param {object} opts
     * @param {function} [logFn]
     * @param {{ syncMode?: boolean, removeOrphans?: boolean }} [runOpts]
     */
    async function syncSelectedLists(webUrl, opts, logFn, runOpts) {
        const write = typeof logFn === 'function' ? logFn : log;
        const o = opts && typeof opts === 'object' ? opts : {};
        const ro = runOpts && typeof runOpts === 'object' ? runOpts : readRunOpts();
        const results = {};
        if (o.schueler) {
            write('—— Schülerinnen ——');
            results.schueler = await syncSchuelerList(webUrl, o.schuelerTitle || 'Schülerinnen', write, ro);
        }
        if (o.faecher) {
            write('—— Fächer ——');
            results.faecher = await syncFaecherList(webUrl, o.faecherTitle || 'Fächer', write, ro);
        }
        if (o.fachgruppen) {
            write('—— Fachgruppen ——');
            results.fachgruppen = await syncFachgruppenList(webUrl, o.fachgruppenTitle || 'Fachgruppen', write, ro);
        }
        if (o.arges) {
            write('—— ARGEs ——');
            results.arges = await syncArgeList(webUrl, o.argesTitle || 'ARGEs', write, ro);
        }
        if (o.klassen) {
            write('—— Klassen ——');
            results.klassen = await syncKlassenList(webUrl, o.klassenTitle || 'Klassen', write, ro);
        }
        if (o.lehrer) {
            write('—— Lehrerinnen ——');
            const lehrerApi = window.ms365SpoLehrerListe;
            if (!lehrerApi || typeof lehrerApi.createList !== 'function') {
                throw new Error('Lehrerlisten-Sync nicht geladen (sharepoint-liste-lehrer.js).');
            }
            const lehrerRun = {
                syncMode: ro.syncMode,
                removeOrphans: ro.removeOrphans,
                personField: ro.lehrerPersonen !== false,
                clearListFirst: false
            };
            results.lehrer = await lehrerApi.createList(
                webUrl,
                o.lehrerTitle || 'Lehrerinnen',
                write,
                lehrerRun
            );
        }
        if (!o.schueler && !o.faecher && !o.fachgruppen && !o.arges && !o.klassen && !o.lehrer) {
            throw new Error('Mindestens eine Liste auswählen.');
        }
        return results;
    }

    async function createSelectedLists(webUrl, opts, logFn) {
        return await syncSelectedLists(webUrl, opts, logFn, readRunOpts());
    }

    function collectOptsFromForm() {
        return {
            schueler: $('spsWantSchueler') && $('spsWantSchueler').checked,
            faecher: $('spsWantFaecher') && $('spsWantFaecher').checked,
            fachgruppen: $('spsWantFachgruppen') && $('spsWantFachgruppen').checked,
            arges: $('spsWantArge') && $('spsWantArge').checked,
            klassen: $('spsWantKlassen') && $('spsWantKlassen').checked,
            schuelerTitle: String($('spsSchuelerName') && $('spsSchuelerName').value || '').trim(),
            faecherTitle: String($('spsFaecherName') && $('spsFaecherName').value || '').trim(),
            fachgruppenTitle: String($('spsFachgruppenName') && $('spsFachgruppenName').value || '').trim(),
            argesTitle: String($('spsArgeName') && $('spsArgeName').value || '').trim(),
            klassenTitle: String($('spsKlassenName') && $('spsKlassenName').value || '').trim()
        };
    }

    async function runSync() {
        const logEl = $('spsLog');
        if (logEl) logEl.textContent = '';
        const webUrl = String($('spsSiteUrl') && $('spsSiteUrl').value || '').trim();
        const runOpts = readRunOpts();
        const listOpts = collectOptsFromForm();
        await syncSelectedLists(webUrl, listOpts, undefined, runOpts);
        const skipPerms = $('spsSkipPerms') && $('spsSkipPerms').checked;
        if (!skipPerms && typeof window.ms365SpoStammdatenApplyPermissions === 'function') {
            try {
                await window.ms365SpoStammdatenApplyPermissions(webUrl, listOpts, log);
            } catch (e) {
                log('Berechtigungen: ' + (e && e.message ? e.message : String(e)));
            }
        } else if (skipPerms) {
            log('Berechtigungen übersprungen (Haken gesetzt).');
        }
        toast(runOpts.syncMode ? 'Listen abgeglichen.' : 'Neue Listen erstellt und befüllt.');
    }

    window.ms365SpoStammdatenListen = {
        syncSchuelerList: syncSchuelerList,
        syncFaecherList: syncFaecherList,
        syncFachgruppenList: syncFachgruppenList,
        syncKlassenList: syncKlassenList,
        syncArgeList: syncArgeList,
        syncSelectedLists: syncSelectedLists,
        createSchuelerList: createSchuelerList,
        createFaecherList: createFaecherList,
        createKlassenList: createKlassenList,
        createSelectedLists: createSelectedLists
    };

    const runBtn = $('spsBtnRun');
    if (runBtn) {
        runBtn.addEventListener('click', function () {
            const runOpts = readRunOpts();
            const msg = runOpts.syncMode
                ? 'Ausgewählte Listen auf der Website abgleichen?\n\nVorhandene Listen (gleicher Name) werden aktualisiert; fehlende Zeilen werden ergänzt.' +
                  (runOpts.removeOrphans
                      ? '\nZeilen, die nicht mehr in den Stammdaten sind, werden aus der Liste entfernt.'
                      : '\nAlte Zeilen in SharePoint bleiben erhalten (kein Entfernen).')
                : 'Für jede Auswahl eine NEUE Liste anlegen und befüllen (auch wenn der Name schon existiert)?';
            if (!window.confirm(msg)) return;
            runSync().catch(function (e) {
                log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                toast('Fehler: ' + (e && e.message ? e.message : e));
            });
        });
    }

    const probeBtn = $('spsBtnProbe');
    if (probeBtn) {
        probeBtn.addEventListener('click', function () {
            if ($('spsLog')) $('spsLog').textContent = '';
            const webUrl = String($('spsSiteUrl') && $('spsSiteUrl').value || '').trim();
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
        if (saved && $('spsSiteUrl') && !$('spsSiteUrl').value) $('spsSiteUrl').value = saved;
    } catch {
        /* ignore */
    }
})();
