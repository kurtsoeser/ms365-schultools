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
            klassenPersonen: !$('spsKlassenPersonen') || $('spsKlassenPersonen').checked,
            lehrerPersonen: !$('spsLehrerPersonen') || $('spsLehrerPersonen').checked
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
        { name: 'KVEmail', displayName: 'KV E-Mail', text: { allowMultipleLines: false, maxLength: 255 } },
        {
            name: 'M365GroupId',
            displayName: 'M365-Gruppen-ID',
            text: { allowMultipleLines: false, maxLength: 80 }
        },
        {
            name: 'M365GroupMail',
            displayName: 'M365-Gruppen-E-Mail',
            text: { allowMultipleLines: false, maxLength: 255 }
        },
        {
            name: 'M365GroupName',
            displayName: 'M365-Gruppenname',
            text: { allowMultipleLines: false, maxLength: 200 }
        },
        {
            name: 'M365Gruppe',
            displayName: 'M365-Gruppe (Team)',
            personOrGroup: { allowMultipleSelection: false, chooseFromType: 'peopleAndGroups' }
        }
    ];

    const COL_KLASSE_PERSON = COL_KLASSE.concat([
        {
            name: 'Schuelerinnen',
            displayName: 'Schülerinnen',
            personOrGroup: { allowMultipleSelection: true, chooseFromType: 'peopleOnly' }
        },
        {
            name: 'Lehrkraefte',
            displayName: 'Lehrkräfte (unterrichtend)',
            personOrGroup: { allowMultipleSelection: true, chooseFromType: 'peopleOnly' }
        },
        {
            name: 'Klassenvorstand',
            displayName: 'KV (Person)',
            personOrGroup: { allowMultipleSelection: false, chooseFromType: 'peopleOnly' }
        }
    ]);

    function loadUnterrichtsbelegungRows() {
        try {
            const api = window.ms365AppDataV2;
            if (api && typeof api.getUnterrichtsbelegung === 'function') {
                const snap = api.getUnterrichtsbelegung();
                if (snap && Array.isArray(snap.rows)) return snap.rows;
            }
        } catch {
            /* ignore */
        }
        try {
            const s = loadSettings();
            if (s && s.unterrichtsbelegung && Array.isArray(s.unterrichtsbelegung.rows)) {
                return s.unterrichtsbelegung.rows;
            }
        } catch {
            /* ignore */
        }
        return null;
    }

    function normTeachingEmail(v) {
        return String(v || '')
            .trim()
            .toLowerCase();
    }

    function normTeachingCode(v) {
        return String(v || '')
            .trim()
            .toUpperCase();
    }

    function classKeyForTeaching(rowOrCode) {
        if (rowOrCode && typeof rowOrCode === 'object') {
            return normTeachingCode(rowOrCode.code) || String(rowOrCode.name || '').trim().toUpperCase();
        }
        return normTeachingCode(rowOrCode) || String(rowOrCode || '').trim().toUpperCase();
    }

    /**
     * Fallback, falls die ES-Modul-Bridge noch nicht geladen ist (z. B. Intranet-Einbettung).
     * @param {Array<object>|null} belegungRows
     * @param {Array<object>} teachers
     * @returns {Map<string, string[]>|null}
     */
    function buildTeachingEmailsByKlasseFallback(belegungRows, teachers) {
        if (!Array.isArray(belegungRows)) return null;
        /** @type {Map<string, string>} */
        const emailByCode = new Map();
        (Array.isArray(teachers) ? teachers : []).forEach(function (t) {
            const code = normTeachingCode(t && t.code);
            const email = normTeachingEmail(t && t.email);
            if (!code || !email || email.indexOf('@') === -1) return;
            if (!emailByCode.has(code)) emailByCode.set(code, email);
        });
        /** @type {Map<string, Map<string, string>>} */
        const byClass = new Map();
        belegungRows.forEach(function (r) {
            const klasse = classKeyForTeaching(r && r.klasse) || String(r && r.klasse || '').trim().toUpperCase();
            if (!klasse) return;
            let email = normTeachingEmail(r && r.lehrerEmail);
            const code = normTeachingCode(r && r.lehrerCode);
            if ((!email || email.indexOf('@') === -1) && code) {
                email = emailByCode.get(code) || '';
            }
            if (!email || email.indexOf('@') === -1) return;
            if (!byClass.has(klasse)) byClass.set(klasse, new Map());
            byClass.get(klasse).set(email, email);
        });
        /** @type {Map<string, string[]>} */
        const out = new Map();
        byClass.forEach(function (map, klasse) {
            out.set(klasse, Array.from(map.keys()).sort());
        });
        return out;
    }

    /**
     * @param {Array<object>|null} belegungRows
     * @param {Array<object>} teachers
     * @returns {Map<string, string[]>|null}
     */
    function buildTeachingEmailsByKlasse(belegungRows, teachers) {
        const logic = window.ms365KlassenlehrerLogic;
        if (logic && typeof logic.teachingTeacherEmailsByKlasse === 'function' && Array.isArray(belegungRows)) {
            return logic.teachingTeacherEmailsByKlasse(belegungRows, teachers);
        }
        return buildTeachingEmailsByKlasseFallback(belegungRows, teachers);
    }

    /**
     * @param {Map<string, string[]>|null} emailsByKlasse
     * @param {object} classRow
     * @returns {string[]}
     */
    function emailsForClassTeaching(emailsByKlasse, classRow) {
        const logic = window.ms365KlassenlehrerLogic;
        if (logic && typeof logic.teachingEmailsForClassRow === 'function') {
            return logic.teachingEmailsForClassRow(emailsByKlasse, classRow) || [];
        }
        if (!emailsByKlasse || typeof emailsByKlasse.get !== 'function') return [];
        const key = classKeyForTeaching(classRow);
        if (key && emailsByKlasse.has(key)) return emailsByKlasse.get(key) || [];
        const nameKey = normTeachingCode(classRow && classRow.name);
        if (nameKey && emailsByKlasse.has(nameKey)) return emailsByKlasse.get(nameKey) || [];
        return [];
    }

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
        const teachers = Array.isArray(s.teachers) ? s.teachers : [];
        const belegungRows = withPersons ? loadUnterrichtsbelegungRows() : null;
        const teachingEmailsByKlasse = withPersons
            ? buildTeachingEmailsByKlasse(belegungRows, teachers)
            : null;
        let teachingEmailCount = 0;
        if (teachingEmailsByKlasse) {
            teachingEmailsByKlasse.forEach(function (emails) {
                teachingEmailCount += Array.isArray(emails) ? emails.length : 0;
            });
        }

        write(
            'Klassen: ' +
                classes.length +
                (withPersons
                    ? ' (Schülerinnen-Personenfeld, ' +
                      students.length +
                      ' Schüler mit E-Mail' +
                      (teachingEmailsByKlasse
                          ? '; Lehrkräfte aus Belegung, ' + teachingEmailCount + ' Zuordnungen'
                          : belegungRows === null
                            ? '; keine Unterrichtsbelegung – Lehrkräfte-Personenfeld bleibt unverändert'
                            : '; Lehrkräfte-Logik nicht geladen') +
                      ')'
                    : '')
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

        const compareKeys = [
            'Title',
            'KlassenCode',
            'Abschlussjahr',
            'KVName',
            'KVEmail',
            'M365GroupId',
            'M365GroupMail',
            'M365GroupName'
        ];

        function classTeamLink(klasseCode) {
            const code = String(klasseCode || '').trim();
            if (!code) return null;
            const want = code.toUpperCase();
            const api = window.ms365AppDataV2;
            if (!api || typeof api.getContainer !== 'function') return null;
            const container = api.getContainer();
            const teams =
                typeof api.normalizeCoreClassTeams === 'function'
                    ? api.normalizeCoreClassTeams((container.core && container.core.classTeams) || [])
                    : (container.core && container.core.classTeams) || [];
            for (let i = 0; i < teams.length; i++) {
                const t = teams[i];
                if (!t) continue;
                const cc = String(t.classCode || '')
                    .trim()
                    .toUpperCase();
                const dn = String(t.displayName || '').trim();
                if ((dn && code === dn) || (cc && want === cc)) {
                    return {
                        groupId: String(t.graphGroupId || '').trim(),
                        mail: String(t.mail || '').trim().toLowerCase(),
                        name: dn || code
                    };
                }
            }
            const setup = typeof api.getSetup === 'function' ? api.getSetup() : null;
            const map = (setup && setup.classGroupMatchByKey) || {};
            const entry = map[want] || map[code] || null;
            if (!entry || typeof entry !== 'object') return null;
            return {
                groupId: String(entry.groupId || entry.graphGroupId || '').trim(),
                mail: String(entry.mail || entry.email || '')
                    .trim()
                    .toLowerCase(),
                name: String(entry.displayName || entry.name || code).trim()
            };
        }

        async function ensureGroupLookupId(groupId) {
            const gid = String(groupId || '').trim();
            if (!gid || typeof G.spoEnsureUser !== 'function') return 0;
            const logon = 'c:0o.c|federateddirectoryclaimprovider|' + gid;
            try {
                let host = '';
                try {
                    host = new URL(url).hostname;
                } catch {
                    host = '';
                }
                if (!host) return 0;
                const spoToken = await G.getGraphToken(['https://' + host + '/AllSites.Write']);
                const digest = await G.getSpoRequestDigest(url, spoToken);
                const principal = await G.spoEnsureUser(url, spoToken, digest, logon);
                return principal && principal.id ? Number(principal.id) : 0;
            } catch (e) {
                write(
                    'M365-Gruppe ' +
                        gid.slice(0, 8) +
                        '…: ' +
                        (e && e.message ? e.message : e)
                );
                return 0;
            }
        }

        async function fieldsForClass(c) {
            const code = String(c.code || '').trim();
            const name = String(c.name || '').trim() || code;
            const link = classTeamLink(code);
            let groupId = String((c && c.m365GroupId) || (link && link.groupId) || '').trim();
            let groupMail = String((c && c.m365GroupMail) || (link && link.mail) || '')
                .trim()
                .toLowerCase();
            let groupName = String((c && c.m365GroupName) || (link && link.name) || '').trim();
            if (groupId && !groupMail) {
                try {
                    const g = await G.graphJson(
                        'GET',
                        '/groups/' + encodeURIComponent(groupId) + '?$select=mail,displayName',
                        token,
                        undefined,
                        'v1.0'
                    );
                    groupMail = String((g && g.mail) || '')
                        .trim()
                        .toLowerCase();
                    if (!groupName) groupName = String((g && g.displayName) || '').trim();
                } catch {
                    /* ignore */
                }
            }
            const fields = {
                Title: name,
                KlassenCode: code,
                Abschlussjahr: String(c.year || '').trim(),
                KVName: String(c.headName || '').trim(),
                KVEmail: String(c.headEmail || '').trim()
            };
            if (groupId) fields.M365GroupId = groupId;
            if (groupMail) fields.M365GroupMail = groupMail;
            if (groupName) fields.M365GroupName = groupName;

            if (groupId) {
                const gLookup = await ensureGroupLookupId(groupId);
                if (gLookup) G.applySinglePersonLookupField(fields, 'M365Gruppe', gLookup);
            }

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

            if (teachingEmailsByKlasse) {
                const teacherEmails = emailsForClassTeaching(teachingEmailsByKlasse, c);
                const teacherIds = [];
                let teacherMiss = 0;
                for (let ti = 0; ti < teacherEmails.length; ti++) {
                    const tem = teacherEmails[ti];
                    const tid = await resolver.lookupIdForEmail(tem);
                    if (tid) teacherIds.push(tid);
                    else teacherMiss++;
                }
                if (teacherMiss) {
                    write(
                        'Klasse ' +
                            (code || name) +
                            ': ' +
                            teacherMiss +
                            ' Lehrkräfte mit E-Mail konnten auf der Site nicht verknüpft werden (ensureUser/Rechte?).'
                    );
                }
                G.applyMultiPersonLookupFields(fields, 'Lehrkraefte', teacherIds);
            }

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
            if (
                personLookupFingerprint(prev, 'M365Gruppe') !==
                personLookupFingerprint(next, 'M365Gruppe')
            )
                return false;
            if (!withPersons) return true;
            if (
                personLookupFingerprint(prev, 'Schuelerinnen') !==
                personLookupFingerprint(next, 'Schuelerinnen')
            )
                return false;
            if (
                teachingEmailsByKlasse &&
                personLookupFingerprint(prev, 'Lehrkraefte') !==
                    personLookupFingerprint(next, 'Lehrkraefte')
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
                summary:
                    'Klassenliste „' +
                    title +
                    '“' +
                    (withPersons
                        ? ' inkl. Personenfelder (Schülerinnen' +
                          (teachingEmailsByKlasse ? ', Lehrkräfte' : '') +
                          ', KV)'
                        : '')
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

    function loadBelegungRowsForSpo(teachers) {
        const rows = loadUnterrichtsbelegungRows();
        if (!Array.isArray(rows)) return [];
        const logic = window.ms365KlassenlehrerLogic;
        if (logic && typeof logic.belegungRowsForSpoSync === 'function') {
            return logic.belegungRowsForSpoSync(rows, teachers) || [];
        }
        // Minimal-Fallback ohne Modul-Bridge
        /** @type {Map<string, object>} */
        const byKey = new Map();
        rows.forEach(function (r) {
            const klasse = String(r && r.klasse || '')
                .trim()
                .toUpperCase();
            if (!klasse) return;
            const email = String(r && r.lehrerEmail || '')
                .trim()
                .toLowerCase();
            const code = String(r && r.lehrerCode || '')
                .trim()
                .toUpperCase();
            const fach = String(r && r.fach || '').trim();
            const teacherPart = email || code;
            if (!teacherPart) return;
            const syncKey = klasse.toLowerCase() + '|' + teacherPart + '|' + fach.toLowerCase();
            if (byKey.has(syncKey)) return;
            byKey.set(syncKey, {
                syncKey: syncKey,
                title: [klasse, code || email, fach].filter(Boolean).join(' · '),
                klasse: klasse,
                lehrerCode: code,
                lehrerEmail: email,
                lehrerName: String(r && r.lehrerName || '').trim(),
                fach: fach
            });
        });
        return Array.from(byKey.values());
    }

    /**
     * Index: KlassenCode/Title → Listenelement-ID (für Lookup).
     */
    function buildKlassenLookupIndex(items) {
        /** @type {Map<string, string>} */
        const byCode = new Map();
        (Array.isArray(items) ? items : []).forEach(function (it) {
            const id = it && it.id != null ? String(it.id) : '';
            if (!id) return;
            const f = (it && it.fields) || {};
            const code = String(f.KlassenCode || '')
                .trim()
                .toLowerCase();
            const title = String(f.Title || '')
                .trim()
                .toLowerCase();
            if (code && !byCode.has(code)) byCode.set(code, id);
            if (title && !byCode.has(title)) byCode.set(title, id);
        });
        return byCode;
    }

    /**
     * Index: LehrerCode / E-Mail → Listenelement-ID.
     */
    function buildLehrerLookupIndex(items) {
        /** @type {{ byCode: Map<string, string>, byEmail: Map<string, string> }} */
        const out = { byCode: new Map(), byEmail: new Map() };
        (Array.isArray(items) ? items : []).forEach(function (it) {
            const id = it && it.id != null ? String(it.id) : '';
            if (!id) return;
            const f = (it && it.fields) || {};
            const code = String(f.LehrerCode || '')
                .trim()
                .toLowerCase();
            const email = String(f.EMail || f.Email || f.UPN || '')
                .trim()
                .toLowerCase();
            if (code && !out.byCode.has(code)) out.byCode.set(code, id);
            if (email && email.indexOf('@') !== -1 && !out.byEmail.has(email)) out.byEmail.set(email, id);
        });
        return out;
    }

    function buildFaecherLookupIndex(items) {
        /** @type {Map<string, string>} */
        const byCode = new Map();
        (Array.isArray(items) ? items : []).forEach(function (it) {
            const id = it && it.id != null ? String(it.id) : '';
            if (!id) return;
            const f = (it && it.fields) || {};
            const code = String(f.FachCode || '')
                .trim()
                .toLowerCase();
            const title = String(f.Title || '')
                .trim()
                .toLowerCase();
            if (code && !byCode.has(code)) byCode.set(code, id);
            if (title && !byCode.has(title)) byCode.set(title, id);
        });
        return byCode;
    }

    /**
     * Eigene Liste Klasse × Lehrkraft × Fach mit echten SharePoint-Lookups.
     *
     * @param {string} webUrl
     * @param {string} listTitle
     * @param {function} [logFn]
     * @param {object} [runOpts]
     * @param {{ klassenTitle?: string, lehrerTitle?: string, faecherTitle?: string }} [deps]
     */
    async function syncLehrerJeKlasseList(webUrl, listTitle, logFn, runOpts, deps) {
        const write = typeof logFn === 'function' ? logFn : log;
        const url = String(webUrl || '').trim();
        const title = String(listTitle || '').trim() || 'Lehrer je Klasse';
        if (!url) throw new Error('Bitte die Adresse der SharePoint-Website eintragen.');

        const ro = runOpts && typeof runOpts === 'object' ? runOpts : readRunOpts();
        const d = deps && typeof deps === 'object' ? deps : {};
        const klassenTitle = String(d.klassenTitle || 'Klassen').trim() || 'Klassen';
        const lehrerTitle = String(d.lehrerTitle || 'Lehrerinnen').trim() || 'Lehrerinnen';
        const faecherTitle = String(d.faecherTitle || 'Fächer').trim() || 'Fächer';

        const s = loadSettings();
        const teachers = Array.isArray(s.teachers) ? s.teachers : [];
        const rows = loadBelegungRowsForSpo(teachers);
        if (!rows.length) {
            throw new Error(
                'Keine Unterrichtsbelegung („Lehrer je Klasse“) in den Stammdaten – zuerst dort pflegen.'
            );
        }

        write('Belegung: ' + rows.length + ' Zeilen (Klasse × Lehrkraft × Fach)');
        const token = await ensureToken();
        const siteId = await resolveSite(url, token, write);
        const syncMode = ro.syncMode !== false;
        const removeOrphans = ro.removeOrphans !== false;

        const klassenList = await findListByDisplayName(token, siteId, klassenTitle);
        if (!klassenList || !klassenList.id) {
            throw new Error(
                'Lookup-Ziel „' +
                    klassenTitle +
                    '“ fehlt – bitte Klassenliste zuerst anlegen/abgleichen.'
            );
        }
        const lehrerList = await findListByDisplayName(token, siteId, lehrerTitle);
        if (!lehrerList || !lehrerList.id) {
            throw new Error(
                'Lookup-Ziel „' +
                    lehrerTitle +
                    '“ fehlt – bitte Lehrerinnenliste zuerst anlegen/abgleichen.'
            );
        }

        write('Lade Lookup-Ziele …');
        const klassenItems = await fetchAllListItems(token, siteId, String(klassenList.id));
        const lehrerItems = await fetchAllListItems(token, siteId, String(lehrerList.id));
        const klassenIdx = buildKlassenLookupIndex(klassenItems);
        const lehrerIdx = buildLehrerLookupIndex(lehrerItems);
        write(
            'Lookup-Index: ' +
                klassenIdx.size +
                ' Klassen, ' +
                (lehrerIdx.byCode.size + lehrerIdx.byEmail.size) +
                ' Lehrkraft-Schlüssel'
        );

        let faecherIdx = null;
        const faecherList = await findListByDisplayName(token, siteId, faecherTitle);
        if (faecherList && faecherList.id) {
            const faecherItems = await fetchAllListItems(token, siteId, String(faecherList.id));
            faecherIdx = buildFaecherLookupIndex(faecherItems);
            write('Optional: Fächer-Lookup (' + faecherIdx.size + ' Einträge)');
        }

        const listMeta = await ensureOrCreateList(siteId, title, token, write, syncMode);
        const textCols = [
            {
                name: 'KlassenCode',
                displayName: 'Klassenkürzel',
                text: { allowMultipleLines: false, maxLength: 40 }
            },
            {
                name: 'LehrerCode',
                displayName: 'Lehrkraft-Kürzel',
                text: { allowMultipleLines: false, maxLength: 40 }
            },
            {
                name: 'LehrerEmail',
                displayName: 'Lehrkraft E-Mail',
                text: { allowMultipleLines: false, maxLength: 255 }
            },
            {
                name: 'Fach',
                displayName: 'Fach',
                text: { allowMultipleLines: false, maxLength: 80 }
            },
            {
                name: 'SyncKey',
                displayName: 'Sync-Schlüssel',
                text: { allowMultipleLines: false, maxLength: 200 }
            }
        ];
        const lookupCols = [
            {
                name: 'Klasse',
                displayName: 'Klasse (Lookup)',
                lookup: {
                    listId: String(klassenList.id),
                    columnName: 'Title',
                    allowMultipleValues: false
                }
            },
            {
                name: 'Lehrkraft',
                displayName: 'Lehrkraft (Lookup)',
                lookup: {
                    listId: String(lehrerList.id),
                    columnName: 'Title',
                    allowMultipleValues: false
                }
            }
        ];
        if (faecherList && faecherList.id) {
            lookupCols.push({
                name: 'Unterrichtsfach',
                displayName: 'Fach (Lookup)',
                lookup: {
                    listId: String(faecherList.id),
                    columnName: 'Title',
                    allowMultipleValues: false
                }
            });
        }
        write('Prüfe Spalten (Text + Lookups) …');
        await addMissingColumns(siteId, listMeta.listId, token, textCols.concat(lookupCols));

        function resolveLehrerItemId(row) {
            const email = String(row.lehrerEmail || '')
                .trim()
                .toLowerCase();
            const code = String(row.lehrerCode || '')
                .trim()
                .toLowerCase();
            if (email && lehrerIdx.byEmail.has(email)) return lehrerIdx.byEmail.get(email);
            if (code && lehrerIdx.byCode.has(code)) return lehrerIdx.byCode.get(code);
            return '';
        }

        function fieldsFromBelegung(row) {
            const klasse = String(row.klasse || '')
                .trim()
                .toUpperCase();
            const fields = {
                Title: String(row.title || '').trim() || klasse,
                KlassenCode: klasse,
                LehrerCode: String(row.lehrerCode || '').trim(),
                LehrerEmail: String(row.lehrerEmail || '').trim().toLowerCase(),
                Fach: String(row.fach || '').trim(),
                SyncKey: String(row.syncKey || '').trim()
            };
            const kId = klassenIdx.get(klasse.toLowerCase()) || '';
            if (kId) fields.KlasseLookupId = Number(kId) || kId;
            const lId = resolveLehrerItemId(row);
            if (lId) fields.LehrkraftLookupId = Number(lId) || lId;
            if (faecherIdx && fields.Fach) {
                const fId = faecherIdx.get(fields.Fach.toLowerCase()) || '';
                if (fId) fields.UnterrichtsfachLookupId = Number(fId) || fId;
            }
            return fields;
        }

        function fieldsEqualBelegung(prev, next) {
            const keys = ['Title', 'KlassenCode', 'LehrerCode', 'LehrerEmail', 'Fach', 'SyncKey'];
            for (let i = 0; i < keys.length; i++) {
                const k = keys[i];
                if (String(prev[k] || '') !== String(next[k] || '')) return false;
            }
            if (personLookupFingerprint(prev, 'Klasse') !== personLookupFingerprint(next, 'Klasse')) {
                return false;
            }
            if (
                personLookupFingerprint(prev, 'Lehrkraft') !== personLookupFingerprint(next, 'Lehrkraft')
            ) {
                return false;
            }
            if (
                faecherIdx &&
                personLookupFingerprint(prev, 'Unterrichtsfach') !==
                    personLookupFingerprint(next, 'Unterrichtsfach')
            ) {
                return false;
            }
            return true;
        }

        let missKlasse = 0;
        let missLehrer = 0;
        rows.forEach(function (row) {
            const klasse = String(row.klasse || '')
                .trim()
                .toLowerCase();
            if (!klassenIdx.has(klasse)) missKlasse++;
            if (!resolveLehrerItemId(row)) missLehrer++;
        });
        if (missKlasse) {
            write(
                'Hinweis: ' +
                    missKlasse +
                    ' Zeilen ohne Klassen-Lookup (Kürzel nicht in „' +
                    klassenTitle +
                    '“).'
            );
        }
        if (missLehrer) {
            write(
                'Hinweis: ' +
                    missLehrer +
                    ' Zeilen ohne Lehrkraft-Lookup (nicht in „' +
                    lehrerTitle +
                    '“).'
            );
        }

        if (syncMode) {
            await syncRowsToList({
                siteId: siteId,
                listId: listMeta.listId,
                token: token,
                rows: rows,
                keyFromRow: function (row) {
                    return String(row.syncKey || '').trim().toLowerCase();
                },
                keyFromItem: function (f) {
                    const sk = String(f.SyncKey || '').trim().toLowerCase();
                    if (sk) return sk;
                    const k = String(f.KlassenCode || '').trim().toLowerCase();
                    const em = String(f.LehrerEmail || '').trim().toLowerCase();
                    const code = String(f.LehrerCode || '').trim().toLowerCase();
                    const fach = String(f.Fach || '').trim().toLowerCase();
                    return k + '|' + (em || code) + '|' + fach;
                },
                fieldsFromRow: fieldsFromBelegung,
                fieldsEqual: fieldsEqualBelegung,
                compareKeys: ['Title', 'KlassenCode', 'LehrerCode', 'LehrerEmail', 'Fach', 'SyncKey'],
                removeOrphans: removeOrphans,
                write: write,
                label: 'Lehrer je Klasse'
            });
        } else {
            await postAllRows(
                siteId,
                listMeta.listId,
                token,
                rows,
                fieldsFromBelegung,
                write,
                'Lehrer-je-Klasse'
            );
        }

        if (listMeta.webUrl) write('Liste im Browser: ' + listMeta.webUrl);
        if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
            window.ms365ActionLog.append({
                tool: 'sharepoint',
                action: 'sync-lehrer-je-klasse-list',
                target: url,
                summary: 'Lehrer je Klasse „' + title + '“ (' + rows.length + ' Zeilen, Lookups)'
            });
        }
        return { listId: listMeta.listId, webUrl: listMeta.webUrl, count: rows.length };
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
        } else if (o.lehrerJeKlasse) {
            const tokenProbe = await ensureToken();
            const siteProbe = await G.resolveSiteFromWebUrl(tokenProbe, String(webUrl || '').trim());
            const siteIdProbe = siteProbe && siteProbe.id ? String(siteProbe.id) : '';
            const kTitle = o.klassenTitle || 'Klassen';
            const existingK = siteIdProbe
                ? await findListByDisplayName(tokenProbe, siteIdProbe, kTitle)
                : null;
            if (!(existingK && existingK.id)) {
                write('—— Klassen (Voraussetzung für Lookups) ——');
                results.klassen = await syncKlassenList(webUrl, kTitle, write, ro);
            }
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
        } else if (o.lehrerJeKlasse) {
            const tokenProbe = await ensureToken();
            const siteProbe = await G.resolveSiteFromWebUrl(tokenProbe, String(webUrl || '').trim());
            const siteIdProbe = siteProbe && siteProbe.id ? String(siteProbe.id) : '';
            const lTitle = o.lehrerTitle || 'Lehrerinnen';
            const existingL = siteIdProbe
                ? await findListByDisplayName(tokenProbe, siteIdProbe, lTitle)
                : null;
            if (!(existingL && existingL.id)) {
                write('—— Lehrerinnen (Voraussetzung für Lookups) ——');
                const lehrerApi = window.ms365SpoLehrerListe;
                if (!lehrerApi || typeof lehrerApi.createList !== 'function') {
                    throw new Error('Lehrerlisten-Sync nicht geladen (sharepoint-liste-lehrer.js).');
                }
                results.lehrer = await lehrerApi.createList(webUrl, lTitle, write, {
                    syncMode: ro.syncMode,
                    removeOrphans: ro.removeOrphans,
                    personField: ro.lehrerPersonen !== false,
                    clearListFirst: false
                });
            }
        }
        if (o.lehrerJeKlasse) {
            write('—— Lehrer je Klasse (Lookups) ——');
            results.lehrerJeKlasse = await syncLehrerJeKlasseList(
                webUrl,
                o.lehrerJeKlasseTitle || 'Lehrer je Klasse',
                write,
                ro,
                {
                    klassenTitle: o.klassenTitle || 'Klassen',
                    lehrerTitle: o.lehrerTitle || 'Lehrerinnen',
                    faecherTitle: o.faecherTitle || 'Fächer'
                }
            );
        }
        if (
            !o.schueler &&
            !o.faecher &&
            !o.fachgruppen &&
            !o.arges &&
            !o.klassen &&
            !o.lehrer &&
            !o.lehrerJeKlasse
        ) {
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
            lehrer: $('spsWantLehrer') && $('spsWantLehrer').checked,
            lehrerJeKlasse: $('spsWantLehrerJeKlasse') && $('spsWantLehrerJeKlasse').checked,
            schuelerTitle: String($('spsSchuelerName') && $('spsSchuelerName').value || '').trim(),
            faecherTitle: String($('spsFaecherName') && $('spsFaecherName').value || '').trim(),
            fachgruppenTitle: String($('spsFachgruppenName') && $('spsFachgruppenName').value || '').trim(),
            argesTitle: String($('spsArgeName') && $('spsArgeName').value || '').trim(),
            klassenTitle: String($('spsKlassenName') && $('spsKlassenName').value || '').trim(),
            lehrerTitle: String($('spsLehrerName') && $('spsLehrerName').value || '').trim(),
            lehrerJeKlasseTitle: String(
                $('spsLehrerJeKlasseName') && $('spsLehrerJeKlasseName').value || ''
            ).trim()
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
        syncLehrerJeKlasseList: syncLehrerJeKlasseList,
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
