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

    function readRunOpts(overrides) {
        const syncMode = !$('splAlwaysNew') || !$('splAlwaysNew').checked;
        const base = {
            syncMode: syncMode,
            removeOrphans: !$('splRemoveOrphans') || $('splRemoveOrphans').checked,
            personField: !$('splPersonField') || $('splPersonField').checked,
            clearListFirst: $('splClearFirst') && $('splClearFirst').checked
        };
        if (overrides && typeof overrides === 'object') {
            Object.keys(overrides).forEach(function (k) {
                base[k] = overrides[k];
            });
        }
        return base;
    }

    function personLookupFingerprint(fields, columnName) {
        const name = String(columnName || '').trim();
        const lidKey = name + 'LookupId';
        const rawIds = fields && fields[lidKey];
        if (rawIds != null && rawIds !== '') return String(rawIds);
        const raw = fields && fields[name];
        if (Array.isArray(raw) && raw[0] && raw[0].LookupId != null) return String(raw[0].LookupId);
        if (raw && typeof raw === 'object' && raw.LookupId != null) return String(raw.LookupId);
        return '';
    }

    function fieldsEqualLehrer(prev, next, compareKeys, withPerson) {
        for (let k = 0; k < compareKeys.length; k++) {
            if (String(prev[compareKeys[k]] || '') !== String(next[compareKeys[k]] || '')) return false;
        }
        if (withPerson) {
            if (personLookupFingerprint(prev, 'Lehrkraft') !== personLookupFingerprint(next, 'Lehrkraft')) return false;
        }
        return true;
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

    function splitFirstLastName(displayName) {
        const s = String(displayName || '').trim();
        if (!s) return { vorname: '', nachname: '' };
        const parts = s.split(/\s+/).filter(Boolean);
        if (parts.length <= 1) return { vorname: parts[0] || '', nachname: '' };
        return { vorname: parts[0], nachname: parts.slice(1).join(' ') };
    }

    async function addColumnsLehrer(siteId, listId, token, withPerson) {
        const defs = [
            {
                name: 'Vorname',
                displayName: 'Vorname',
                text: { allowMultipleLines: false, maxLength: 120 }
            },
            {
                name: 'Nachname',
                displayName: 'Nachname',
                text: { allowMultipleLines: false, maxLength: 120 }
            },
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
        if (withPerson) {
            defs.push({
                name: 'Lehrkraft',
                displayName: 'Lehrkraft (Person)',
                personOrGroup: { allowMultipleSelection: false, chooseFromType: 'peopleOnly' }
            });
        }
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
            if (existing.has(defs[i].name)) continue;
            await G.graphJson('POST', base, token, defs[i], 'v1.0');
            await G.sleep(120);
        }
    }

    function teacherFieldsBase(t) {
        const email = String(t.email || '').trim();
        const name = String(t.name || '').trim() || String(t.code || '').trim();
        const code = String(t.code || '').trim();
        const parts = splitFirstLastName(name);
        return {
            Title: name,
            Vorname: parts.vorname,
            Nachname: parts.nachname,
            LehrerCode: code,
            EMail: email,
            UPN: email
        };
    }

    async function teacherFieldsFull(t, resolver, withPerson) {
        const fields = teacherFieldsBase(t);
        if (!withPerson || !resolver) return fields;
        const email = String(t.email || '').trim();
        if (!email) return fields;
        const id = await resolver.lookupIdForEmail(email);
        if (id && typeof G.applySinglePersonLookupField === 'function') {
            G.applySinglePersonLookupField(fields, 'Lehrkraft', id);
        }
        return fields;
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

    async function deleteAllListItems(token, siteId, listId, write) {
        const itemsPath = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/items';
        const items = await fetchAllListItems(token, siteId, listId);
        let deleted = 0;
        for (let i = 0; i < items.length; i++) {
            const itemId = items[i] && items[i].id != null ? String(items[i].id) : '';
            if (!itemId) continue;
            await G.graphJson('DELETE', itemsPath + '/' + encodeURIComponent(itemId), token, undefined, 'v1.0');
            deleted++;
            if (deleted % 25 === 0) write('… ' + deleted + ' Zeilen gelöscht');
            await G.sleep(60);
        }
        return deleted;
    }

    function indexLehrerItemsByCode(items) {
        const byCode = new Map();
        const duplicateItems = [];
        const noCodeItems = [];
        for (let i = 0; i < items.length; i++) {
            const it = items[i];
            const f = (it && it.fields) || {};
            const code = String(f.LehrerCode || '').trim().toLowerCase();
            if (!code) {
                noCodeItems.push(it);
                continue;
            }
            if (!byCode.has(code)) {
                byCode.set(code, it);
                continue;
            }
            duplicateItems.push(it);
        }
        return { byCode: byCode, duplicateItems: duplicateItems, noCodeItems: noCodeItems };
    }

    async function createLehrerList(webUrl, listTitle, logFn, runOpts) {
        const write = typeof logFn === 'function' ? logFn : log;
        const url = String(webUrl || '').trim();
        const title = String(listTitle || '').trim() || 'Lehrerinnen';
        if (!url) throw new Error('Bitte die Adresse der SharePoint-Website eintragen.');

        const teachers = loadTeachers();
        if (!teachers.length) {
            throw new Error('Keine Lehrkräfte in den Stammdaten – zuerst dort pflegen.');
        }

        const ro =
            runOpts && typeof runOpts === 'object'
                ? runOpts
                : { syncMode: true, removeOrphans: true, personField: true, clearListFirst: false };
        const withPerson = ro.personField !== false;
        const clearListFirst = ro.clearListFirst === true;
        if (clearListFirst) ro.syncMode = true;

        write('Lehrkräfte aus lokalem Speicher: ' + teachers.length);
        const token = await ensureToken();
        write('Löse Website auf …');
        const site = await G.resolveSiteFromWebUrl(token, url);
        const siteId = site && site.id ? String(site.id) : '';
        const siteTitle = site && site.displayName ? String(site.displayName) : '';
        if (!siteId) throw new Error('Site-ID fehlt in der Graph-Antwort.');
        write('Site: ' + (siteTitle || siteId));

        let listId = '';
        let listWeb = '';
        if (ro.syncMode) {
            const existing = await findListByDisplayName(token, siteId, title);
            if (existing && existing.id) {
                listId = String(existing.id);
                listWeb = existing.webUrl ? String(existing.webUrl) : '';
                write('Liste „' + title + '" vorhanden – Abgleich.');
            } else {
                write('Erstelle Liste „' + title + '" …');
                const created = await G.graphJson(
                    'POST',
                    G.graphPathSite(siteId) + '/lists',
                    token,
                    { displayName: title, list: { template: 'genericList' } },
                    'v1.0'
                );
                listId = created && created.id ? String(created.id) : '';
                listWeb = created && created.webUrl ? String(created.webUrl) : '';
                if (!listId) throw new Error('Listen-ID fehlt in der Antwort.');
                write('Liste angelegt, ID: ' + listId);
            }
        } else {
            write('Erstelle neue Liste „' + title + '" …');
            const created = await G.graphJson(
                'POST',
                G.graphPathSite(siteId) + '/lists',
                token,
                { displayName: title, list: { template: 'genericList' } },
                'v1.0'
            );
            listId = created && created.id ? String(created.id) : '';
            listWeb = created && created.webUrl ? String(created.webUrl) : '';
            if (!listId) throw new Error('Listen-ID fehlt in der Antwort.');
            write('Liste angelegt, ID: ' + listId);
        }

        write(
            'Prüfe Spalten (Vorname, Nachname, Kürzel, E-Mail, UPN' +
                (withPerson ? ', Personenfeld Lehrkraft' : '') +
                ') …'
        );
        await addColumnsLehrer(siteId, listId, token, withPerson);

        let resolver = null;
        if (withPerson && typeof G.createSitePersonResolver === 'function') {
            write('Personen: User Information List / ensureUser …');
            resolver = await G.createSitePersonResolver(url, token, siteId, { write: write, ensureMissing: true });
        }

        const itemsPath = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/items';
        const compareKeys = ['Title', 'Vorname', 'Nachname', 'LehrerCode', 'EMail', 'UPN'];
        let ok = 0;
        let personMiss = 0;

        if (clearListFirst) {
            write('Leere die Liste – alle vorhandenen Zeilen werden entfernt …');
            const cleared = await deleteAllListItems(token, siteId, listId, write);
            write(cleared + ' Zeilen gelöscht. Spiele Stammdaten neu ein …');
            for (let i = 0; i < teachers.length; i++) {
                const fields = await teacherFieldsFull(teachers[i], resolver, withPerson);
                if (withPerson && String(teachers[i].email || '').trim() && !personLookupFingerprint(fields, 'Lehrkraft')) {
                    personMiss++;
                }
                await G.graphJson('POST', itemsPath, token, { fields: fields }, 'v1.0');
                ok++;
                if (ok % 25 === 0) write('… ' + ok + ' Lehrkräfte eingetragen');
                await G.sleep(80);
            }
            if (withPerson && personMiss) {
                write(
                    'Hinweis: ' +
                        personMiss +
                        ' Lehrkräfte ohne Personen-Verknüpfung (E-Mail fehlt oder Benutzer auf der Site nicht auflösbar).'
                );
            }
            write('Neu eingespielt: ' + ok + ' Lehrkräfte (Liste war zuvor leer).');
        } else if (ro.syncMode) {
            const items = await fetchAllListItems(token, siteId, listId);
            const indexed = indexLehrerItemsByCode(items);
            const byCode = indexed.byCode;
            const desired = new Set();
            let added = 0;
            let updated = 0;
            let deleted = 0;
            for (let i = 0; i < teachers.length; i++) {
                const t = teachers[i];
                const code = String(t.code || '').trim().toLowerCase();
                if (!code) continue;
                desired.add(code);
                const next = await teacherFieldsFull(t, resolver, withPerson);
                if (withPerson && String(t.email || '').trim() && !personLookupFingerprint(next, 'Lehrkraft')) {
                    personMiss++;
                }
                const existing = byCode.get(code);
                if (!existing) {
                    await G.graphJson('POST', itemsPath, token, { fields: next }, 'v1.0');
                    added++;
                    ok++;
                    await G.sleep(80);
                    continue;
                }
                const prev = existing.fields || {};
                const changed = !fieldsEqualLehrer(prev, next, compareKeys, withPerson);
                if (changed) {
                    const itemId = existing.id ? String(existing.id) : '';
                    if (itemId) {
                        await G.graphJson(
                            'PATCH',
                            itemsPath + '/' + encodeURIComponent(itemId) + '/fields',
                            token,
                            next,
                            'v1.0'
                        );
                        updated++;
                        ok++;
                        await G.sleep(80);
                    }
                } else {
                    ok++;
                }
            }
            async function deleteListItem(it) {
                const itemId = it && it.id ? String(it.id) : '';
                if (!itemId) return;
                await G.graphJson('DELETE', itemsPath + '/' + encodeURIComponent(itemId), token, undefined, 'v1.0');
                deleted++;
                await G.sleep(60);
            }

            for (let d = 0; d < indexed.duplicateItems.length; d++) {
                await deleteListItem(indexed.duplicateItems[d]);
            }
            if (indexed.duplicateItems.length) {
                write('Doppelte Zeilen (gleiches Kürzel): ' + indexed.duplicateItems.length + ' entfernt.');
            }

            if (ro.removeOrphans) {
                const keys = Array.from(byCode.keys());
                for (let k = 0; k < keys.length; k++) {
                    const code = keys[k];
                    if (desired.has(code)) continue;
                    await deleteListItem(byCode.get(code));
                }
                for (let n = 0; n < indexed.noCodeItems.length; n++) {
                    await deleteListItem(indexed.noCodeItems[n]);
                }
            }
            if (withPerson && personMiss) {
                write(
                    'Hinweis: ' +
                        personMiss +
                        ' Lehrkräfte ohne Personen-Verknüpfung (E-Mail fehlt oder Benutzer auf der Site nicht auflösbar).'
                );
            }
            write(
                'Abgleich: ' + added + ' neu, ' + updated + ' geändert, ' + ok + ' gesamt' +
                    (ro.removeOrphans ? ', ' + deleted + ' entfernt' : '')
            );
        } else {
            for (let i = 0; i < teachers.length; i++) {
                const fields = await teacherFieldsFull(teachers[i], resolver, withPerson);
                await G.graphJson('POST', itemsPath, token, { fields: fields }, 'v1.0');
                ok++;
                if (ok % 10 === 0) write('… ' + ok + ' Zeilen geschrieben');
                await G.sleep(80);
            }
            write('Fertig: ' + ok + ' Lehrkräfte als Listenelemente.');
        }

        if (listWeb) write('Liste im Browser: ' + listWeb);
        if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
            window.ms365ActionLog.append({
                tool: 'sharepoint',
                action: clearListFirst
                    ? 'replace-lehrer-list'
                    : ro.syncMode
                      ? 'sync-lehrer-list'
                      : 'create-lehrer-list',
                target: url,
                summary: 'Lehrerliste „' + title + '“ – ' + ok + ' Einträge'
            });
        }
        return { listId: listId, webUrl: listWeb, count: ok };
    }

    async function runCreate(overrides) {
        const logEl = $('splLog');
        if (logEl) logEl.textContent = '';
        const webUrl = String($('splSiteUrl') && $('splSiteUrl').value || '').trim();
        const listTitle = String($('splListName') && $('splListName').value || '').trim() || 'Lehrerinnen';
        const ro = readRunOpts(overrides);
        const created = await createLehrerList(webUrl, listTitle, undefined, ro);
        if (ro.clearListFirst) toast('Lehrerliste geleert und neu eingespielt.');
        else toast(ro.syncMode ? 'Lehrerliste abgeglichen.' : 'Lehrerliste erstellt und befüllt.');
        return created;
    }

    window.ms365SpoLehrerListe = { createList: createLehrerList, syncList: createLehrerList };

    const runBtn = $('splBtnRun');
    if (runBtn) {
        runBtn.addEventListener('click', function () {
            const ro = readRunOpts();
            let msg;
            if (ro.clearListFirst) {
                msg =
                    'Alle Zeilen in der Liste „' +
                    (String($('splListName') && $('splListName').value || '').trim() || 'Lehrerinnen') +
                    '“ löschen und danach nur die Lehrkräfte aus den Stammdaten neu anlegen?\n\nDoppelte Einträge verschwinden; kurzzeitig ist die Liste leer.';
            } else if (ro.syncMode) {
                msg = 'Lehrerliste auf der Website abgleichen (vorhandene Liste nach Name, sonst neu anlegen)?';
            } else {
                msg = 'Neue Lehrerliste anlegen und alle Lehrkräfte eintragen?';
            }
            if (!window.confirm(msg)) return;
            runCreate().catch(function (e) {
                log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                toast('Fehler: ' + (e && e.message ? e.message : e));
            });
        });
    }

    const replaceBtn = $('splBtnReplace');
    if (replaceBtn) {
        replaceBtn.addEventListener('click', function () {
            const listTitle =
                String($('splListName') && $('splListName').value || '').trim() || 'Lehrerinnen';
            const msg =
                'Liste „' +
                listTitle +
                '“ komplett leeren und mit den Stammdaten neu füllen?\n\nAlle bisherigen Zeilen (auch Duplikate) werden gelöscht.';
            if (!window.confirm(msg)) return;
            runCreate({ clearListFirst: true, syncMode: true }).catch(function (e) {
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
