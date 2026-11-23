/**
 * tenant-settings-ui-classes-panel.js
 * @param {Record<string, any>} ctx
 */
import { collapseActionCell, editTableRow, presentText, presentValue } from './tenant-workbench-chrome.js';

export function installTenantClassesPanel(ctx) {
        ctx.renderClassesTableFromTextarea = function renderClassesTableFromTextarea() {
            if (!ctx.classesTbody) return;
            ctx.updateTableSortIndicators(ctx.classesTable, ctx.classesSortState);
            const rows = ctx.getClassesFromTextarea();
            ctx.classesTbody.replaceChildren();

            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 7;
                td.style.color = 'var(--muted)';
                td.textContent = 'Noch keine Einträge – Importieren oder „+ Klasse“.';
                tr.appendChild(td);
                ctx.classesTbody.appendChild(tr);
                ctx.classesTableEditFocus = null;
                return;
            }

            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');
                tr.dataset.classRow = String(idx);

                const tdCode = document.createElement('td');
                ctx.bindClassEditableCell(tdCode, idx, 'code', (cell, r) => {
                    const code = presentValue(r && r.code);
                    cell.innerHTML = code ? `<code>${code}</code>` : '–';
                });

                const tdYear = document.createElement('td');
                ctx.bindClassEditableCell(tdYear, idx, 'year', (cell, r) => {
                    cell.textContent = presentText(r && r.year);
                });

                const tdName = document.createElement('td');
                ctx.bindClassEditableCell(tdName, idx, 'name', (cell, r) => {
                    cell.textContent = presentText(r && r.name);
                });

                const tdHead = document.createElement('td');
                ctx.bindClassEditableCell(tdHead, idx, 'headName', (cell, r) => {
                    cell.textContent = presentText(r && r.headName);
                });

                const tdEmail = document.createElement('td');
                ctx.bindClassEditableCell(tdEmail, idx, 'headEmail', (cell, r) => {
                    cell.textContent = presentText(r && r.headEmail);
                });

                const tdMs = ctx.buildClassGroupMatchCell(tr, row);

                const tdAction = document.createElement('td');
                tdAction.className = 'action-cell';
                tdAction.style.whiteSpace = 'nowrap';
                tdAction.style.display = 'flex';
                tdAction.style.gap = '6px';
                tdAction.style.alignItems = 'center';
                const btnCheck = document.createElement('button');
                btnCheck.type = 'button';
                btnCheck.className = 'mini-btn';
                btnCheck.style.background = 'var(--brand1)';
                btnCheck.title = 'Klassengruppe in Microsoft 365 prüfen';
                btnCheck.innerHTML = '<i class="bi bi-microsoft" aria-hidden="true"></i>';
                btnCheck.addEventListener('click', async () => {
                    btnCheck.disabled = true;
                    try {
                        const res = await ctx.verifyClassGroupForRow(row);
                        ctx.setSummary(
                            (row.code || row.name || 'Klasse') + ': ' + (res && res.found ? 'Klassengruppe gefunden' : 'keine Klassengruppe gefunden'),
                            res && res.found ? 'ok' : 'warn'
                        );
                    } catch (e) {
                        ctx.setSummary('Klassenabgleich: ' + (e && e.message ? e.message : String(e)), 'warn');
                    } finally {
                        ctx.renderClassesTableFromTextarea();
                    }
                });
                tdAction.appendChild(btnCheck);
                const matchKey = ctx.classMatchKey(row);
                const linked = matchKey ? ctx.getClassGroupMatchByKey(matchKey) : null;
                const btnCreateClassGroup = document.createElement('button');
                btnCreateClassGroup.type = 'button';
                btnCreateClassGroup.className = 'mini-btn';
                btnCreateClassGroup.style.background = '#198754';
                btnCreateClassGroup.title = linked && linked.groupId
                    ? 'Bereits mit einer Microsoft-365-Gruppe verknüpft'
                    : 'Klassengruppe in Microsoft 365 anlegen (Alias nach Stammdaten-Schema)';
                btnCreateClassGroup.innerHTML = '<i class="bi bi-plus-lg" aria-hidden="true"></i>';
                if (linked && linked.groupId) {
                    btnCreateClassGroup.disabled = true;
                    btnCreateClassGroup.style.opacity = '0.45';
                }
                btnCreateClassGroup.addEventListener('click', async () => {
                    const all = ctx.getClassesFromTextarea();
                    const liveRow = all[idx] || row;
                    const nick = ctx.classRowMailNickname(liveRow);
                    const displayName = ctx.normStr(liveRow.name) || ctx.normStr(liveRow.code) || nick;
                    if (!nick) {
                        ctx.setSummary(
                            'Klasse „' +
                                (liveRow.code || liveRow.name || '?') +
                                '“: Kein Mail-Nickname – Abschlussjahr und Abkürzung prüfen.',
                            'warn'
                        );
                        return;
                    }
                    const memberEmails = ctx.collectStudentEmailsForClassRow(liveRow);
                    const ownerLine = ctx.normStr(liveRow.headEmail) || '–';
                    const memberLine = memberEmails.length
                        ? memberEmails.length + ' Schüler:in(nen) aus der Stammliste'
                        : 'keine Schüler aus der Stammliste';
                    if (
                        !(await ctx.dlgConfirm(
                            'Microsoft-365-Gruppe für „' +
                                (liveRow.code || displayName) +
                                '“ anlegen?\n\n' +
                                'Anzeigename: ' +
                                displayName +
                                '\nMail-Nickname: ' +
                                nick +
                                '\nBesitzer (KV): ' +
                                ownerLine +
                                '\nMitglieder: ' +
                                memberLine,
                            { title: 'Klassengruppe anlegen', okText: 'Anlegen' }
                        ))
                    ) {
                        return;
                    }
                    btnCreateClassGroup.disabled = true;
                    btnCreateClassGroup.setAttribute('aria-busy', 'true');
                    try {
                        const res = await ctx.createClassGroupForRow(liveRow, { memberEmails });
                        if (res && res.error) {
                            ctx.setSummary(res.error, 'warn');
                        } else if (res && res.skipped && res.reason === 'already-linked') {
                            ctx.setSummary('Klasse ist bereits mit einer M365-Gruppe verknüpft.', 'warn');
                        } else if (res && res.created) {
                            ctx.setSummary(
                                'Klassengruppe angelegt: ' +
                                    (res.displayName || displayName) +
                                    ' (' +
                                    (res.mailNickname || nick) +
                                    ').',
                                'ok'
                            );
                            ctx.renderStatusOverview();
                        }
                    } catch (e) {
                        ctx.setSummary('Anlegen: ' + (e && e.message ? e.message : String(e)), 'warn');
                    } finally {
                        btnCreateClassGroup.removeAttribute('aria-busy');
                        ctx.renderClassesTableFromTextarea();
                    }
                });
                tdAction.appendChild(btnCreateClassGroup);
                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'mini-btn';
                btnDel.textContent = '✕';
                btnDel.title = 'Zeile löschen';
                btnDel.addEventListener('click', () => {
                    const all = ctx.getClassesFromTextarea();
                    all.splice(idx, 1);
                    ctx.setClassesTextareaFromRows(all);
                    ctx.renderClassesTableFromTextarea();
                });
                tdAction.appendChild(btnDel);
                collapseActionCell(tdAction, { onEdit: function () { editTableRow(tr); } });

                tr.append(tdCode, tdYear, tdName, tdHead, tdEmail, tdMs, tdAction);
                ctx.classesTbody.appendChild(tr);
            });

            if (ctx.classesTableEditFocus) {
                const focus = ctx.classesTableEditFocus;
                ctx.classesTableEditFocus = null;
                const rowIdx = focus.rowIdx;
                requestAnimationFrame(function () {
                    const tr = ctx.classesTbody.children[rowIdx];
                    if (tr) tr.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
                    ctx.beginClassTableCellEdit(rowIdx, focus.field);
                });
            }
            if (ctx.klassenlehrerPanelApi && typeof ctx.klassenlehrerPanelApi.refresh === 'function') {
                ctx.klassenlehrerPanelApi.refresh();
            }
        }

        ctx.importClassesFromM365Groups = async function importClassesFromM365Groups() {
            try {
                const searchUnified =
                    ctx.graphApi() && typeof ctx.graphApi().searchUnifiedGroups === 'function'
                        ? async function (q, token) {
                              const tok = token || (await ctx.graphApi().getGraphToken());
                              return ctx.graphApi().searchUnifiedGroups(tok, q);
                          }
                        : undefined;
                const picked = await ctx.pickEntraGroupsMulti({
                    title: 'Klassengruppen aus Microsoft 365',
                    hint:
                        'Vorschläge mit jg, klasse und class. Ausgewählte Gruppen werden als Klassen übernommen und für den M365-Abgleich verknüpft (typisches Alias: jg2030-1ak).',
                    initialQueries: ['jg', 'klasse', 'class'],
                    searchFn: searchUnified
                });
                if (!picked.length) return;
                const all = ctx.getClassesFromTextarea();
                const byCode = new Map();
                all.forEach(function (r) {
                    const c = ctx.normCode(r.code);
                    if (c) byCode.set(c, r);
                });
                let added = 0;
                let linked = 0;
                let skipped = 0;
                picked.forEach(function (g) {
                    const derived = ctx.deriveClassEntryFromGroup(g);
                    const code = ctx.normCode(derived.code);
                    if (!code) {
                        skipped++;
                        return;
                    }
                    let row;
                    if (!byCode.has(code)) {
                        row = {
                            code,
                            year: derived.year || '',
                            name: derived.name || code,
                            headName: '',
                            headEmail: ''
                        };
                        byCode.set(code, row);
                        added++;
                    } else {
                        row = byCode.get(code);
                        if (!ctx.normStr(row.year) && derived.year) row.year = derived.year;
                        if (!ctx.normStr(row.name) && derived.name) row.name = derived.name;
                    }
                    ctx.linkClassGroupMatch(row, g);
                    linked++;
                });
                const next = Array.from(byCode.values()).sort(function (a, b) {
                    return ctx.normCode(a.code).localeCompare(ctx.normCode(b.code));
                });
                let headFromOwners = { kvFilled: 0 };
                try {
                    headFromOwners = await ctx.fillRegisterRowsHeadFromGroupOwners(next, picked, {
                        deriveCode: function (g) {
                            return ctx.deriveClassEntryFromGroup(g).code;
                        },
                        getTeachers: ctx.getTeachersFromTextarea,
                        fetchGroupOwners: ctx.graphApi().fetchGroupOwners,
                        getGraphToken: function () {
                            return ctx.graphApi().getGraphToken();
                        }
                    });
                } catch (headErr) {
                    if (typeof console !== 'undefined' && console.warn) {
                        console.warn('Klassen Besitzer:', headErr && headErr.message ? headErr.message : headErr);
                    }
                }
                // Auch bereits verknüpfte Klassen ohne KV: Besitzer aus M365 nachlesen
                let kvExisting = { kvFilled: 0 };
                try {
                    const enrichCtx = ctx.buildClassEnrichContext();
                    const graphKv = await ctx.enrichClassesKvFromGraphOwners(next, {
                        classTeams: enrichCtx.classTeams,
                        classGroupMatchByKey: enrichCtx.classGroupMatchByKey,
                        teachers: ctx.getTeachersFromTextarea(),
                        fetchGroupOwners: ctx.graphApi().fetchGroupOwners,
                        getGraphToken: function () {
                            return ctx.graphApi().getGraphToken();
                        }
                    });
                    if (graphKv && Array.isArray(graphKv.classes)) {
                        next.length = 0;
                        graphKv.classes.forEach(function (r) {
                            next.push(r);
                        });
                    }
                    kvExisting = graphKv.stats || { kvFilled: 0 };
                } catch (kvErr) {
                    if (typeof console !== 'undefined' && console.warn) {
                        console.warn('Klassen KV-Enrich:', kvErr && kvErr.message ? kvErr.message : kvErr);
                    }
                }
                const kvTotal = (headFromOwners.kvFilled || 0) + (kvExisting.kvFilled || 0);
                ctx.setClassesTextareaFromRows(next);
                ctx.renderClassesTableFromTextarea();
                ctx.scheduleAutoSave();
                const parts = [];
                if (added) parts.push(added + ' neue Klasse(n)');
                if (linked) parts.push(linked + ' verknüpft');
                if (kvTotal) {
                    parts.push(kvTotal + ' Klassenvorstand/Vorstände aus Gruppenbesitzer');
                }
                if (skipped) parts.push(skipped + ' übersprungen (kein Kürzel)');
                ctx.setSummary(
                    'Microsoft 365: ' + (parts.length ? parts.join(' · ') : 'Keine Änderung.'),
                    added || linked || kvTotal ? 'ok' : 'warn'
                );
            } catch (e) {
                const msg = e && e.message ? e.message : String(e);
                ctx.setSummary('Import aus Microsoft 365: ' + msg, 'warn');
            }
        }

        ctx.buildClassEnrichContext = function buildClassEnrichContext() {
            const teams = [];
            const matchMap = {};
            let prior = new Map();
            try {
                const api = window.ms365AppDataV2;
                if (api && typeof api.getContainer === 'function') {
                    const c = api.getContainer();
                    if (c && c.core && Array.isArray(c.core.classTeams)) {
                        const raw = c.core.classTeams;
                        teams.push(
                            ...(typeof api.normalizeCoreClassTeams === 'function'
                                ? api.normalizeCoreClassTeams(raw)
                                : raw)
                        );
                    }
                    if (c && c.years && c.years.byLabel) {
                        prior = ctx.collectPriorClassFieldsByCode(c.years.byLabel);
                    }
                }
                const setup =
                    api && typeof api.getSetup === 'function' ? api.getSetup() : {};
                if (setup && setup.classGroupMatchByKey && typeof setup.classGroupMatchByKey === 'object') {
                    Object.assign(matchMap, setup.classGroupMatchByKey);
                }
            } catch {
                // ignore
            }
            const deriveNick =
                typeof window.ms365DeriveClassStableMailNickname === 'function'
                    ? window.ms365DeriveClassStableMailNickname
                    : null;
            return {
                classTeams: teams,
                classGroupMatchByKey: matchMap,
                priorByCode: prior,
                deriveStableMailNickname: deriveNick
            };
        }

        ctx.applyClassEnrichFromLocalLinks = function applyClassEnrichFromLocalLinks(classes, opts) {
            const options = opts || {};
            const result = ctx.enrichClassesFromLinkedGroups(classes || [], ctx.buildClassEnrichContext());
            if (!result.changed) return result;
            if (options.persist && typeof ctx.save === 'function') {
                const cur = ctx.load();
                const saved = ctx.save(Object.assign({}, cur, { classes: result.classes }));
                if (ctx.taClasses) {
                    ctx.taClasses.value = ctx.classesToLines(saved.classes || []);
                }
                ctx.renderClassesTableFromTextarea();
                if (options.summary !== false && ctx.summary) {
                    const st = result.stats || {};
                    ctx.setSummary(
                        'Klassenliste ergänzt: ' +
                            (st.yearFilled || 0) +
                            ' Abschlussjahr(e), ' +
                            (st.kvFilled || 0) +
                            ' Klassenvorstand/KV-Daten aus lokalen Verknüpfungen.',
                        'ok'
                    );
                }
                ctx.dispatchTenantSettingsChanged(saved, options.reason || 'class-enrich-local');
            }
            return result;
        }

        ctx.buildSubjectEnrichContext = function buildSubjectEnrichContext() {
            let legacySubjects = [];
            let structureRows = [];
            let appDataContainer = null;
            try {
                const raw = localStorage.getItem('ms365-tenant-settings-v1');
                if (raw) {
                    const o = JSON.parse(raw);
                    if (o && Array.isArray(o.subjects)) legacySubjects = o.subjects.slice();
                }
            } catch {
                // ignore
            }
            try {
                const api = window.ms365AppDataV2;
                if (api && typeof api.getContainer === 'function') {
                    const c = api.getContainer();
                    appDataContainer = c;
                    if (c && c.structure && Array.isArray(c.structure.rows)) {
                        structureRows = c.structure.rows.slice();
                    }
                }
            } catch {
                // ignore
            }
            const cur = ctx.load();
            return {
                legacySubjects: legacySubjects,
                structureRows: structureRows,
                arges: (cur && cur.arges) || [],
                appDataContainer: appDataContainer,
                storage: typeof localStorage !== 'undefined' ? localStorage : null
            };
        }

        ctx.applySubjectEnrichFromSources = function applySubjectEnrichFromSources(subjects, opts) {
            const options = opts || {};
            const result = ctx.enrichSubjectsFromCatalogSources(subjects || [], ctx.buildSubjectEnrichContext());
            if (!result.changed) return result;
            if (options.persist && typeof ctx.save === 'function') {
                const cur = ctx.load();
                const saved = ctx.save(Object.assign({}, cur, { subjects: result.subjects }));
                if (ctx.taSubjects) {
                    ctx.taSubjects.value = (saved.subjects || [])
                        .map((x) => `${x.code};${x.name || ''}`.trim())
                        .join('\n');
                }
                ctx.renderSubjectsTableFromTextarea();
                if (options.summary !== false && ctx.summary) {
                    const st = result.stats || {};
                    ctx.setSummary(
                        'Fächerliste ergänzt: ' +
                            (st.added || 0) +
                            ' Kürzel neu, ' +
                            (st.namesFilled || 0) +
                            ' Bezeichnungen nachgetragen (Quellen: Unterrichtsteams/Belegung, ARGE, Schulstruktur, ggf. Legacy-Speicher).',
                        'ok'
                    );
                }
                ctx.dispatchTenantSettingsChanged(saved, options.reason || 'subject-enrich-local');
            }
            return result;
        }

}
