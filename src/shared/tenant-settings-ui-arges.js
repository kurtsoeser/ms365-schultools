/**
 * tenant-settings-ui-arges.js
 * @param {Record<string, any>} ctx
 */
import { collapseActionCell, editTableRow, groupLinkStatusView, paintStatusCell } from './tenant-workbench-chrome.js';

export function installTenantArgesPanel(ctx) {
        ctx.argesToLines = function argesToLines(rows) {
            return (rows || [])
                .map((x) => {
                    const list = Array.isArray(x.subjects) ? x.subjects : [];
                    const subj = list.map((s) => ctx.normCode(s)).filter(Boolean).join(',');
                    return `${ctx.normCode(x.code)};${ctx.normStr(x.name || '')};${subj};${ctx.normStr(x.headName || '')};${ctx.normStr(x.headEmail || '').toLowerCase()}`.trim();
                })
                .join('\n');
        }

        ctx.getArgesFromTextarea = function getArgesFromTextarea() {
            return typeof ctx.parseLinesToArges === 'function' ? ctx.parseLinesToArges(ctx.taArges ? ctx.taArges.value : '') : [];
        }

        ctx.setArgesTextareaFromRows = function setArgesTextareaFromRows(rows) {
            if (!ctx.taArges) return;
            ctx.taArges.value = ctx.argesToLines(rows);
        }

        ctx.getSubjectCatalogLink = function getSubjectCatalogLink(code) {
            const api = window.ms365AppDataV2;
            if (api && typeof api.getCatalogLink === 'function') return api.getCatalogLink('subject', code);
            const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
            const links = (setup && setup.catalogLinks) || [];
            const c = ctx.normCode(code);
            return links.find(function (x) {
                return x && x.kind === 'subject' && ctx.normCode(x.code) === c;
            });
        }

        ctx.upsertSubjectCatalogLink = function upsertSubjectCatalogLink(code, group) {
            const api = window.ms365AppDataV2;
            if (!api || typeof api.upsertCatalogLink !== 'function') return;
            api.upsertCatalogLink({
                kind: 'subject',
                code: ctx.normCode(code),
                graphGroupId: group && group.id ? String(group.id) : '',
                displayName: group && group.displayName ? String(group.displayName) : '',
                mailNickname: group && group.mailNickname ? String(group.mailNickname) : '',
                mode: 'matched'
            });
        }

        ctx.subjectGroupMailPrefixForDerive = function subjectGroupMailPrefixForDerive() {
            const api = window.ms365AppDataV2;
            const su = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
            if (su && su.subjectGroupMailPrefix != null && !ctx.normStr(su.subjectGroupMailPrefix)) {
                return '';
            }
            return ctx.normStr(su && su.subjectGroupMailPrefix) || 'fach';
        }

        ctx.deriveArgeEntryFromGroup = function deriveArgeEntryFromGroup(g) {
            const nick = ctx.normStr(g && g.mailNickname).toLowerCase();
            const dn = ctx.normStr(g && g.displayName);
            const prefixes = ['arge-', 'arge.', 'fachgruppe-', 'fachgruppe.', 'fg-', 'fg.', 'ag-', 'ag.'];
            let rest = nick;
            for (let i = 0; i < prefixes.length; i++) {
                if (rest.startsWith(prefixes[i])) {
                    rest = rest.slice(prefixes[i].length);
                    break;
                }
            }
            let code = rest ? ctx.normCode(rest.replace(/[^a-z0-9]/gi, '')) : '';
            if (!code && dn) {
                const m = dn.match(/^(?:ARGE|Fachgruppe|FG|AG)\s*[-–:]?\s*(.+)$/i);
                if (m) {
                    const token = ctx.normStr(m[1]).split(/\s+/)[0] || '';
                    code = ctx.normCode(token);
                }
            }
            if (!code && nick) code = ctx.normCode(nick.replace(/[^a-z0-9]/gi, '').slice(0, 20));
            const name = dn || code || ctx.normStr(g && g.mailNickname);
            return { code, name };
        }

        ctx.buildArgeGroupMatchCell = function buildArgeGroupMatchCell(tr, row) {
            const td = document.createElement('td');
            td.style.fontSize = '0.88em';
            td.style.lineHeight = '1.35';
            const code = ctx.normCode(row && row.code);
            if (!code) {
                paintStatusCell(td, {
                    kind: 'unlinked',
                    label: 'Nicht verknüpft',
                    title: 'Kürzel nötig für den Abgleich'
                });
                return td;
            }
            const link = ctx.getArgeCatalogLink(code);
            paintStatusCell(
                td,
                groupLinkStatusView(
                    link,
                    !code ? 'Kürzel nötig für den Abgleich' : link && link.syncStatus === 'notFound'
                        ? 'Keine passende ARGE-Gruppe gefunden'
                        : 'Noch nicht mit einer Microsoft-365-Gruppe abgeglichen'
                ),
                tr
            );
            return td;
        }

        ctx.importArgesFromM365Groups = async function importArgesFromM365Groups() {
            try {
                const searchUnified =
                    ctx.graphApi() && typeof ctx.graphApi().searchUnifiedGroups === 'function'
                        ? async function (q, token) {
                              const tok = token || (await ctx.graphApi().getGraphToken());
                              return ctx.graphApi().searchUnifiedGroups(tok, q);
                          }
                        : undefined;
                const picked = await ctx.pickEntraGroupsMulti({
                    title: 'ARGE-Gruppen aus Microsoft 365',
                    hint:
                        'Es werden Vorschläge mit den Suchbegriffen arge, fg und ag geladen. Ausgewählte Gruppen werden als ARGE-Stammdaten übernommen und in catalogLinks verknüpft.',
                    initialQueries: ['arge', 'fg', 'ag'],
                    searchFn: searchUnified
                });
                if (!picked.length) return;
                const all = ctx.getArgesFromTextarea();
                const byCode = new Map();
                all.forEach(function (r) {
                    const c = ctx.normCode(r.code);
                    if (c) byCode.set(c, r);
                });
                let added = 0;
                let linked = 0;
                let skipped = 0;
                picked.forEach(function (g) {
                    const derived = ctx.deriveArgeEntryFromGroup(g);
                    const code = ctx.normCode(derived.code);
                    if (!code) {
                        skipped++;
                        return;
                    }
                    if (!byCode.has(code)) {
                        byCode.set(code, {
                            code,
                            name: derived.name || code,
                            subjects: [],
                            headName: '',
                            headEmail: ''
                        });
                        added++;
                    } else {
                        const ex = byCode.get(code);
                        if (!ctx.normStr(ex.name) && derived.name) ex.name = derived.name;
                    }
                    ctx.upsertArgeCatalogLink(code, g);
                    linked++;
                });
                const next = Array.from(byCode.values()).sort(function (a, b) {
                    return ctx.normCode(a.code).localeCompare(ctx.normCode(b.code));
                });
                let headFromOwners = { kvFilled: 0 };
                try {
                    headFromOwners = await ctx.fillRegisterRowsHeadFromGroupOwners(next, picked, {
                        deriveCode: function (g) {
                            return ctx.deriveArgeEntryFromGroup(g).code;
                        },
                        getTeachers: ctx.getTeachersFromTextarea,
                        fetchGroupOwners: ctx.graphApi().fetchGroupOwners,
                        getGraphToken: function () {
                            return ctx.graphApi().getGraphToken();
                        }
                    });
                } catch (headErr) {
                    if (typeof console !== 'undefined' && console.warn) {
                        console.warn('ARGE Besitzer:', headErr && headErr.message ? headErr.message : headErr);
                    }
                }
                ctx.setArgesTextareaFromRows(next);
                ctx.renderArgesTableFromTextarea();
                ctx.scheduleAutoSave();
                const parts = [];
                if (added) parts.push(added + ' neue ARGE');
                if (linked) parts.push(linked + ' verknüpft');
                if (headFromOwners.kvFilled) parts.push(headFromOwners.kvFilled + ' Leitung(en) aus Gruppenbesitzer');
                if (skipped) parts.push(skipped + ' übersprungen (kein Kürzel)');
                ctx.setSummary(
                    'Microsoft 365: ' + (parts.length ? parts.join(' · ') : 'Keine Änderung.'),
                    added || linked || headFromOwners.kvFilled ? 'ok' : 'warn'
                );
            } catch (e) {
                const msg = e && e.message ? e.message : String(e);
                ctx.setSummary('Import aus Microsoft 365: ' + msg, 'warn');
            }
        }

        ctx.buildSubjectM365Cell = function buildSubjectM365Cell(row) {
            const td = document.createElement('td');
            td.style.fontSize = '0.88em';
            td.style.lineHeight = '1.35';
            const code = ctx.normCode(row && row.code);
            if (!code) {
                paintStatusCell(td, {
                    kind: 'unlinked',
                    label: 'Nicht verknüpft',
                    title: 'Kürzel nötig für die Verknüpfung'
                });
                return td;
            }
            const link = ctx.getSubjectCatalogLink(code);
            paintStatusCell(
                td,
                groupLinkStatusView(
                    link,
                    !code
                        ? 'Kürzel nötig für die Verknüpfung'
                        : 'Noch nicht mit einer Microsoft-365-Fachgruppe verknüpft'
                )
            );
            return td;
        }

        ctx.importSubjectsFromM365Groups = async function importSubjectsFromM365Groups() {
            try {
                const searchUnified =
                    ctx.graphApi() && typeof ctx.graphApi().searchUnifiedGroups === 'function'
                        ? async function (q, token) {
                              const tok = token || (await ctx.graphApi().getGraphToken());
                              return ctx.graphApi().searchUnifiedGroups(tok, q);
                          }
                        : undefined;
                const prefix = ctx.subjectGroupMailPrefixForDerive();
                const picked = await ctx.pickEntraGroupsMulti({
                    title: 'Fachgruppen aus Microsoft 365',
                    hint:
                        'Vorschläge mit fach, fachgruppe und Ihrem Präfix („' +
                        prefix +
                        '“). Ausgewählte Gruppen werden als Fachgruppen-Stammdaten übernommen und in catalogLinks verknüpft (wie ARGE & Fachgruppen).',
                    initialQueries: [prefix, 'fach', 'fachgruppe'],
                    searchFn: searchUnified
                });
                if (!picked.length) return;
                const all = ctx.getSubjectsFromTextarea();
                const byCode = new Map();
                all.forEach(function (r) {
                    const c = ctx.normCode(r.code);
                    if (c) byCode.set(c, r);
                });
                let added = 0;
                let linked = 0;
                let skipped = 0;
                const deriveOpts = { mailPrefix: prefix };
                picked.forEach(function (g) {
                    const derived = ctx.deriveSubjectEntryFromGroup(g, deriveOpts);
                    const code = ctx.normCode(derived.code);
                    if (!code) {
                        skipped++;
                        return;
                    }
                    if (!byCode.has(code)) {
                        byCode.set(code, { code, name: derived.name || code });
                        added++;
                    } else {
                        const ex = byCode.get(code);
                        if (!ctx.normStr(ex.name) && derived.name) ex.name = derived.name;
                    }
                    ctx.upsertSubjectCatalogLink(code, g);
                    linked++;
                });
                const next = Array.from(byCode.values()).sort(function (a, b) {
                    return ctx.normCode(a.code).localeCompare(ctx.normCode(b.code));
                });
                ctx.setSubjectsTextareaFromRows(next);
                ctx.renderSubjectsTableFromTextarea();
                ctx.scheduleAutoSave();
                const parts = [];
                if (added) parts.push(added + ' neue Fachgruppe(n)');
                if (linked) parts.push(linked + ' verknüpft');
                if (skipped) parts.push(skipped + ' übersprungen (kein Kürzel)');
                ctx.setSummary(
                    'Microsoft 365: ' + (parts.length ? parts.join(' · ') : 'Keine Änderung.'),
                    added || linked ? 'ok' : 'warn'
                );
            } catch (e) {
                const msg = e && e.message ? e.message : String(e);
                ctx.setSummary('Import aus Microsoft 365: ' + msg, 'warn');
            }
        }

        ctx.commitArgeTableCell = function commitArgeTableCell(idx, field, next, meta) {
            const all = ctx.getArgesFromTextarea();
            if (!all[idx]) {
                ctx.renderArgesTableFromTextarea();
                return;
            }
            if (!meta || !meta.cancelled) {
                if (field === 'code') {
                    const prev = all[idx].code;
                    const newCode = ctx.normCode(next);
                    if (prev && newCode && ctx.normCode(prev) !== ctx.normCode(newCode)) {
                        const api = window.ms365AppDataV2;
                        if (api && typeof api.renameCatalogLink === 'function') {
                            try {
                                api.renameCatalogLink('arge', prev, newCode);
                            } catch {
                                /* Verknüpfung bleibt am alten Kürzel */
                            }
                        }
                    }
                    all[idx].code = newCode;
                } else if (field === 'name') {
                    all[idx].name = ctx.normStr(next);
                } else if (field === 'subjects') {
                    all[idx].subjects = String(next || '')
                        .split(/[,\s|]+/)
                        .map((x) => ctx.normCode(x))
                        .filter(Boolean);
                } else if (field === 'headName') {
                    all[idx].headName = ctx.normStr(next);
                } else if (field === 'headEmail') {
                    all[idx].headEmail = ctx.normStr(next).toLowerCase();
                }
            }
            ctx.setArgesTextareaFromRows(all);
            if (meta && meta.tab) {
                const at = ctx.ARGE_TABLE_EDIT_FIELDS.indexOf(field);
                const step = meta.tab === 'prev' ? -1 : 1;
                const nextField = ctx.ARGE_TABLE_EDIT_FIELDS[at + step];
                if (nextField) {
                    ctx.argesTableEditFocus = { rowIdx: idx, field: nextField };
                }
            }
            ctx.renderArgesTableFromTextarea();
            ctx.scheduleAutoSave();
        }

        ctx.beginArgeTableCellEdit = function beginArgeTableCellEdit(idx, field) {
            if (!ctx.argesTbody) return;
            const tr = ctx.argesTbody.children[idx];
            if (!tr) return;
            const td = tr.querySelector('[data-arge-field="' + field + '"]');
            if (!td) return;
            const rows = ctx.getArgesFromTextarea();
            const row = rows[idx];
            if (!row) return;
            let initial = '';
            if (field === 'code') initial = row.code || '';
            else if (field === 'name') initial = row.name || '';
            else if (field === 'subjects') initial = (Array.isArray(row.subjects) ? row.subjects : []).join(', ');
            else if (field === 'headName') initial = row.headName || '';
            else if (field === 'headEmail') initial = row.headEmail || '';
            ctx.startCellEdit(
                td,
                initial,
                (next, meta) => ctx.commitArgeTableCell(idx, field, next, meta),
                { tabAdvance: true }
            );
        }

        ctx.bindArgeEditableCell = function bindArgeEditableCell(td, idx, field, renderCellContent) {
            td.dataset.argeField = field;
            td.classList.add('is-editable-cell');
            td.title = 'Klicken oder Doppelklick – Tabulator springt zur nächsten Spalte';
            const open = () => ctx.beginArgeTableCellEdit(idx, field);
            td.addEventListener('click', (e) => {
                if (e.target.closest('input.cell-editor')) return;
                open();
            });
            td.addEventListener('dblclick', (e) => {
                e.preventDefault();
                open();
            });
            renderCellContent(td, ctx.getArgesFromTextarea()[idx]);
        }

        ctx.renderArgesTableFromTextarea = function renderArgesTableFromTextarea() {
            if (!ctx.argesTbody) return;
            ctx.updateTableSortIndicators(ctx.argesTable, ctx.argesSortState);
            const rows = ctx.getArgesFromTextarea();
            ctx.argesTbody.replaceChildren();

            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 7;
                td.style.color = 'var(--muted)';
                td.textContent = 'Noch keine Einträge – Importieren oder „+ ARGE“.';
                tr.appendChild(td);
                ctx.argesTbody.appendChild(tr);
                ctx.argesTableEditFocus = null;
                return;
            }

            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');

                const tdCode = document.createElement('td');
                ctx.bindArgeEditableCell(tdCode, idx, 'code', (cell, r) => {
                    cell.innerHTML = `<code>${ctx.escapeHtml((r && r.code) || '')}</code>`;
                });

                const tdName = document.createElement('td');
                ctx.bindArgeEditableCell(tdName, idx, 'name', (cell, r) => {
                    cell.textContent = (r && r.name) || '';
                });

                const tdSubjects = document.createElement('td');
                ctx.bindArgeEditableCell(tdSubjects, idx, 'subjects', (cell, r) => {
                    cell.textContent = (Array.isArray(r && r.subjects) ? r.subjects : []).join(', ');
                });

                const tdHead = document.createElement('td');
                ctx.bindArgeEditableCell(tdHead, idx, 'headName', (cell, r) => {
                    cell.textContent = (r && r.headName) || '';
                });

                const tdEmail = document.createElement('td');
                ctx.bindArgeEditableCell(tdEmail, idx, 'headEmail', (cell, r) => {
                    cell.textContent = (r && r.headEmail) || '';
                });

                const tdM365 = ctx.buildArgeGroupMatchCell(tr, row);

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
                btnCheck.title = 'ARGE-Gruppe in Microsoft 365 prüfen';
                btnCheck.innerHTML = '<i class="bi bi-microsoft" aria-hidden="true"></i>';
                btnCheck.addEventListener('click', async () => {
                    btnCheck.disabled = true;
                    try {
                        const res = await ctx.verifyArgeGroupForRow(row);
                        ctx.setSummary(
                            (row.code || row.name || 'ARGE') +
                                ': ' +
                                (res && res.found ? 'Gruppe gefunden' : 'keine Gruppe gefunden'),
                            res && res.found ? 'ok' : 'warn'
                        );
                    } catch (e) {
                        ctx.setSummary('ARGE-Abgleich: ' + (e && e.message ? e.message : String(e)), 'warn');
                    } finally {
                        ctx.renderArgesTableFromTextarea();
                        ctx.renderStatusOverview();
                    }
                });
                tdAction.appendChild(btnCheck);

                const link = ctx.getArgeCatalogLink(row.code);
                const btnCreate = document.createElement('button');
                btnCreate.type = 'button';
                btnCreate.className = 'mini-btn';
                btnCreate.style.background = '#198754';
                btnCreate.title =
                    link && ctx.normStr(link.graphGroupId)
                        ? 'Bereits mit einer Microsoft-365-Gruppe verknüpft'
                        : 'ARGE-Gruppe in Microsoft 365 anlegen';
                btnCreate.innerHTML = '<i class="bi bi-plus-lg" aria-hidden="true"></i>';
                if (link && ctx.normStr(link.graphGroupId)) {
                    btnCreate.disabled = true;
                    btnCreate.style.opacity = '0.45';
                }
                btnCreate.addEventListener('click', async () => {
                    const all = ctx.getArgesFromTextarea();
                    const liveRow = all[idx] || row;
                    const nick = ctx.argeRowMailNickname(liveRow);
                    const displayName = ctx.normStr(liveRow.name) || ctx.normStr(liveRow.code) || nick;
                    if (!nick) {
                        ctx.setSummary('ARGE „' + (liveRow.code || '?') + '“: Kein Mail-Nickname ableitbar.', 'warn');
                        return;
                    }
                    const ownerLine = ctx.normStr(liveRow.headEmail) || '–';
                    if (
                        !(await ctx.dlgConfirm(
                            'Microsoft-365-Gruppe für „' +
                                (liveRow.code || displayName) +
                                '“ anlegen?\n\n' +
                                'Anzeigename: ' +
                                displayName +
                                '\nMail-Nickname: ' +
                                nick +
                                '\nBesitzer (Leitung): ' +
                                ownerLine,
                            { title: 'ARGE-Gruppe anlegen', okText: 'Anlegen' }
                        ))
                    ) {
                        return;
                    }
                    btnCreate.disabled = true;
                    btnCreate.setAttribute('aria-busy', 'true');
                    try {
                        const res = await ctx.createArgeGroupForRow(liveRow);
                        if (res && res.error) {
                            ctx.setSummary(res.error, 'warn');
                        } else if (res && res.skipped && res.reason === 'already-linked') {
                            ctx.setSummary('ARGE ist bereits mit einer M365-Gruppe verknüpft.', 'warn');
                        } else if (res && res.created) {
                            ctx.setSummary(
                                'ARGE-Gruppe angelegt: ' + (res.displayName || displayName) + ' (' + (res.mailNickname || nick) + ').',
                                'ok'
                            );
                            ctx.renderStatusOverview();
                        }
                    } catch (e) {
                        ctx.setSummary('Anlegen: ' + (e && e.message ? e.message : String(e)), 'warn');
                    } finally {
                        btnCreate.removeAttribute('aria-busy');
                        ctx.renderArgesTableFromTextarea();
                    }
                });
                tdAction.appendChild(btnCreate);

                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'mini-btn';
                btnDel.textContent = '✕';
                btnDel.title = 'Zeile löschen';
                btnDel.addEventListener('click', () => {
                    const all = ctx.getArgesFromTextarea();
                    all.splice(idx, 1);
                    ctx.setArgesTextareaFromRows(all);
                    ctx.renderArgesTableFromTextarea();
                    ctx.scheduleAutoSave();
                });
                tdAction.appendChild(btnDel);
                collapseActionCell(tdAction, { onEdit: function () { editTableRow(tr); } });

                tr.append(tdCode, tdName, tdSubjects, tdHead, tdEmail, tdM365, tdAction);
                ctx.argesTbody.appendChild(tr);
            });

            if (ctx.argesTableEditFocus) {
                const focus = ctx.argesTableEditFocus;
                ctx.argesTableEditFocus = null;
                const rowIdx = focus.rowIdx;
                requestAnimationFrame(function () {
                    const tr = ctx.argesTbody.children[rowIdx];
                    if (tr) tr.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
                    ctx.beginArgeTableCellEdit(rowIdx, focus.field);
                });
            }
        }

}
