/**
 * tenant-settings-ui-students-panel.js
 * @param {Record<string, any>} ctx
 */
import { collapseActionCell, editTableRow, presentText, presentValue } from './tenant-workbench-chrome.js';

export function installTenantStudentsPanel(ctx) {
        ctx.renderStudentLifecyclePanel = function renderStudentLifecyclePanel(prev, next) {
            const host = document.getElementById('tenantStudentLifecycle');
            const listEl = document.getElementById('tenantStudentLifecycleList');
            const metaEl = document.getElementById('tenantStudentLifecycleMeta');
            const lc = ctx.lifecycleApi();
            if (!host || !lc) return;
            const diff = lc.diffStudents(prev, next);
            const ctx = ctx.getLifecycleContext();
            const preview = lc.previewMemberships(diff, ctx.classTeams, ctx.schuelerGroupId);
            ctx.lastLifecyclePreview = preview;
            if (!lc.hasMembershipWork(preview)) {
                host.hidden = true;
                if (listEl) listEl.replaceChildren();
                return;
            }
            const sum = lc.summarizePreview(preview);
            host.hidden = false;
            if (metaEl) {
                metaEl.textContent =
                    sum.join +
                    ' Aufnahme(n), ' +
                    sum.leave +
                    ' Entfernung(en) in ' +
                    sum.groupCount +
                    ' Gruppe(n). Nur E-Mails mit zugeordneter Microsoft-365-Gruppe erscheinen hier.';
            }
            if (listEl) {
                listEl.replaceChildren();
                preview.groups.forEach(function (g) {
                    const li = document.createElement('li');
                    const joinN = g.join.length;
                    const leaveN = g.leave.length;
                    li.textContent =
                        (g.label || g.groupId) +
                        ': +' +
                        joinN +
                        (joinN ? ' (' + g.join.slice(0, 4).join(', ') + (joinN > 4 ? ' …' : '') + ')' : '') +
                        ', −' +
                        leaveN +
                        (leaveN ? ' (' + g.leave.slice(0, 4).join(', ') + (leaveN > 4 ? ' …' : '') + ')' : '');
                    listEl.appendChild(li);
                });
            }
        }

        ctx.applyStudentLifecyclePreview = async function applyStudentLifecyclePreview() {
            const lc = ctx.lifecycleApi();
            const gug = window.ms365GraphUnifiedGroups;
            if (!lc || !gug || !ctx.lastLifecyclePreview || !lc.hasMembershipWork(ctx.lastLifecyclePreview)) {
                ctx.setSummary('Keine Mitgliedschaftsänderungen zum Anwenden.', 'warn');
                return;
            }
            const btn = document.getElementById('tenantStudentLifecycleApply');
            if (btn) btn.disabled = true;
            try {
                const token = await gug.getGraphToken();
                let ok = 0;
                let fail = 0;
                for (let i = 0; i < ctx.lastLifecyclePreview.groups.length; i++) {
                    const g = ctx.lastLifecyclePreview.groups[i];
                    if (g.join && g.join.length) {
                        const r = await gug.syncEmailsToGroup(token, g.groupId, g.join, g.label, function () {});
                        ok += r.ok || 0;
                        fail += r.fail || 0;
                    }
                    if (g.leave && g.leave.length && typeof gug.removeEmailsFromGroup === 'function') {
                        const r = await gug.removeEmailsFromGroup(token, g.groupId, g.leave, g.label, function () {});
                        ok += r.ok || 0;
                        fail += r.fail || 0;
                    }
                }
                ctx.setSummary(
                    'Mitgliedschaften in Microsoft 365 angewendet: ' + ok + ' OK' + (fail ? ', ' + fail + ' Fehler' : '') + '.',
                    fail ? 'warn' : 'ok'
                );
                const host = document.getElementById('tenantStudentLifecycle');
                if (host) host.hidden = true;
            } catch (e) {
                ctx.setSummary('Mitgliedschaften: ' + (e.message || e), 'warn');
            } finally {
                if (btn) btn.disabled = false;
            }
        }

        ctx.setStudentsTextareaFromRows = function setStudentsTextareaFromRows(rows) {
            if (!ctx.taStudents) return;
            ctx.taStudents.value = ctx.studentsToLines(rows);
        }

        ctx.getStudentsDirectoryMatchMap = function getStudentsDirectoryMatchMap() {
            try {
                const api = window.ms365AppDataV2;
                const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
                return setup && setup.directoryMatchByEmail ? setup.directoryMatchByEmail : {};
            } catch {
                return {};
            }
        }

        /** Update one student field without rebuilding all ~N table rows. */
        ctx.commitStudentCellEdit = function commitStudentCellEdit(idx, field, next, meta, td) {
            const prevRows = ctx.getStudentsFromTextarea();
            const all = prevRows.slice();
            if (!all[idx]) {
                ctx.renderStudentsTableFromTextarea();
                return;
            }
            const prev = all[idx][field];
            let value = prev;
            if (!(meta && meta.cancelled)) {
                if (field === 'email') value = ctx.normStr(next).toLowerCase();
                else value = ctx.normStr(next);
            }
            all[idx][field] = value;
            ctx.setStudentsTextareaFromRows(all);
            if (field === 'email') {
                ctx.renderStudentSammelLifecyclePanel(prevRows, all);
            }

            const filterActive =
                (ctx.studentsClassFilter && ctx.studentsClassFilter.value) ||
                (ctx.studentsTextFilter && String(ctx.studentsTextFilter.value || '').trim());
            if (field === 'klasse' || filterActive) {
                ctx.renderStudentsTableFromTextarea();
                ctx.scheduleAutoSave();
                return;
            }

            td.textContent = value || '';
            td.title = 'Doppelklick zum Bearbeiten';

            if (field === 'email') {
                const tr = td.parentElement;
                if (tr) {
                    const oldMs = tr.children[5];
                    const newMs = ctx.buildDirectoryMatchCell(tr, value, ctx.getStudentsDirectoryMatchMap());
                    if (oldMs) tr.replaceChild(newMs, oldMs);
                    else tr.appendChild(newMs);
                    tr.style.background = '';
                    const match = value ? ctx.getDirectoryMatchByEmail(value) : null;
                    if (match && match.graphUserId) {
                        tr.style.background = 'color-mix(in srgb, #0d8050 8%, transparent)';
                    }
                }
            }
            ctx.scheduleAutoSave();
        }

        ctx.updateStudentsSortIndicators = function updateStudentsSortIndicators() {
            ctx.updateTableSortIndicators(ctx.studentsTable, ctx.studentsSortState);
        }

        ctx.applySubjectsSort = function applySubjectsSort(key) {
            if (!key || (key !== 'code' && key !== 'name')) return;
            ctx.toggleTenantSortState(ctx.subjectsSortState, key);
            ctx.setSubjectsTextareaFromRows(ctx.sortTenantRows(ctx.getSubjectsFromTextarea(), key, ctx.subjectsSortState.dir));
            ctx.renderSubjectsTableFromTextarea();
            ctx.scheduleAutoSave();
        }

        ctx.applyArgesSort = function applyArgesSort(key) {
            if (!key || (key !== 'code' && key !== 'name' && key !== 'subjects')) return;
            ctx.toggleTenantSortState(ctx.argesSortState, key);
            ctx.setArgesTextareaFromRows(ctx.sortTenantRows(ctx.getArgesFromTextarea(), key, ctx.argesSortState.dir));
            ctx.renderArgesTableFromTextarea();
            ctx.scheduleAutoSave();
        }

        ctx.applyClassesSort = function applyClassesSort(key) {
            const allowed = ['code', 'year', 'name', 'headName', 'headEmail'];
            if (!key || allowed.indexOf(key) < 0) return;
            ctx.toggleTenantSortState(ctx.classesSortState, key);
            ctx.setClassesTextareaFromRows(ctx.sortTenantRows(ctx.getClassesFromTextarea(), key, ctx.classesSortState.dir));
            ctx.renderClassesTableFromTextarea();
            ctx.scheduleAutoSave();
        }

        ctx.applySgaSort = function applySgaSort(key) {
            if (!key || (key !== 'scope' && key !== 'name' && key !== 'email')) return;
            ctx.toggleTenantSortState(ctx.sgaSortState, key);
            ctx.setSgaTextareaFromRows(ctx.sortTenantRows(ctx.getSgaFromTextarea(), key, ctx.sgaSortState.dir));
            ctx.renderSgaTableFromTextarea();
            ctx.scheduleAutoSave();
        }

        ctx.applyStudentCouncilSort = function applyStudentCouncilSort(key) {
            if (!key || (key !== 'klasse' && key !== 'name' && key !== 'email')) return;
            ctx.toggleTenantSortState(ctx.studentCouncilSortState, key);
            ctx.setStudentCouncilTextareaFromRows(
                ctx.sortTenantRows(ctx.getStudentCouncilFromTextarea(), key, ctx.studentCouncilSortState.dir)
            );
            ctx.renderStudentCouncilTableFromTextarea();
            ctx.scheduleAutoSave();
        }

        ctx.syncStudentsClassFilterOptions = function syncStudentsClassFilterOptions(rows) {
            if (!ctx.studentsClassFilter) return;
            const prev = String(ctx.studentsClassFilter.value || '');
            const codes = new Set();
            let hasEmpty = false;
            (rows || []).forEach((row) => {
                const k = ctx.normStr(row && row.klasse);
                if (!k) hasEmpty = true;
                else codes.add(k);
            });
            const sorted = Array.from(codes).sort((a, b) => a.localeCompare(b, 'de', { sensitivity: 'base' }));
            const signature = (hasEmpty ? '1' : '0') + '|' + sorted.join('\n');
            if (ctx.studentsClassFilter.dataset.sig !== signature) {
                ctx.studentsClassFilter.dataset.sig = signature;
                ctx.studentsClassFilter.replaceChildren();
                const optAll = document.createElement('option');
                optAll.value = '';
                optAll.textContent = 'Alle Klassen';
                ctx.studentsClassFilter.appendChild(optAll);
                if (hasEmpty) {
                    const optEmpty = document.createElement('option');
                    optEmpty.value = '__empty__';
                    optEmpty.textContent = 'Ohne Klasse';
                    ctx.studentsClassFilter.appendChild(optEmpty);
                }
                sorted.forEach((code) => {
                    const opt = document.createElement('option');
                    opt.value = code;
                    opt.textContent = code;
                    ctx.studentsClassFilter.appendChild(opt);
                });
            }
            let next = '';
            if (prev === '__empty__' && hasEmpty) {
                next = '__empty__';
            } else if (prev) {
                const match = sorted.find((c) => c.toUpperCase() === prev.toUpperCase());
                if (match) next = match;
            }
            if (ctx.studentsClassFilter.value !== next) {
                ctx.studentsClassFilter.value = next;
            }
        }

        ctx.studentRowMatchesFilters = function studentRowMatchesFilters(row, classFilter, textQuery) {
            const klasse = ctx.normStr(row && row.klasse);
            if (classFilter === '__empty__') {
                if (klasse) return false;
            } else if (classFilter) {
                if (klasse.toUpperCase() !== classFilter.toUpperCase()) return false;
            }
            const tokens = String(textQuery || '')
                .trim()
                .toLowerCase()
                .split(/\s+/)
                .filter(Boolean);
            if (!tokens.length) return true;
            const parts = [klasse, row && row.name, row && row.email];
            const pairs = Array.isArray(row && row.parentPairs) ? row.parentPairs : [];
            pairs.forEach((p) => {
                if (!p) return;
                parts.push(p.name, p.email);
            });
            const hay = parts
                .map((v) => String(v || '').toLowerCase())
                .join('\n');
            return tokens.every((t) => hay.includes(t));
        }

        ctx.updateStudentsFilterMeta = function updateStudentsFilterMeta(total, visible) {
            if (!ctx.studentsFilterMeta) return;
            const classActive = ctx.studentsClassFilter && ctx.studentsClassFilter.value;
            const textActive = ctx.studentsTextFilter && String(ctx.studentsTextFilter.value || '').trim();
            if (!total || (!classActive && !textActive)) {
                ctx.studentsFilterMeta.hidden = true;
                ctx.studentsFilterMeta.textContent = '';
                return;
            }
            ctx.studentsFilterMeta.hidden = false;
            ctx.studentsFilterMeta.textContent =
                visible === total
                    ? total + ' Einträge angezeigt'
                    : visible + ' von ' + total + ' Einträgen angezeigt';
        }

        ctx.applyStudentsSort = function applyStudentsSort(key) {
            if (!key || (key !== 'klasse' && key !== 'name' && key !== 'email')) return;
            if (ctx.studentsSortState.key === key) {
                ctx.studentsSortState.dir = ctx.studentsSortState.dir === 1 ? -1 : 1;
            } else {
                ctx.studentsSortState.key = key;
                ctx.studentsSortState.dir = 1;
            }
            const api = window.ms365StudentListImport;
            const all = ctx.getStudentsFromTextarea();
            const sorted =
                api && typeof api.sortRows === 'function'
                    ? api.sortRows(all, ctx.studentsSortState.key, ctx.studentsSortState.dir)
                    : all.slice().sort((a, b) => {
                          const av = String((a && a[key]) || '');
                          const bv = String((b && b[key]) || '');
                          return av.localeCompare(bv, 'de', { sensitivity: 'base' }) * ctx.studentsSortState.dir;
                      });
            ctx.setStudentsTextareaFromRows(sorted);
            ctx.renderStudentsTableFromTextarea();
            ctx.scheduleAutoSave();
        }

        ctx.renderStudentsTableFromTextarea = function renderStudentsTableFromTextarea() {
            if (!ctx.studentsTbody) return;
            const rows = ctx.getStudentsFromTextarea();
            ctx.updateStudentsSortIndicators();
            ctx.syncStudentsClassFilterOptions(rows);
            const classFilter = ctx.studentsClassFilter ? String(ctx.studentsClassFilter.value || '') : '';
            const textQuery = ctx.studentsTextFilter ? String(ctx.studentsTextFilter.value || '') : '';
            const visible = [];
            rows.forEach((row, idx) => {
                if (ctx.studentRowMatchesFilters(row, classFilter, textQuery)) {
                    visible.push({ row, idx });
                }
            });
            ctx.updateStudentsFilterMeta(rows.length, visible.length);

            if (!rows.length) {
                ctx.studentsTbody.replaceChildren();
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 7;
                td.style.color = 'var(--muted)';
                td.textContent = 'Noch keine Einträge – Importieren oder „+ Schüler:in“.';
                tr.appendChild(td);
                ctx.studentsTbody.appendChild(tr);
                return;
            }

            if (!visible.length) {
                ctx.studentsTbody.replaceChildren();
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 7;
                td.style.color = 'var(--muted)';
                td.textContent = 'Keine Treffer für Klasse/Suche – Filter anpassen.';
                tr.appendChild(td);
                ctx.studentsTbody.appendChild(tr);
                return;
            }

            const dirMatchMap = ctx.getStudentsDirectoryMatchMap();

            function buildParentCell(pair, extraCount) {
                const td = document.createElement('td');
                td.className = 'students-parent-cell';
                const p = pair || {};
                const pName = presentValue(p.name);
                const pMail = presentValue(p.email).toLowerCase();
                if (!pName && !pMail) {
                    td.textContent = '–';
                    td.style.color = 'var(--muted)';
                    td.title = 'Kein Erziehungsberechtigter';
                    return td;
                }
                td.textContent = pName || pMail;
                if (extraCount > 0) {
                    const more = document.createElement('span');
                    more.className = 'students-parent-more';
                    more.textContent = ' +' + extraCount;
                    td.appendChild(more);
                }
                const tipParts = [];
                if (pName && pMail) tipParts.push(pName);
                if (pMail) tipParts.push(pMail);
                else if (pName) tipParts.push(pName);
                if (extraCount > 0) tipParts.push('(+ ' + extraCount + ' weitere)');
                td.title = tipParts.join(' · ') || 'Erziehungsberechtigte:r';
                return td;
            }

            const frag = document.createDocumentFragment();
            visible.forEach(({ row, idx }) => {
                const tr = document.createElement('tr');
                tr.dataset.studentIdx = String(idx);

                const tdClass = document.createElement('td');
                const klasseShown = presentValue(row.klasse);
                tdClass.innerHTML = klasseShown
                    ? '<code>' + ctx.escapeHtml(klasseShown) + '</code>'
                    : '–';
                tdClass.title = 'Doppelklick zum Bearbeiten';
                tdClass.addEventListener('dblclick', () => {
                    ctx.startCellEdit(tdClass, klasseShown, (next, meta) => {
                        ctx.commitStudentCellEdit(idx, 'klasse', next, meta, tdClass);
                    });
                });

                const tdName = document.createElement('td');
                tdName.textContent = presentText(row.name);
                tdName.title = 'Doppelklick zum Bearbeiten';
                tdName.addEventListener('dblclick', () => {
                    ctx.startCellEdit(tdName, presentValue(row.name), (next, meta) => {
                        ctx.commitStudentCellEdit(idx, 'name', next, meta, tdName);
                    });
                });

                const tdEmail = document.createElement('td');
                tdEmail.textContent = presentText(row.email);
                tdEmail.title = 'Doppelklick zum Bearbeiten';
                tdEmail.addEventListener('dblclick', () => {
                    ctx.startCellEdit(tdEmail, presentValue(row.email), (next, meta) => {
                        ctx.commitStudentCellEdit(idx, 'email', next, meta, tdEmail);
                    });
                });

                const pairs = Array.isArray(row.parentPairs) ? row.parentPairs : [];
                const tdParent1 = buildParentCell(pairs[0], 0);
                const tdParent2 = buildParentCell(pairs[1], Math.max(0, pairs.length - 2));

                const tdMs = ctx.buildDirectoryMatchCell(tr, row.email, dirMatchMap);

                const tdAction = document.createElement('td');
                tdAction.className = 'action-cell';
                tdAction.style.whiteSpace = 'nowrap';
                tdAction.style.display = 'flex';
                tdAction.style.gap = '6px';
                tdAction.style.alignItems = 'center';
                const emKey = ctx.normStr(row.email || '').toLowerCase();
                const dir = emKey ? dirMatchMap[emKey] || null : null;
                if (row.email && row.name) {
                    const btnCheck = document.createElement('button');
                    btnCheck.type = 'button';
                    btnCheck.className = 'mini-btn';
                    btnCheck.style.background = 'var(--brand1)';
                    btnCheck.title = 'Diese E‑Mail in Microsoft Entra prüfen';
                    btnCheck.innerHTML = '<i class="bi bi-microsoft" aria-hidden="true"></i>';
                    if (dir && dir.graphUserId) {
                        btnCheck.disabled = true;
                    } else {
                        btnCheck.addEventListener('click', async () => {
                            btnCheck.disabled = true;
                            try {
                                await ctx.verifyAdminDirectoryEmail(row.email);
                            } finally {
                                ctx.renderStudentsTableFromTextarea();
                            }
                        });
                    }
                    tdAction.appendChild(btnCheck);

                    const btnCreate = document.createElement('button');
                    btnCreate.type = 'button';
                    btnCreate.className = 'mini-btn';
                    btnCreate.style.background = '#11cdef';
                    btnCreate.title = 'Neuen Benutzer in Microsoft Entra ID anlegen (User.ReadWrite.All)';
                    btnCreate.innerHTML = '<i class="bi bi-person-plus" aria-hidden="true"></i>';
                    if (dir && dir.graphUserId) {
                        btnCreate.disabled = true;
                    } else {
                        btnCreate.addEventListener('click', async () => {
                            btnCreate.disabled = true;
                            try {
                                await ctx.createAdminEntraUserInteractive(row.email, row.name);
                            } finally {
                                ctx.renderStudentsTableFromTextarea();
                            }
                        });
                    }
                    tdAction.appendChild(btnCreate);
                }
                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'mini-btn';
                btnDel.textContent = '✕';
                btnDel.title = 'Zeile löschen';
                btnDel.addEventListener('click', () => {
                    const all = ctx.getStudentsFromTextarea();
                    all.splice(idx, 1);
                    ctx.setStudentsTextareaFromRows(all);
                    ctx.renderStudentsTableFromTextarea();
                    ctx.scheduleAutoSave();
                });
                tdAction.appendChild(btnDel);
                collapseActionCell(tdAction, { onEdit: function () { editTableRow(tr); } });

                tr.append(tdClass, tdName, tdEmail, tdParent1, tdParent2, tdMs, tdAction);
                frag.appendChild(tr);
            });
            ctx.studentsTbody.replaceChildren(frag);
        }

        ctx.classesToLines = function classesToLines(rows) {
            return ctx.classesToLinesPure(rows);
        }

        ctx.getClassesFromTextarea = function getClassesFromTextarea() {
            return typeof ctx.parseLinesToClasses === 'function' ? ctx.parseLinesToClasses(ctx.taClasses ? ctx.taClasses.value : '') : [];
        }

        ctx.setClassesTextareaFromRows = function setClassesTextareaFromRows(rows) {
            if (!ctx.taClasses) return;
            ctx.taClasses.value = ctx.classesToLines(rows);
        }

        ctx.commitClassTableCell = function commitClassTableCell(idx, field, next, meta) {
            const all = ctx.getClassesFromTextarea();
            if (!all[idx]) {
                ctx.renderClassesTableFromTextarea();
                return;
            }
            if (!meta || !meta.cancelled) {
                if (field === 'code') {
                    all[idx].code = ctx.normCode(next);
                } else if (field === 'year') {
                    const n = ctx.normStr(next);
                    all[idx].year = /^\d{4}$/.test(n) ? n : '';
                } else if (field === 'name') {
                    all[idx].name = ctx.normStr(next);
                } else if (field === 'headName') {
                    all[idx].headName = ctx.normStr(next);
                } else if (field === 'headEmail') {
                    all[idx].headEmail = ctx.normStr(next).toLowerCase();
                }
            }
            ctx.setClassesTextareaFromRows(all);
            if (meta && meta.tab) {
                const at = ctx.CLASS_TABLE_EDIT_FIELDS.indexOf(field);
                const step = meta.tab === 'prev' ? -1 : 1;
                const nextField = ctx.CLASS_TABLE_EDIT_FIELDS[at + step];
                if (nextField) {
                    ctx.classesTableEditFocus = { rowIdx: idx, field: nextField };
                }
            }
            ctx.renderClassesTableFromTextarea();
            ctx.scheduleAutoSave();
        }

        ctx.collectStudentEmailsForClassRow = function collectStudentEmailsForClassRow(row) {
            if (!row || typeof ctx.getStudentsFromTextarea !== 'function') return [];
            const seen = new Set();
            const out = [];
            ctx.getStudentsFromTextarea().forEach(function (s) {
                if (!ctx.studentBelongsToClassRow(s, row, null)) return;
                const em = ctx.normStr(s && s.email).toLowerCase();
                if (!em || em.indexOf('@') === -1 || seen.has(em)) return;
                seen.add(em);
                out.push(em);
            });
            return out;
        }

        ctx.beginClassTableCellEdit = function beginClassTableCellEdit(idx, field) {
            if (!ctx.classesTbody) return;
            const tr = ctx.classesTbody.children[idx];
            if (!tr) return;
            const td = tr.querySelector('[data-class-field="' + field + '"]');
            if (!td) return;
            const rows = ctx.getClassesFromTextarea();
            const row = rows[idx];
            if (!row) return;
            let initial = '';
            if (field === 'code') initial = presentValue(row.code);
            else if (field === 'year') initial = presentValue(row.year);
            else if (field === 'name') initial = presentValue(row.name);
            else if (field === 'headName') initial = presentValue(row.headName);
            else if (field === 'headEmail') initial = presentValue(row.headEmail);
            ctx.startCellEdit(
                td,
                initial,
                (next, meta) => ctx.commitClassTableCell(idx, field, next, meta),
                { tabAdvance: true }
            );
        }

        ctx.bindClassEditableCell = function bindClassEditableCell(td, idx, field, renderCellContent) {
            td.dataset.classField = field;
            td.classList.add('is-editable-cell');
            td.title = 'Klicken oder Doppelklick – Tabulator springt zur nächsten Spalte';
            const open = () => ctx.beginClassTableCellEdit(idx, field);
            td.addEventListener('click', (e) => {
                if (e.target.closest('input.cell-editor')) return;
                open();
            });
            td.addEventListener('dblclick', (e) => {
                e.preventDefault();
                open();
            });
            renderCellContent(td, ctx.getClassesFromTextarea()[idx]);
        }

}
