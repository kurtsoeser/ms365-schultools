/**
 * tenant-settings-ui-student-council.js
 * @param {Record<string, any>} ctx
 */
import {
    bindOccupancyEdit,
    collapseActionCell,
    editTableRow,
    markEmptyOccupancy,
    presentText,
    presentValue
} from './tenant-workbench-chrome.js';

export function installTenantStudentCouncilPanel(ctx) {
        ctx.studentsToLines = function studentsToLines(rows) {
            return ctx.studentsToLinesPure(rows);
        }

        ctx.getStudentsFromTextarea = function getStudentsFromTextarea() {
            return typeof ctx.parseLinesToStudents === 'function' ? ctx.parseLinesToStudents(ctx.taStudents ? ctx.taStudents.value : '') : [];
        }

        ctx.studentCouncilToLines = function studentCouncilToLines(rows) {
            return (rows || [])
                .map((x) => `${ctx.normStr(x.klasse || '')};${ctx.normStr(x.name || '')};${ctx.normStr(x.email || '').toLowerCase()}`.trim())
                .filter(Boolean)
                .join('\n');
        }

        ctx.getStudentCouncilFromTextarea = function getStudentCouncilFromTextarea() {
            return typeof ctx.parseLinesToStudentCouncil === 'function'
                ? ctx.parseLinesToStudentCouncil(ctx.taStudentCouncil ? ctx.taStudentCouncil.value : '')
                : [];
        }

        ctx.setStudentCouncilTextareaFromRows = function setStudentCouncilTextareaFromRows(rows) {
            if (!ctx.taStudentCouncil) return;
            ctx.taStudentCouncil.value = ctx.studentCouncilToLines(rows);
            if (!ctx.sgaStudentCouncilSyncLock) {
                ctx.pushSgaStudentsFromCouncilRows(rows);
            }
        }

        ctx.studentCouncilPickSlotCount = function studentCouncilPickSlotCount(rows) {
            const n = Array.isArray(rows) ? rows.length : 0;
            return Math.max(ctx.STUDENT_COUNCIL_PICK_MIN_SLOTS, n);
        }

        ctx.studentRowFromPickEmail = function studentRowFromPickEmail(email, students) {
            const em = ctx.normStr(email).toLowerCase();
            if (!em) return { klasse: '', name: '', email: '' };
            const hit = (students || []).find(function (s) {
                return ctx.normStr(s && s.email).toLowerCase() === em;
            });
            if (!hit) return { klasse: '', name: '', email: em };
            return {
                klasse: ctx.normStr(hit.klasse),
                name: ctx.normStr(hit.name),
                email: em
            };
        }

        ctx.commitStudentCouncilPickRows = function commitStudentCouncilPickRows(slotRows) {
            const cleaned = (slotRows || []).filter(function (r) {
                return ctx.normStr(r && r.email) || ctx.normStr(r && r.name) || ctx.normStr(r && r.klasse);
            });
            ctx.setStudentCouncilTextareaFromRows(cleaned);
            ctx.renderStudentCouncilTableFromTextarea();
            ctx.scheduleAutoSave();
        }

        ctx.renderStudentCouncilPickBoard = function renderStudentCouncilPickBoard() {
            if (!ctx.studentCouncilPickSlotsHost) return;
            const rows = ctx.getStudentCouncilFromTextarea();
            const slotN = ctx.studentCouncilPickSlotCount(rows);
            const students = ctx.getStudentsFromTextarea().map(function (s) {
                return Object.assign({}, s, { _sgaPickLabel: ctx.studentPickLabel(s) });
            });
            ctx.studentCouncilPickSlotsHost.replaceChildren();

            for (let i = 0; i < slotN; i++) {
                const slot = rows[i] || { klasse: '', name: '', email: '' };
                const em = presentValue(slot.email).toLowerCase();
                const shownName = presentValue(slot.name);
                const wrap = document.createElement('div');
                wrap.className = 'ts-occupancy' + (em || shownName ? ' is-filled' : ' is-empty');

                const head = document.createElement('div');
                head.className = 'ts-sv-pick-board__head';

                const num = document.createElement('span');
                num.className = 'ts-sga-board__num';
                num.textContent = String(i + 1);
                head.appendChild(num);

                if (slotN > ctx.STUDENT_COUNCIL_PICK_MIN_SLOTS || em) {
                    const btnRm = document.createElement('button');
                    btnRm.type = 'button';
                    btnRm.className = 'ts-sv-pick-board__clear';
                    btnRm.title = 'Platz entfernen';
                    btnRm.setAttribute('aria-label', 'Platz entfernen');
                    btnRm.innerHTML = '<i class="bi bi-x-lg" aria-hidden="true"></i>';
                    btnRm.addEventListener('click', function () {
                        const current = ctx.getStudentCouncilFromTextarea();
                        const slotNNow = ctx.studentCouncilPickSlotCount(current);
                        const next = [];
                        for (let j = 0; j < slotNNow; j++) {
                            if (j !== i) next.push(current[j] ? Object.assign({}, current[j]) : { klasse: '', name: '', email: '' });
                        }
                        ctx.commitStudentCouncilPickRows(next);
                        ctx.renderStudentCouncilPickBoard();
                    });
                    head.appendChild(btnRm);
                }

                wrap.appendChild(head);

                const sel = document.createElement('select');
                sel.className = 'ts-sga-board__pick';
                sel.title = 'Schüler:in aus der Stammdaten-Schülerliste wählen';
                ctx.fillSgaPickSelect(sel, students, em, '— aus Schülerliste wählen —');
                sel.addEventListener('change', function () {
                    const chosen = ctx.normStr(sel.value).toLowerCase();
                    const current = ctx.getStudentCouncilFromTextarea();
                    const slotNNow = ctx.studentCouncilPickSlotCount(current);
                    const next = [];
                    for (let j = 0; j < slotNNow; j++) {
                        next.push(current[j] ? Object.assign({}, current[j]) : { klasse: '', name: '', email: '' });
                    }
                    if (chosen) {
                        const dup = next.some(function (r, idx) {
                            return idx !== i && ctx.normStr(r.email).toLowerCase() === chosen;
                        });
                        if (dup) {
                            ctx.setSummary('Diese:r Schüler:in ist bereits als Vertreter:in eingetragen.', 'warn');
                            sel.value = em || '';
                            return;
                        }
                        next[i] = ctx.studentRowFromPickEmail(chosen, students);
                    } else {
                        next[i] = { klasse: '', name: '', email: '' };
                    }
                    ctx.commitStudentCouncilPickRows(next);
                    ctx.renderStudentCouncilPickBoard();
                });
                wrap.appendChild(sel);

                const preview = document.createElement('div');
                preview.className = 'ts-sga-board__preview muted';
                if (em || shownName) {
                    const who = shownName || em;
                    const meta = [presentValue(slot.klasse), shownName ? em : ''].filter(Boolean).join(' · ');
                    preview.textContent = '';
                    const nameEl = document.createElement('span');
                    nameEl.className = 'ts-occupancy__name';
                    nameEl.textContent = who;
                    preview.appendChild(nameEl);
                    if (meta) {
                        const metaEl = document.createElement('span');
                        metaEl.className = 'ts-occupancy__meta';
                        metaEl.textContent = meta;
                        preview.appendChild(metaEl);
                    }
                } else if (students.length) {
                    markEmptyOccupancy(preview);
                } else {
                    preview.textContent = 'Schülerliste ist leer – zuerst im Tab Schüler pflegen.';
                }
                wrap.appendChild(preview);
                bindOccupancyEdit(wrap, preview, sel);

                ctx.studentCouncilPickSlotsHost.appendChild(wrap);
            }
        }

        ctx.renderStudentCouncilTableFromTextarea = function renderStudentCouncilTableFromTextarea() {
            if (!ctx.studentCouncilTbody) return;
            ctx.updateTableSortIndicators(ctx.studentCouncilTable, ctx.studentCouncilSortState);
            const rows = ctx.getStudentCouncilFromTextarea();
            ctx.studentCouncilTbody.replaceChildren();
            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 5;
                td.style.color = 'var(--muted)';
                td.textContent = 'Noch keine Einträge – oben aus der Schülerliste wählen oder „+ Zeile“.';
                tr.appendChild(td);
                ctx.studentCouncilTbody.appendChild(tr);
                return;
            }
            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');
                const tdKlasse = document.createElement('td');
                tdKlasse.textContent = presentText(row.klasse);
                tdKlasse.title = 'Doppelklick zum Bearbeiten';
                tdKlasse.addEventListener('dblclick', () => {
                    ctx.startCellEdit(tdKlasse, presentValue(row.klasse), (next, meta) => {
                        const all = ctx.getStudentCouncilFromTextarea();
                        if (!all[idx]) return ctx.renderStudentCouncilTableFromTextarea();
                        const prev = all[idx].klasse;
                        all[idx].klasse = meta && meta.cancelled ? prev : ctx.normStr(next);
                        ctx.setStudentCouncilTextareaFromRows(all);
                        ctx.renderStudentCouncilTableFromTextarea();
                        ctx.renderStudentCouncilPickBoard();
                        ctx.scheduleAutoSave();
                    });
                });
                const tdName = document.createElement('td');
                tdName.textContent = presentText(row.name);
                tdName.title = 'Doppelklick zum Bearbeiten';
                tdName.addEventListener('dblclick', () => {
                    ctx.startCellEdit(tdName, presentValue(row.name), (next, meta) => {
                        const all = ctx.getStudentCouncilFromTextarea();
                        if (!all[idx]) return ctx.renderStudentCouncilTableFromTextarea();
                        const prev = all[idx].name;
                        all[idx].name = meta && meta.cancelled ? prev : ctx.normStr(next);
                        ctx.setStudentCouncilTextareaFromRows(all);
                        ctx.renderStudentCouncilTableFromTextarea();
                        ctx.renderStudentCouncilPickBoard();
                        ctx.scheduleAutoSave();
                    });
                });
                const tdEmail = document.createElement('td');
                tdEmail.textContent = presentText(row.email);
                tdEmail.title = 'Doppelklick zum Bearbeiten';
                tdEmail.addEventListener('dblclick', () => {
                    ctx.startCellEdit(tdEmail, presentValue(row.email), (next, meta) => {
                        const all = ctx.getStudentCouncilFromTextarea();
                        if (!all[idx]) return ctx.renderStudentCouncilTableFromTextarea();
                        const prev = all[idx].email;
                        all[idx].email = meta && meta.cancelled ? prev : ctx.normStr(next).toLowerCase();
                        ctx.setStudentCouncilTextareaFromRows(all);
                        ctx.renderStudentCouncilTableFromTextarea();
                        ctx.renderStudentCouncilPickBoard();
                        ctx.scheduleAutoSave();
                    });
                });
                const tdMs = ctx.buildDirectoryMatchCell(tr, row.email);

                const tdAction = document.createElement('td');
                tdAction.className = 'action-cell';
                const emSv = ctx.normStr(row.email || '').toLowerCase();
                const dirSv = emSv && emSv.indexOf('@') !== -1 ? ctx.getDirectoryMatchByEmail(emSv) : null;
                const btnCheck = document.createElement('button');
                btnCheck.type = 'button';
                btnCheck.className = 'mini-btn';
                btnCheck.title = 'In Microsoft Entra prüfen';
                btnCheck.innerHTML = '<i class="bi bi-microsoft" aria-hidden="true"></i>';
                if (dirSv && dirSv.graphUserId) {
                    btnCheck.disabled = true;
                } else {
                btnCheck.addEventListener('click', async () => {
                    const em = ctx.normStr(row.email || '').toLowerCase();
                    if (!em || em.indexOf('@') === -1) return ctx.setSummary('Keine gültige E-Mail für den Abgleich.', 'warn');
                    try {
                        const token = await ctx.graphApi().getGraphToken();
                        const u = await ctx.graphApi().resolveUserByEmail(token, em);
                        const iso = new Date().toISOString();
                        const setup = window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function' ? window.ms365AppDataV2.getSetup() : {};
                        const updates = Object.assign({}, (setup && setup.directoryMatchByEmail) || {});
                        if (u && u.id) {
                            updates[em] = ctx.directoryMatchUserPayload(u, iso);
                            ctx.setSummary(row.name + ': gefunden (' + (u.displayName || em) + ')', 'ok');
                        } else {
                            updates[em] = { notFound: true, checkedAt: iso };
                            ctx.setSummary(row.name + ': nicht in Microsoft Entra gefunden.', 'warn');
                        }
                        window.ms365AppDataV2.patchSetup({ directoryMatchByEmail: updates });
                        ctx.renderStudentCouncilTableFromTextarea();
                        ctx.renderStudentCouncilPickBoard();
                    } catch (e) {
                        ctx.setSummary('Fehler beim Prüfen: ' + (e && e.message ? e.message : String(e)), 'warn');
                    }
                });
                }
                tdAction.appendChild(btnCheck);

                const btnCreate = document.createElement('button');
                btnCreate.type = 'button';
                btnCreate.className = 'mini-btn';
                btnCreate.title = 'Benutzer in Microsoft Entra anlegen';
                btnCreate.innerHTML = '<i class="bi bi-person-plus" aria-hidden="true"></i>';
                if (dirSv && dirSv.graphUserId) {
                    btnCreate.disabled = true;
                } else {
                btnCreate.addEventListener('click', () => {
                    if (typeof window.ms365TenantSettingsCreateUser === 'function') {
                        window.ms365TenantSettingsCreateUser(row, () => ctx.renderStudentCouncilTableFromTextarea());
                    } else {
                        ctx.setSummary('Anlegen nicht verfügbar – Modul nicht geladen.', 'warn');
                    }
                });
                }
                tdAction.appendChild(btnCreate);

                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'mini-btn';
                btnDel.textContent = '✕';
                btnDel.title = 'Zeile löschen';
                btnDel.addEventListener('click', () => {
                    const all = ctx.getStudentCouncilFromTextarea();
                    all.splice(idx, 1);
                    ctx.setStudentCouncilTextareaFromRows(all);
                    ctx.renderStudentCouncilTableFromTextarea();
                    ctx.renderStudentCouncilPickBoard();
                    ctx.scheduleAutoSave();
                });
                tdAction.appendChild(btnDel);
                ctx.applyDirectoryRowActionCellStyles(tdAction, btnCheck, btnCreate);
                collapseActionCell(tdAction, { onEdit: function () { editTableRow(tr); } });

                tr.append(tdKlasse, tdName, tdEmail, tdMs, tdAction);
                ctx.studentCouncilTbody.appendChild(tr);
            });
        }

}
