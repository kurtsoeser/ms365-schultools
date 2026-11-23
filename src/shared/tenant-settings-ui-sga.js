/**
 * tenant-settings-ui-sga.js
 * @param {Record<string, any>} ctx
 */
import {
    bindOccupancyEdit,
    collapseActionCell,
    editTableRow,
    markEmptyOccupancy,
    presentValue
} from './tenant-workbench-chrome.js';

export function installTenantSgaPanel(ctx) {
        ctx.sgaToLines = function sgaToLines(rows) {
            return (rows || [])
                .map((x) => {
                    const scope =
                        x.scope === 'teacher'
                            ? 'Lehrer'
                            : x.scope === 'student'
                              ? 'Schueler'
                              : x.scope === 'external'
                                ? 'Extern'
                                : '';
                    return `${scope};${ctx.normStr(x.name || '')};${ctx.normStr(x.email || '').toLowerCase()}`.trim();
                })
                .filter(Boolean)
                .join('\n');
        }

        ctx.getSgaFromTextarea = function getSgaFromTextarea() {
            return typeof ctx.parseLinesToSga === 'function' ? ctx.parseLinesToSga(ctx.taSga ? ctx.taSga.value : '') : [];
        }

        ctx.setSgaTextareaFromRows = function setSgaTextareaFromRows(rows) {
            if (!ctx.taSga) return;
            ctx.taSga.value = ctx.sgaToLines(rows);
        }

        /** @type {{ scope: string, name: string, email: string }[]} */
        ctx.sgaBoardOverflowRows = [];

        ctx.sgaScopeLabel = function sgaScopeLabel(scope) {
            if (typeof window.ms365TenantSettingsSgaScopeLabel === 'function') {
                return window.ms365TenantSettingsSgaScopeLabel(scope);
            }
            if (scope === 'teacher') return 'Lehrer';
            if (scope === 'student') return 'Schülervertretung';
            if (scope === 'external') return 'Extern (Eltern)';
            return scope || '—';
        }

        ctx.partitionSgaForBoard = function partitionSgaForBoard() {
            const fn = window.ms365TenantSettingsSgaPartition;
            return typeof fn === 'function' ? fn(ctx.getSgaFromTextarea()) : { board: { teachers: [], students: [], externals: [] }, overflow: [] };
        }

        /** Verhindert Ping-Pong zwischen SGA-Schülerplätzen und Schülervertretung. */
        ctx.sgaStudentCouncilSyncLock = false;

        ctx.sgaStudentSlotCount = function sgaStudentSlotCount() {
            return typeof window.ms365TenantSettingsSgaBoardSlots === 'number'
                ? window.ms365TenantSettingsSgaBoardSlots
                : 3;
        }

        ctx.pushSgaStudentsFromCouncilRows = function pushSgaStudentsFromCouncilRows(councilRows) {
            if (ctx.sgaStudentCouncilSyncLock) return;
            const fn = window.ms365TenantSettingsSgaRowsFromBoard;
            if (typeof fn !== 'function') return;
            ctx.sgaStudentCouncilSyncLock = true;
            try {
                const part = ctx.partitionSgaForBoard();
                const b = part.board || { teachers: [], students: [], externals: [] };
                b.students = ctx.sgaStudentsFromStudentCouncil(councilRows, ctx.sgaStudentSlotCount());
                ctx.sgaBoardOverflowRows = part.overflow || [];
                ctx.setSgaTextareaFromRows(fn(b, ctx.sgaBoardOverflowRows));
                ctx.renderSgaTableFromTextarea();
                if (typeof ctx.renderSgaBoard === 'function') ctx.renderSgaBoard();
            } finally {
                ctx.sgaStudentCouncilSyncLock = false;
            }
        }

        ctx.pushStudentCouncilFromSgaBoard = function pushStudentCouncilFromSgaBoard(board) {
            if (!ctx.taStudentCouncil) return;
            const slotN = ctx.sgaStudentSlotCount();
            const next = ctx.studentCouncilFromSgaStudents(
                (board && board.students) || [],
                ctx.getStudentsFromTextarea(),
                ctx.getStudentCouncilFromTextarea(),
                slotN
            );
            ctx.sgaStudentCouncilSyncLock = true;
            try {
                ctx.setStudentCouncilTextareaFromRows(next);
                if (typeof ctx.renderStudentCouncilTableFromTextarea === 'function') {
                    ctx.renderStudentCouncilTableFromTextarea();
                }
                if (typeof ctx.renderStudentCouncilPickBoard === 'function') {
                    ctx.renderStudentCouncilPickBoard();
                }
            } finally {
                ctx.sgaStudentCouncilSyncLock = false;
            }
        }

        ctx.reconcileSgaAndStudentCouncil = function reconcileSgaAndStudentCouncil() {
            if (ctx.sgaStudentCouncilSyncLock) return;
            const part = ctx.partitionSgaForBoard();
            const board = part.board || { teachers: [], students: [], externals: [] };
            const council = ctx.getStudentCouncilFromTextarea();
            const source = ctx.resolveStudentRepSyncSource(
                board.students,
                council,
                ctx.sgaStudentSlotCount()
            );
            if (source === 'none') return;
            if (source === 'sga') ctx.pushStudentCouncilFromSgaBoard(board);
            else if (source === 'council') ctx.pushSgaStudentsFromCouncilRows(council);
        }

        ctx.commitSgaBoard = function commitSgaBoard(board, overflow) {
            const fn = window.ms365TenantSettingsSgaRowsFromBoard;
            if (typeof fn !== 'function') return;
            ctx.sgaBoardOverflowRows = Array.isArray(overflow) ? overflow : [];
            ctx.setSgaTextareaFromRows(fn(board, ctx.sgaBoardOverflowRows));
            ctx.renderSgaTableFromTextarea();
            if (!ctx.sgaStudentCouncilSyncLock) {
                ctx.pushStudentCouncilFromSgaBoard(board);
            }
            ctx.scheduleAutoSave();
        }

        ctx.teacherPickLabel = function teacherPickLabel(row) {
            const code = presentValue(row && row.code);
            const name = presentValue(row && row.name);
            const em = presentValue(row && row.email);
            if (code && name) return code + ' — ' + name + (em ? ' (' + em + ')' : '');
            if (name) return name + (em ? ' (' + em + ')' : '');
            return em || '—';
        }

        ctx.studentPickLabel = function studentPickLabel(row) {
            const klasse = presentValue(row && row.klasse);
            const name = presentValue(row && row.name);
            const em = presentValue(row && row.email);
            const head = klasse ? klasse + ' — ' + (name || em) : name || em || '—';
            return em && name ? head + ' (' + em + ')' : head;
        }

        ctx.fillSgaPickSelect = function fillSgaPickSelect(sel, options, selectedEmail, emptyLabel) {
            sel.replaceChildren();
            const opt0 = document.createElement('option');
            opt0.value = '';
            opt0.textContent = emptyLabel || '— aus Listen wählen —';
            sel.appendChild(opt0);
            (options || []).forEach(function (row) {
                const em = ctx.normStr(row && row.email).toLowerCase();
                if (!em || em.indexOf('@') === -1) return;
                const opt = document.createElement('option');
                opt.value = em;
                opt.textContent = row._sgaPickLabel || em;
                if (selectedEmail && em === selectedEmail) opt.selected = true;
                sel.appendChild(opt);
            });
            if (selectedEmail && !sel.value) {
                const extra = document.createElement('option');
                extra.value = selectedEmail;
                extra.textContent = selectedEmail + ' (nicht in Liste)';
                extra.selected = true;
                sel.appendChild(extra);
            }
        }

        ctx.renderSgaBoard = function renderSgaBoard() {
            if (!ctx.sgaBoardTeachersHost && !ctx.sgaBoardStudentsHost && !ctx.sgaBoardExternalsHost) return;
            const part = ctx.partitionSgaForBoard();
            ctx.sgaBoardOverflowRows = part.overflow || [];
            const board = part.board || { teachers: [], students: [], externals: [] };
            const slotN =
                typeof window.ms365TenantSettingsSgaBoardSlots === 'number'
                    ? window.ms365TenantSettingsSgaBoardSlots
                    : 3;
            const teachers = ctx.getTeachersFromTextarea().map(function (t) {
                return Object.assign({}, t, { _sgaPickLabel: ctx.teacherPickLabel(t) });
            });
            const students = ctx.getStudentsFromTextarea().map(function (s) {
                return Object.assign({}, s, { _sgaPickLabel: ctx.studentPickLabel(s) });
            });

            function renderPickColumn(host, scopeKey, scopeValue, slots, pickRows) {
                if (!host) return;
                host.replaceChildren();
                for (let i = 0; i < slotN; i++) {
                    const slot = (slots && slots[i]) || { name: '', email: '' };
                    const em = presentValue(slot.email).toLowerCase();
                    const shownName = presentValue(slot.name);
                    const wrap = document.createElement('div');
                    wrap.className = 'ts-sga-board__slot' + (em || shownName ? ' is-filled' : ' is-empty');

                    const num = document.createElement('span');
                    num.className = 'ts-sga-board__num';
                    num.textContent = String(i + 1);
                    wrap.appendChild(num);

                    const sel = document.createElement('select');
                    sel.className = 'ts-sga-board__pick';
                    sel.title = 'Person aus der Stammdatenliste wählen';
                    ctx.fillSgaPickSelect(sel, pickRows, em, '— aus Listen wählen —');
                    sel.addEventListener('change', function () {
                        const p = ctx.partitionSgaForBoard();
                        const b = p.board;
                        const chosen = ctx.normStr(sel.value).toLowerCase();
                        let name = '';
                        let email = '';
                        if (chosen) {
                            const hit = pickRows.find(function (r) {
                                return ctx.normStr(r.email).toLowerCase() === chosen;
                            });
                            if (hit) {
                                name = ctx.normStr(hit.name) || ctx.normStr(hit.code);
                                email = chosen;
                            } else {
                                email = chosen;
                                name = ctx.normStr(slot.name);
                            }
                        }
                        b[scopeKey][i] = { name: name, email: email };
                        ctx.commitSgaBoard(b, p.overflow);
                        ctx.renderSgaBoard();
                    });
                    wrap.appendChild(sel);

                    const preview = document.createElement('div');
                    preview.className = 'ts-sga-board__preview muted';
                    if (em || shownName) {
                        preview.textContent = shownName ? shownName + (em ? ' · ' + em : '') : em;
                    } else {
                        markEmptyOccupancy(preview);
                    }
                    wrap.appendChild(preview);
                    bindOccupancyEdit(wrap, preview, sel);

                    host.appendChild(wrap);
                }
            }

            function renderExternalColumn(host, slots) {
                if (!host) return;
                host.replaceChildren();
                for (let i = 0; i < slotN; i++) {
                    const slot = (slots && slots[i]) || { name: '', email: '' };
                    const filled = !!(presentValue(slot.name) || presentValue(slot.email));
                    const wrap = document.createElement('div');
                    wrap.className = 'ts-sga-board__slot' + (filled ? ' is-filled' : ' is-empty');

                    const num = document.createElement('span');
                    num.className = 'ts-sga-board__num';
                    num.textContent = String(i + 1);
                    wrap.appendChild(num);

                    const inpName = document.createElement('input');
                    inpName.type = 'text';
                    inpName.className = 'ts-sga-board__name';
                    inpName.placeholder = 'Name';
                    inpName.value = slot.name || '';
                    inpName.autocomplete = 'off';

                    const inpEmail = document.createElement('input');
                    inpEmail.type = 'email';
                    inpEmail.className = 'ts-sga-board__email';
                    inpEmail.placeholder = 'E-Mail';
                    inpEmail.value = slot.email || '';
                    inpEmail.autocomplete = 'off';

                    const commitSlot = function () {
                        const p = ctx.partitionSgaForBoard();
                        const b = p.board;
                        b.externals[i] = {
                            name: ctx.normStr(inpName.value),
                            email: ctx.normStr(inpEmail.value).toLowerCase()
                        };
                        ctx.commitSgaBoard(b, p.overflow);
                    };

                    inpName.addEventListener('change', commitSlot);
                    inpEmail.addEventListener('change', commitSlot);
                    inpName.addEventListener('blur', commitSlot);
                    inpEmail.addEventListener('blur', commitSlot);

                    const preview = document.createElement('div');
                    preview.className = 'ts-sga-board__preview';
                    if (filled) {
                        const extName = presentValue(slot.name);
                        const extMail = presentValue(slot.email);
                        preview.textContent =
                            (extName || 'Ohne Name') + (extMail ? ' · ' + extMail : '');
                    } else {
                        markEmptyOccupancy(preview);
                    }
                    wrap.appendChild(inpName);
                    wrap.appendChild(inpEmail);
                    wrap.appendChild(preview);
                    bindOccupancyEdit(wrap, preview, inpName);
                    host.appendChild(wrap);
                }
            }

            renderPickColumn(ctx.sgaBoardTeachersHost, 'teachers', 'teacher', board.teachers, teachers);
            renderPickColumn(ctx.sgaBoardStudentsHost, 'students', 'student', board.students, students);
            renderExternalColumn(ctx.sgaBoardExternalsHost, board.externals);
        }

        ctx.applyDirectoryRowActionCellStyles = function applyDirectoryRowActionCellStyles(tdAction, btnCheck, btnCreate) {
            tdAction.style.whiteSpace = 'nowrap';
            tdAction.style.display = 'flex';
            tdAction.style.gap = '6px';
            tdAction.style.alignItems = 'center';
            if (btnCheck) btnCheck.style.background = 'var(--brand1)';
            if (btnCreate) btnCreate.style.background = '#11cdef';
        }

        ctx.renderSgaTableFromTextarea = function renderSgaTableFromTextarea() {
            if (!ctx.sgaTbody) return;
            ctx.updateTableSortIndicators(ctx.sgaTable, ctx.sgaSortState);
            const rows = ctx.getSgaFromTextarea();
            ctx.sgaTbody.replaceChildren();
            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 5;
                td.style.color = 'var(--muted)';
                td.textContent = 'Noch keine SGA-Mitglieder – oben einfügen oder „+ Zeile“.';
                tr.appendChild(td);
                ctx.sgaTbody.appendChild(tr);
                return;
            }
            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');
                const tdScope = document.createElement('td');
                const scopeKey = ctx.normStr(row.scope) || 'unknown';
                const badge = document.createElement('span');
                badge.className = 'ts-sga-origin-badge ts-sga-origin-badge--' + (scopeKey || 'unknown');
                badge.textContent = row.scope ? ctx.sgaScopeLabel(row.scope) : 'Sonstige';
                tdScope.appendChild(badge);

                const tdName = document.createElement('td');
                tdName.textContent = row.name || '';
                tdName.title = 'Doppelklick zum Bearbeiten';
                tdName.addEventListener('dblclick', () => {
                    ctx.startCellEdit(tdName, row.name, (next, meta) => {
                        const all = ctx.getSgaFromTextarea();
                        if (!all[idx]) return ctx.renderSgaTableFromTextarea();
                        const prev = all[idx].name;
                        all[idx].name = meta && meta.cancelled ? prev : ctx.normStr(next);
                        ctx.setSgaTextareaFromRows(all);
                        ctx.renderSgaTableFromTextarea();
                        ctx.renderSgaBoard();
                        ctx.scheduleAutoSave();
                    });
                });

                const tdEmail = document.createElement('td');
                tdEmail.textContent = row.email || '';
                tdEmail.title = 'Doppelklick zum Bearbeiten';
                tdEmail.addEventListener('dblclick', () => {
                    ctx.startCellEdit(tdEmail, row.email, (next, meta) => {
                        const all = ctx.getSgaFromTextarea();
                        if (!all[idx]) return ctx.renderSgaTableFromTextarea();
                        const prev = all[idx].email;
                        all[idx].email = meta && meta.cancelled ? prev : ctx.normStr(next).toLowerCase();
                        ctx.setSgaTextareaFromRows(all);
                        ctx.renderSgaTableFromTextarea();
                        ctx.renderSgaBoard();
                        ctx.scheduleAutoSave();
                    });
                });

                const tdMs = ctx.buildDirectoryMatchCell(tr, row.email);

                const tdAction = document.createElement('td');
                tdAction.className = 'action-cell';
                const emRow = ctx.normStr(row.email || '').toLowerCase();
                const dirSga = emRow && emRow.indexOf('@') !== -1 ? ctx.getDirectoryMatchByEmail(emRow) : null;
                const btnCheck = document.createElement('button');
                btnCheck.type = 'button';
                btnCheck.className = 'mini-btn';
                btnCheck.title = 'In Microsoft Entra prüfen';
                btnCheck.innerHTML = '<i class="bi bi-microsoft" aria-hidden="true"></i>';
                if (dirSga && dirSga.graphUserId) {
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
                        ctx.renderSgaTableFromTextarea();
                        ctx.renderSgaBoard();
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
                if (dirSga && dirSga.graphUserId) {
                    btnCreate.disabled = true;
                } else {
                btnCreate.addEventListener('click', () => {
                    if (typeof window.ms365TenantSettingsCreateUser === 'function') {
                        window.ms365TenantSettingsCreateUser(row, () => ctx.renderSgaTableFromTextarea());
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
                    const all = ctx.getSgaFromTextarea();
                    all.splice(idx, 1);
                    ctx.setSgaTextareaFromRows(all);
                    ctx.renderSgaTableFromTextarea();
                    ctx.renderSgaBoard();
                    ctx.scheduleAutoSave();
                });
                tdAction.appendChild(btnDel);
                ctx.applyDirectoryRowActionCellStyles(tdAction, btnCheck, btnCreate);
                collapseActionCell(tdAction, { onEdit: function () { editTableRow(tr); } });

                tr.append(tdScope, tdName, tdEmail, tdMs, tdAction);
                ctx.sgaTbody.appendChild(tr);
            });
        }

}
