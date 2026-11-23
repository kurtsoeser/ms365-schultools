/**
 * tenant-settings-ui-teachers-table.js
 * @param {Record<string, any>} ctx
 */
import { collapseActionCell, editTableRow, presentText, presentValue } from './tenant-workbench-chrome.js';

export function installTenantTeachersTable(ctx) {
    ctx.getTeachersFromTextarea = function getTeachersFromTextarea() {
        return typeof ctx.parseLinesToTeachers === 'function' ? ctx.parseLinesToTeachers(ctx.taTeachers ? ctx.taTeachers.value : '') : [];
    }

    ctx.setTeachersTextareaFromRows = function setTeachersTextareaFromRows(rows) {
        if (!ctx.taTeachers) return;
        ctx.taTeachers.value = ctx.teachersToLines(rows);
    }

    ctx.sortValueForTenantRow = function sortValueForTenantRow(row, key) {
        if (!row) return '';
        if (key === 'subjects') {
            return (Array.isArray(row.subjects) ? row.subjects : [])
                .map((s) => ctx.normCode(s))
                .filter(Boolean)
                .join(',');
        }
        if (key === 'code' || key === 'year') return ctx.normCode(String(row[key] || ''));
        if (key === 'email' || key === 'headEmail') return ctx.normStr(row[key]).toLowerCase();
        if (key === 'scope') return ctx.normStr(row.scope).toLowerCase();
        if (key === 'klasse') return ctx.normStr(row.klasse).toLowerCase();
        if (key === 'personName') return ctx.normStr(row.personName);
        if (key === 'tier') {
            const t = ctx.inferAdminTierForRole({ name: row.name, code: row.code, tier: row.tier });
            return ctx.adminTierLabel(t);
        }
        return ctx.normStr(row[key]);
    }

    ctx.sortTenantRows = function sortTenantRows(rows, key, dir) {
        return (rows || [])
            .slice()
            .sort((a, b) => {
                const av = ctx.sortValueForTenantRow(a, key);
                const bv = ctx.sortValueForTenantRow(b, key);
                return av.localeCompare(bv, 'de', { sensitivity: 'base', numeric: true }) * dir;
            });
    }

    ctx.updateTableSortIndicators = function updateTableSortIndicators(table, sortState) {
        if (!table) return;
        table.querySelectorAll('th.is-sortable').forEach((th) => {
            const label = th.dataset.label || th.textContent.replace(/\s*[▲▼]\s*$/, '').trim();
            th.dataset.label = label;
            th.textContent = label;
            th.setAttribute('aria-sort', 'none');
            if (sortState.key && th.dataset.sortKey === sortState.key) {
                const ind = document.createElement('span');
                ind.className = 'teachers-sort-ind';
                ind.setAttribute('aria-hidden', 'true');
                ind.textContent = sortState.dir === -1 ? '▼' : '▲';
                th.appendChild(ind);
                th.setAttribute('aria-sort', sortState.dir === -1 ? 'descending' : 'ascending');
            }
        });
    }

    ctx.toggleTenantSortState = function toggleTenantSortState(sortState, key) {
        if (sortState.key === key) sortState.dir = sortState.dir === 1 ? -1 : 1;
        else {
            sortState.key = key;
            sortState.dir = 1;
        }
    }

    ctx.bindTenantTableSort = function bindTenantTableSort(table, sortState, allowedKeys, onSortKey) {
        if (!table || table.dataset.tenantSortBound) return;
        table.dataset.tenantSortBound = '1';
        table.querySelectorAll('th.is-sortable').forEach((th) => {
            const key = th.dataset.sortKey;
            if (!key || (allowedKeys && allowedKeys.indexOf(key) < 0)) return;
            th.addEventListener('click', () => onSortKey(key));
            th.addEventListener('keydown', (e) => {
                if (e.key === 'Enter' || e.key === ' ') {
                    e.preventDefault();
                    onSortKey(key);
                }
            });
        });
    }

    ctx.updateTeachersSortIndicators = function updateTeachersSortIndicators() {
        ctx.updateTableSortIndicators(ctx.teachersTable, ctx.teachersSortState);
    }

    ctx.applyTeachersSort = function applyTeachersSort(key) {
        if (!key || (key !== 'code' && key !== 'name' && key !== 'email')) return;
        if (ctx.teachersSortState.key === key) {
            ctx.teachersSortState.dir = ctx.teachersSortState.dir === 1 ? -1 : 1;
        } else {
            ctx.teachersSortState.key = key;
            ctx.teachersSortState.dir = 1;
        }
        const api = window.ms365TeacherListImport;
        const all = ctx.getTeachersFromTextarea();
        const sorted =
            api && typeof api.sortRows === 'function'
                ? api.sortRows(all, ctx.teachersSortState.key, ctx.teachersSortState.dir)
                : all.slice().sort((a, b) => {
                      const av = String((a && a[key]) || '');
                      const bv = String((b && b[key]) || '');
                      return av.localeCompare(bv, 'de', { sensitivity: 'base' }) * ctx.teachersSortState.dir;
                  });
        ctx.setTeachersTextareaFromRows(sorted);
        ctx.renderTeachersTableFromTextarea();
        ctx.scheduleAutoSave();
    }

    ctx.startCellEdit = function startCellEdit(td, initialValue, onCommit, editOpts) {
        const prevText = String(initialValue ?? '');
        const input = document.createElement('input');
        input.className = 'cell-editor';
        input.type = 'text';
        input.value = prevText;
        td.replaceChildren(input);
        input.focus();
        input.select();

        let finished = false;
        const finish = (value, meta) => {
            if (finished) return;
            finished = true;
            onCommit(value, meta);
        };
        const commit = () => finish(ctx.normStr(input.value));
        const cancel = () => finish(prevText, { cancelled: true });

        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                commit();
            } else if (e.key === 'Escape') {
                e.preventDefault();
                cancel();
            } else if (e.key === 'Tab' && editOpts && editOpts.tabAdvance) {
                e.preventDefault();
                finish(ctx.normStr(input.value), { tab: e.shiftKey ? 'prev' : 'next' });
            }
        });
        input.addEventListener('blur', () => commit());
    }

    /**
     * Lehrkraft ohne E-Mail in Entra suchen und E-Mail (+ Match) in die Zeile schreiben.
     * @param {number} idx
     * @param {string} nameHint
     */
    ctx.pickTeacherFromEntraByName = async function pickTeacherFromEntraByName(idx, nameHint) {
        const hintName = ctx.normStr(nameHint);
        const nameParts = hintName.split(/\s+/).filter(Boolean);
        // Nachname zuerst (Graph: startswith displayName/mail) – oft treffsicherer.
        const searchSeed =
            nameParts.length > 1 ? nameParts[nameParts.length - 1] : hintName;
        let picked;
        try {
            picked = await ctx.pickEntraUser({
                title: 'Lehrkraft in Microsoft 365 suchen',
                hint:
                    'Person ohne gespeicherte E‑Mail zuordnen. Treffer übernimmt E‑Mail und M365‑Abgleich' +
                    (hintName ? ' (Vorschlag: ' + hintName + ')' : '') +
                    '.',
                initialQuery: searchSeed.length >= 2 ? searchSeed : hintName
            });
        } catch (e) {
            ctx.setSummary('M365-Suche: ' + (e && e.message ? e.message : String(e)), 'warn');
            return;
        }
        if (!picked) return;
        const mail = ctx.normStr(picked.mail || '').toLowerCase();
        if (!mail || mail.indexOf('@') === -1) {
            ctx.setSummary('Gewählte Person hat keine E‑Mail/UPN – bitte andere Person wählen.', 'warn');
            return;
        }
        const all = ctx.getTeachersFromTextarea();
        if (!all[idx]) return;
        all[idx].email = mail;
        if (!ctx.normStr(all[idx].name) && picked.displayName) {
            all[idx].name = ctx.normStr(picked.displayName);
        }
        ctx.setTeachersTextareaFromRows(all);
        const iso = new Date().toISOString();
        ctx.patchDirectoryMatchKeys(
            [mail],
            ctx.directoryMatchUserPayload(
                {
                    id: picked.id,
                    displayName: picked.displayName,
                    userPrincipalName: mail
                },
                iso
            )
        );
        ctx.setSummary(
            'E‑Mail aus Microsoft 365 übernommen: ' +
                (picked.displayName || mail) +
                ' → ' +
                mail,
            'ok'
        );
        ctx.scheduleAutoSave();
    }

    ctx.renderTeachersTableFromTextarea = function renderTeachersTableFromTextarea() {
        if (!ctx.teachersTbody) return;
        const rows = ctx.getTeachersFromTextarea();
        ctx.teachersTbody.replaceChildren();
        ctx.updateTeachersSortIndicators();

        if (!rows.length) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 5;
            td.style.color = 'var(--muted)';
            td.textContent = 'Noch keine Einträge – Importieren oder „+ Lehrkraft“.';
            tr.appendChild(td);
            ctx.teachersTbody.appendChild(tr);
            return;
        }

        rows.forEach((row, idx) => {
            const tr = document.createElement('tr');

            const tdCode = document.createElement('td');
            tdCode.innerHTML = `<code>${row.code || ''}</code>`;
            tdCode.title = 'Doppelklick zum Bearbeiten';
            tdCode.addEventListener('dblclick', () => {
                ctx.startCellEdit(tdCode, row.code, (next, meta) => {
                    const all = ctx.getTeachersFromTextarea();
                    if (!all[idx]) return ctx.renderTeachersTableFromTextarea();
                    const prev = all[idx].code;
                    all[idx].code = meta && meta.cancelled ? prev : ctx.normCode(next);
                    ctx.setTeachersTextareaFromRows(all);
                    ctx.renderTeachersTableFromTextarea();
                });
            });

            const tdName = document.createElement('td');
            tdName.textContent = presentText(row.name);
            tdName.title = 'Doppelklick zum Bearbeiten';
            tdName.addEventListener('dblclick', () => {
                ctx.startCellEdit(tdName, presentValue(row.name), (next, meta) => {
                    const all = ctx.getTeachersFromTextarea();
                    if (!all[idx]) return ctx.renderTeachersTableFromTextarea();
                    const prev = all[idx].name;
                    all[idx].name = meta && meta.cancelled ? prev : ctx.normStr(next);
                    ctx.setTeachersTextareaFromRows(all);
                    ctx.renderTeachersTableFromTextarea();
                });
            });

            const tdEmail = document.createElement('td');
            tdEmail.textContent = presentText(row.email);
            tdEmail.title = 'Doppelklick zum Bearbeiten';
            tdEmail.addEventListener('dblclick', () => {
                ctx.startCellEdit(tdEmail, presentValue(row.email), (next, meta) => {
                    const all = ctx.getTeachersFromTextarea();
                    if (!all[idx]) return ctx.renderTeachersTableFromTextarea();
                    const prev = all[idx].email;
                    all[idx].email = meta && meta.cancelled ? prev : ctx.normStr(next).toLowerCase();
                    ctx.setTeachersTextareaFromRows(all);
                    ctx.renderTeachersTableFromTextarea();
                });
            });

            const tdMs = ctx.buildDirectoryMatchCell(tr, row.email);

            const tdAction = document.createElement('td');
            tdAction.className = 'action-cell';
            tdAction.style.whiteSpace = 'nowrap';
            tdAction.style.display = 'flex';
            tdAction.style.gap = '6px';
            tdAction.style.alignItems = 'center';
            const dir = ctx.getDirectoryMatchByEmail(row.email);
            const hasEmail = !!(row.email && String(row.email).includes('@'));
            // Ohne E-Mail: in M365 nach Name suchen und Adresse übernehmen.
            if (!hasEmail && row.name) {
                const btnSearch = document.createElement('button');
                btnSearch.type = 'button';
                btnSearch.className = 'mini-btn';
                btnSearch.style.background = 'var(--brand1)';
                btnSearch.title = 'Person in Microsoft 365 suchen und E‑Mail übernehmen';
                btnSearch.innerHTML = '<i class="bi bi-search" aria-hidden="true"></i>';
                btnSearch.addEventListener('click', async () => {
                    btnSearch.disabled = true;
                    try {
                        await ctx.pickTeacherFromEntraByName(idx, row.name);
                    } finally {
                        ctx.renderTeachersTableFromTextarea();
                    }
                });
                tdAction.appendChild(btnSearch);
            }
            if (hasEmail && row.name) {
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
                            ctx.renderTeachersTableFromTextarea();
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
                            ctx.renderTeachersTableFromTextarea();
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
                const all = ctx.getTeachersFromTextarea();
                all.splice(idx, 1);
                ctx.setTeachersTextareaFromRows(all);
                ctx.renderTeachersTableFromTextarea();
            });
            tdAction.appendChild(btnDel);
            collapseActionCell(tdAction, { onEdit: function () { editTableRow(tr); } });

            tr.append(tdCode, tdName, tdEmail, tdMs, tdAction);
            ctx.teachersTbody.appendChild(tr);
        });
    }

}
