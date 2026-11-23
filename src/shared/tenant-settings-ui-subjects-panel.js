/**
 * tenant-settings-ui-subjects-panel.js
 * @param {Record<string, any>} ctx
 */
import { collapseActionCell, editTableRow } from './tenant-workbench-chrome.js';

export function installTenantSubjectsPanel(ctx) {
        ctx.subjectsToLines = function subjectsToLines(rows) {
            return (rows || [])
                .map((x) => `${ctx.normCode(x.code)};${ctx.normStr(x.name || '')}`.trim())
                .filter(Boolean)
                .join('\n');
        }

        ctx.getSubjectsFromTextarea = function getSubjectsFromTextarea() {
            return typeof ctx.parseLinesToSubjects === 'function' ? ctx.parseLinesToSubjects(ctx.taSubjects ? ctx.taSubjects.value : '') : [];
        }

        ctx.setSubjectsTextareaFromRows = function setSubjectsTextareaFromRows(rows) {
            if (!ctx.taSubjects) return;
            ctx.taSubjects.value = ctx.subjectsToLines(rows);
        }

        ctx.renderSubjectsTableFromTextarea = function renderSubjectsTableFromTextarea() {
            if (!ctx.subjectsTbody) return;
            ctx.updateTableSortIndicators(ctx.subjectsTable, ctx.subjectsSortState);
            const rows = ctx.getSubjectsFromTextarea();
            ctx.subjectsTbody.replaceChildren();

            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 4;
                td.style.color = 'var(--muted)';
                td.textContent = 'Noch keine Einträge – Importieren oder „+ Fach“.';
                tr.appendChild(td);
                ctx.subjectsTbody.appendChild(tr);
                return;
            }

            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');

                const tdCode = document.createElement('td');
                tdCode.innerHTML = `<code>${row.code || ''}</code>`;
                tdCode.title = 'Doppelklick zum Bearbeiten';
                tdCode.addEventListener('dblclick', () => {
                    ctx.startCellEdit(tdCode, row.code, (next, meta) => {
                        const all = ctx.getSubjectsFromTextarea();
                        if (!all[idx]) return ctx.renderSubjectsTableFromTextarea();
                        const prev = all[idx].code;
                        const newCode = meta && meta.cancelled ? prev : ctx.normCode(next);
                        if (prev && newCode && ctx.normCode(prev) !== ctx.normCode(newCode)) {
                            const api = window.ms365AppDataV2;
                            if (api && typeof api.renameCatalogLink === 'function') {
                                try {
                                    api.renameCatalogLink('subject', prev, newCode);
                                } catch {
                                    /* Verknüpfung bleibt am alten Kürzel */
                                }
                            }
                        }
                        all[idx].code = newCode;
                        ctx.setSubjectsTextareaFromRows(all);
                        ctx.renderSubjectsTableFromTextarea();
                        ctx.scheduleAutoSave();
                    });
                });

                const tdName = document.createElement('td');
                tdName.textContent = row.name || '';
                tdName.title = 'Doppelklick zum Bearbeiten';
                tdName.addEventListener('dblclick', () => {
                    ctx.startCellEdit(tdName, row.name, (next, meta) => {
                        const all = ctx.getSubjectsFromTextarea();
                        if (!all[idx]) return ctx.renderSubjectsTableFromTextarea();
                        const prev = all[idx].name;
                        all[idx].name = meta && meta.cancelled ? prev : ctx.normStr(next);
                        ctx.setSubjectsTextareaFromRows(all);
                        ctx.renderSubjectsTableFromTextarea();
                        ctx.scheduleAutoSave();
                    });
                });

                const tdAction = document.createElement('td');
                tdAction.className = 'action-cell';
                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'mini-btn';
                btnDel.textContent = '✕';
                btnDel.title = 'Zeile löschen';
                btnDel.addEventListener('click', () => {
                    const all = ctx.getSubjectsFromTextarea();
                    all.splice(idx, 1);
                    ctx.setSubjectsTextareaFromRows(all);
                    ctx.renderSubjectsTableFromTextarea();
                    ctx.scheduleAutoSave();
                });
                tdAction.appendChild(btnDel);
                collapseActionCell(tdAction, { onEdit: function () { editTableRow(tr); } });

                tr.append(tdCode, tdName, ctx.buildSubjectM365Cell(row), tdAction);
                ctx.subjectsTbody.appendChild(tr);
            });
        }

}
