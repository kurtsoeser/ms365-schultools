/**
 * tenant-settings-ui-admin.js
 * @param {Record<string, any>} ctx
 */
export function installTenantAdminPanel(ctx) {
        ctx.renderAdminRolesTableFromTextarea = function renderAdminRolesTableFromTextarea() {
            if (!ctx.adminRolesTbody) return;
            const rows = ctx.getAdminRolesFromTextarea();
            ctx.adminRolesTbody.replaceChildren();

            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 3;
                td.style.color = 'var(--muted)';
                td.textContent = 'Noch keine Rollen – oben einfügen, „+ Rolle“ oder Standardrollen.';
                tr.appendChild(td);
                ctx.adminRolesTbody.appendChild(tr);
                return;
            }

            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');

                const tdCode = document.createElement('td');
                tdCode.innerHTML = `<code>${row.code || ''}</code>`;
                tdCode.title = 'Doppelklick zum Bearbeiten';
                tdCode.addEventListener('dblclick', () => {
                    ctx.startCellEdit(tdCode, row.code, (next, meta) => {
                        const all = ctx.getAdminRolesFromTextarea();
                        if (!all[idx]) return ctx.renderAdminRolesTableFromTextarea();
                        const prev = all[idx].code;
                        all[idx].code = meta && meta.cancelled ? prev : ctx.normCode(next);
                        ctx.setAdminRolesTextareaFromRows(all);
                        ctx.renderAdminRolesTableFromTextarea();
                        ctx.renderAdminTableFromTextarea();
                        ctx.scheduleAutoSave();
                    });
                });

                const tdName = document.createElement('td');
                tdName.textContent = row.name || '';
                tdName.title = 'Doppelklick zum Umbenennen (Personen werden mitgezogen)';
                tdName.addEventListener('dblclick', () => {
                    ctx.startCellEdit(tdName, row.name, (next, meta) => {
                        const all = ctx.getAdminRolesFromTextarea();
                        if (!all[idx]) return ctx.renderAdminRolesTableFromTextarea();
                        const prev = all[idx].name;
                        const nextName = meta && meta.cancelled ? prev : ctx.normStr(next);
                        if (nextName && nextName !== prev && typeof window.ms365TenantSettingsRenameAdminRole === 'function') {
                            const renamed = window.ms365TenantSettingsRenameAdminRole(all, ctx.getAdminFromTextarea(), prev, nextName);
                            ctx.setAdminRolesTextareaFromRows(renamed.roles);
                            ctx.setAdminTextareaFromRows(renamed.admin);
                        } else {
                            all[idx].name = nextName;
                            ctx.setAdminRolesTextareaFromRows(all);
                        }
                        ctx.renderAdminRolesTableFromTextarea();
                        ctx.renderAdminTableFromTextarea();
                        ctx.scheduleAutoSave();
                    });
                });

                const tdAction = document.createElement('td');
                tdAction.className = 'action-cell';
                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'mini-btn';
                btnDel.textContent = '✕';
                btnDel.title = 'Rolle löschen';
                btnDel.addEventListener('click', () => {
                    const people = ctx.getAdminFromTextarea().filter((p) => {
                        if (typeof window.ms365TenantSettingsPersonMatchesAdminRole === 'function') {
                            return window.ms365TenantSettingsPersonMatchesAdminRole(p, row);
                        }
                        return ctx.normStr(p.role).toLowerCase() === ctx.normStr(row.name).toLowerCase();
                    });
                    if (people.length) {
                        window.alert(
                            'Diese Rolle ist noch ' +
                                people.length +
                                ' Person(en) zugeordnet. Bitte zuerst umbenennen oder die Personen entfernen.'
                        );
                        return;
                    }
                    const all = ctx.getAdminRolesFromTextarea();
                    all.splice(idx, 1);
                    ctx.setAdminRolesTextareaFromRows(all);
                    ctx.renderAdminRolesTableFromTextarea();
                    ctx.renderAdminTableFromTextarea();
                    ctx.scheduleAutoSave();
                });
                tdAction.appendChild(btnDel);

                tr.append(tdCode, tdName, tdAction);
                ctx.adminRolesTbody.appendChild(tr);
            });
        }

        ctx.renderAdminTableFromTextarea = function renderAdminTableFromTextarea() {
            if (!ctx.adminTbody) return;
            const rows = ctx.getAdminFromTextarea();
            const roleCatalog = ctx.getAdminRolesFromTextarea();
            ctx.adminTbody.replaceChildren();

            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 4;
                td.style.color = 'var(--muted)';
                td.textContent = 'Noch keine Einträge – oben einfügen oder „+ Person“.';
                tr.appendChild(td);
                ctx.adminTbody.appendChild(tr);
                return;
            }

            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');

                const tdRole = document.createElement('td');
                const sel = document.createElement('select');
                sel.style.width = '100%';
                sel.style.font = 'inherit';
                const names = [];
                const seenN = new Set();
                roleCatalog.forEach((r) => {
                    const n = ctx.normStr(r.name || r.code);
                    if (!n) return;
                    const k = n.toLowerCase();
                    if (seenN.has(k)) return;
                    seenN.add(k);
                    names.push(n);
                });
                const current = ctx.normStr(row.role);
                if (current && !seenN.has(current.toLowerCase())) names.unshift(current);
                const optEmpty = document.createElement('option');
                optEmpty.value = '';
                optEmpty.textContent = '— Rolle —';
                sel.appendChild(optEmpty);
                names.forEach((n) => {
                    const opt = document.createElement('option');
                    opt.value = n;
                    opt.textContent = n;
                    if (current && n.toLowerCase() === current.toLowerCase()) opt.selected = true;
                    sel.appendChild(opt);
                });
                sel.addEventListener('change', () => {
                    const all = ctx.getAdminFromTextarea();
                    if (!all[idx]) return ctx.renderAdminTableFromTextarea();
                    all[idx].role = ctx.normStr(sel.value);
                    ctx.setAdminTextareaFromRows(all);
                    ctx.renderAdminTableFromTextarea();
                    ctx.scheduleAutoSave();
                });
                tdRole.appendChild(sel);

                const tdName = document.createElement('td');
                tdName.textContent = row.name || '';
                tdName.title = 'Doppelklick zum Bearbeiten';
                tdName.addEventListener('dblclick', () => {
                    ctx.startCellEdit(tdName, row.name, (next, meta) => {
                        const all = ctx.getAdminFromTextarea();
                        if (!all[idx]) return ctx.renderAdminTableFromTextarea();
                        const prev = all[idx].name;
                        all[idx].name = meta && meta.cancelled ? prev : ctx.normStr(next);
                        ctx.setAdminTextareaFromRows(all);
                        ctx.renderAdminTableFromTextarea();
                        ctx.scheduleAutoSave();
                    });
                });

                const tdEmail = document.createElement('td');
                tdEmail.textContent = row.email || '';
                tdEmail.title = 'Doppelklick zum Bearbeiten';
                tdEmail.addEventListener('dblclick', () => {
                    ctx.startCellEdit(tdEmail, row.email, (next, meta) => {
                        const all = ctx.getAdminFromTextarea();
                        if (!all[idx]) return ctx.renderAdminTableFromTextarea();
                        const prev = all[idx].email;
                        all[idx].email = meta && meta.cancelled ? prev : ctx.normStr(next).toLowerCase();
                        ctx.setAdminTextareaFromRows(all);
                        ctx.renderAdminTableFromTextarea();
                        ctx.scheduleAutoSave();
                    });
                });

                const tdAction = document.createElement('td');
                tdAction.className = 'action-cell';
                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'mini-btn';
                btnDel.textContent = '✕';
                btnDel.title = 'Person löschen';
                btnDel.addEventListener('click', () => {
                    const all = ctx.getAdminFromTextarea();
                    all.splice(idx, 1);
                    ctx.setAdminTextareaFromRows(all);
                    ctx.renderAdminTableFromTextarea();
                    ctx.scheduleAutoSave();
                });
                tdAction.appendChild(btnDel);

                tr.append(tdRole, tdName, tdEmail, tdAction);
                ctx.adminTbody.appendChild(tr);
            });
        }

}
