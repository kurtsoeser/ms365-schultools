/**
 * tenant-settings-ui-bind-events-inventory.js
 * @param {Record<string, any>} ctx
 */
export function installTenantBindEventsInventory(ctx) {
    ctx.setTenantDeltaProgress = function setTenantDeltaProgress(visible, text, ratio) {
        const wrap = document.getElementById('tenantDeltaProgressWrap');
        const txt = document.getElementById('tenantDeltaProgressText');
        const bar = document.getElementById('tenantDeltaProgressBar');
        const pct = document.getElementById('tenantDeltaProgressPct');
        if (wrap) wrap.style.display = visible ? '' : 'none';
        if (txt && text) txt.textContent = String(text);
        const r = typeof ratio === 'number' && isFinite(ratio) ? Math.max(0, Math.min(1, ratio)) : null;
        if (bar) bar.style.width = r === null ? '0%' : String(Math.round(r * 100)) + '%';
        if (pct) pct.textContent = r === null ? '–' : String(Math.round(r * 100)) + ' %';
    }

    ctx.collectTenantDeltaItems = function collectTenantDeltaItems(api) {
        const out = [];
        if (!api || typeof api.loadStructureState !== 'function') return out;
        const { rows } = api.loadStructureState();
        const links = api.loadMatchLinks() || {};
        (rows || []).forEach((r) => {
            if (!r) return;
            const t = String(r.typ || '');
            const gid =
                ctx.normStr(r.tenantGroupId) ||
                (links[String(r.id)] && ctx.normStr(links[String(r.id)].tenantGroupId));
            const groupLike =
                t === 'Gruppe' ||
                t === 'Arbeitsgemeinschaft' ||
                t === 'Klasse' ||
                t === 'Jahrgang' ||
                t === 'SchuelerInnen' ||
                t === 'LehrerInnen' ||
                t === 'Verwaltung';
            if (groupLike) {
                if (!gid) {
                    const sug = typeof api.computeCreateSuggestion === 'function' ? api.computeCreateSuggestion(r) : null;
                    const can = !!(sug && sug.displayName && sug.mailNick);
                    let action = can
                        ? 'M365‑Gruppe/Team per Graph anlegen (Mail‑Nickname vorhanden).'
                        : 'Schema/Bezeichnung prüfen (Mail‑Nickname leer).';
                    if (t === 'Klasse' && can) {
                        action = 'Klassen‑Team/Gruppe per Graph anlegen (Schema: meist Microsoft‑365‑Gruppe).';
                    } else if (t === 'Jahrgang' && can) {
                        action = 'Jahrgangs‑Gruppe per Graph anlegen (Schema: Jahrgang/Jg‑Suffix oder Anzeigename).';
                    } else if (t === 'Arbeitsgemeinschaft' && can) {
                        action = 'ARGE als Microsoft‑365‑Gruppe per Graph anlegen.';
                    }
                    out.push({
                        kind: 'group',
                        r,
                        action,
                        canProvision: can
                    });
                }
                return;
            }
            if (t === 'Kursteam') {
                if (gid) return;
                const kt =
                    window.ms365StructureRules &&
                    typeof window.ms365StructureRules.resolveKursteamKlasseFach === 'function'
                        ? window.ms365StructureRules.resolveKursteamKlasseFach(r, rows)
                        : {
                              klasse: ctx.normStr(r.ktKlasse),
                              fach: ctx.normStr(r.ktFach),
                              hasBoth: !!(ctx.normStr(r.ktKlasse) && ctx.normStr(r.ktFach))
                          };
                if (!kt.hasBoth) {
                    out.push({
                        kind: 'group',
                        r,
                        action:
                            'Kursteam: Klasse und Fach fehlen (Feld „Klasse“/„Fach“ oder Kursteam unter einer „Klasse“-Zeile mit gesetztem Fach).',
                        canProvision: false
                    });
                    return;
                }
                const sug = typeof api.computeCreateSuggestion === 'function' ? api.computeCreateSuggestion(r) : null;
                const can = !!(sug && sug.displayName && sug.mailNick);
                out.push({
                    kind: 'group',
                    r,
                    action: can ? 'Kursteam (Unified + Team) per Graph anlegen.' : 'Kursteam: Mail‑Nickname/Schema prüfen.',
                    canProvision: can
                });
                return;
            }
            const linkRow = links[String(r.id)] || null;
            const linkedUserForPerson = ctx.normStr(r.tenantUserId) || (linkRow && ctx.normStr(linkRow.tenantUserId));
            if (t === 'Person' && !linkedUserForPerson) {
                const dn = ctx.normStr(r.personName) || ctx.normStr(r.bezeichnung);
                const em = ctx.normStr(r.personEmail).toLowerCase();
                const can = !!(dn && em && em.indexOf('@') !== -1);
                out.push({
                    kind: 'person',
                    r,
                    action: can
                        ? 'Entra‑Benutzer per Graph anlegen (Name + E‑Mail vorhanden).'
                        : 'Person: Name und gültige E‑Mail/UPN in Schritt 3 ergänzen.',
                    canProvision: can
                });
            }
        });
        return out;
    }

    ctx.renderTenantInventoryMatchRows = function renderTenantInventoryMatchRows() {
        const tbody = document.getElementById('tenantInvMatchTbody');
        const sum = document.getElementById('tenantInvSummary');
        if (!tbody) return;
        const api = window.ms365TenantInventory;
        if (!api || typeof api.loadStructureState !== 'function') {
            tbody.innerHTML =
                '<tr><td colspan="3" class="muted">Technik wird geladen … Seite ggf. neu laden, falls diese Zeile bleibt.</td></tr>';
            return;
        }
        const { rows } = api.loadStructureState();
        const links = api.loadMatchLinks() || {};
        const cache = typeof api.readCache === 'function' ? api.readCache() : { rows: [], users: [] };
        const groups = cache.rows || [];
        const users = cache.users || [];
        const orgRows = (rows || []).filter((r) => {
            if (!r) return false;
            const t = String(r.typ || '');
            return (
                t === 'SchuelerInnen' ||
                t === 'LehrerInnen' ||
                t === 'Verwaltung' ||
                t === 'Jahrgang' ||
                t === 'Klasse' ||
                t === 'Gruppe' ||
                t === 'Kursteam' ||
                t === 'Person' ||
                t === 'Arbeitsgemeinschaft'
            );
        });
        if (!orgRows.length) {
            tbody.innerHTML =
                '<tr><td colspan="3" class="muted">Keine Einträge vom Typ Jahrgang, Klasse, Gruppe, Kursteam, ARGE (Arbeitsgemeinschaft) oder Person in der SOLL‑Struktur (Schritt 3).</td></tr>';
            if (sum) {
                sum.style.display = '';
                sum.textContent =
                    'Cache: ' + groups.length + ' Gruppe(n)/Team(s), ' + users.length + ' Benutzerkonto/-konten (nach „Tenant laden“).';
            }
            return;
        }
        tbody.replaceChildren();
        orgRows.forEach((r) => {
            const tr = document.createElement('tr');
            const rid = String(r.id || '');
            const linked =
                ctx.normStr(r.tenantGroupId) ||
                (links[rid] && ctx.normStr(links[rid].tenantGroupId)) ||
                '';
            const c1 = document.createElement('td');
            c1.textContent = (r.bezeichnung || '–') + ' · ' + (r.typ || '');
            const c2 = document.createElement('td');
            const c3 = document.createElement('td');
            c3.className = 'tenant-inv-actions';
            const t = String(r.typ || '');
            if (t === 'Person') {
                const pNote = document.createElement('div');
                pNote.className = 'muted';
                pNote.style.fontSize = '0.82em';
                pNote.style.lineHeight = '1.35';
                const linkP = links[rid] || null;
                const uidShow = ctx.normStr(r.tenantUserId) || (linkP && ctx.normStr(linkP.tenantUserId));
                let label = '';
                if (uidShow) {
                    const hit = users.find(function (u) {
                        return u && String(u.id) === String(uidShow);
                    });
                    label = hit
                        ? (ctx.normStr(hit.displayName) || ctx.normStr(hit.userPrincipalName) || uidShow) + ' · ' + uidShow
                        : uidShow;
                }
                pNote.textContent = uidShow ? 'Entra: ' + label : '— · Benutzer (Abgleich / Schritt 5)';
                c2.appendChild(pNote);
                c3.textContent = '';
            } else {
                const sel = document.createElement('select');
                sel.setAttribute('data-tenant-inv-sel', rid);
                const o0 = document.createElement('option');
                o0.value = '';
                o0.textContent = '(keine Entra‑Gruppe)';
                sel.appendChild(o0);
                groups.forEach((g) => {
                    const o = document.createElement('option');
                    o.value = String(g.id || '');
                    o.textContent = (g.bezeichnung || '(ohne Name)') + ' · ' + (g.typ || '') + (g.alias ? ' · ' + g.alias : '');
                    sel.appendChild(o);
                });
                sel.value = linked || '';
                c2.appendChild(sel);
                const bSug = document.createElement('button');
                bSug.type = 'button';
                bSug.className = 'btn small-btn tenant-inv-icon-btn';
                bSug.setAttribute('data-tenant-inv-suggest', rid);
                bSug.setAttribute('aria-label', 'Vorschlag');
                bSug.title = 'Vorschlag ins Dropdown (ohne Speichern)';
                bSug.innerHTML = '<i class="bi bi-magic"></i>';
                const bSave = document.createElement('button');
                bSave.type = 'button';
                bSave.className = 'btn btn-success small-btn tenant-inv-icon-btn';
                bSave.setAttribute('data-tenant-inv-save', rid);
                bSave.setAttribute('aria-label', 'Speichern');
                bSave.title = 'Verknüpfung speichern';
                bSave.innerHTML = '<i class="bi bi-check2"></i>';
                const bClr = document.createElement('button');
                bClr.type = 'button';
                bClr.className = 'btn small-btn tenant-inv-icon-btn';
                bClr.setAttribute('data-tenant-inv-clear', rid);
                bClr.setAttribute('aria-label', 'Verknüpfung löschen');
                bClr.title = 'Verknüpfung löschen';
                bClr.innerHTML = '<i class="bi bi-x-lg"></i>';
                c3.appendChild(bSug);
                c3.appendChild(bSave);
                c3.appendChild(bClr);
            }
            tr.appendChild(c1);
            tr.appendChild(c2);
            tr.appendChild(c3);
            tbody.appendChild(tr);
        });
        if (sum) {
            sum.style.display = '';
            sum.textContent =
                'Cache: ' +
                groups.length +
                ' Gruppe(n)/Team(s), ' +
                users.length +
                ' Benutzer · ' +
                orgRows.length +
                ' Match‑Zeilen.';
        }
    }

    ctx.renderTenantDeltaRows = function renderTenantDeltaRows() {
        const tbody = document.getElementById('tenantDeltaTbody');
        const sum = document.getElementById('tenantDeltaSummary');
        if (!tbody) return;
        const api = window.ms365TenantInventory;
        if (!api || typeof api.loadStructureState !== 'function') {
            tbody.innerHTML = '<tr><td colspan="3" class="muted">–</td></tr>';
            return;
        }
        const items = ctx.collectTenantDeltaItems(api);
        tbody.replaceChildren();
        if (!items.length) {
            tbody.innerHTML =
                '<tr><td colspan="3" class="muted">Keine offenen Differenzen (Kursteams nur mit gültigem Klasse‑/Fach‑Kontext; Personen nur mit Name und E‑Mail).</td></tr>';
        } else {
            items.forEach((it) => {
                const tr = document.createElement('tr');
                const a = document.createElement('td');
                a.textContent = it.r.bezeichnung || '–';
                const b = document.createElement('td');
                b.textContent = it.r.typ || '';
                const c = document.createElement('td');
                c.style.display = 'flex';
                c.style.flexWrap = 'wrap';
                c.style.gap = '8px';
                c.style.alignItems = 'center';
                const span = document.createElement('span');
                span.textContent = it.action;
                span.style.flex = '1';
                span.style.minWidth = '120px';
                c.appendChild(span);
                if (it.canProvision && typeof api.provisionGroupRow === 'function' && it.kind === 'group') {
                    const one = document.createElement('button');
                    one.type = 'button';
                    one.className = 'btn btn-success small-btn';
                    one.setAttribute('data-tenant-delta-one-group', String(it.r.id));
                    one.textContent = 'Jetzt anlegen';
                    c.appendChild(one);
                }
                if (it.canProvision && typeof api.provisionPersonRow === 'function' && it.kind === 'person') {
                    const one = document.createElement('button');
                    one.type = 'button';
                    one.className = 'btn btn-success small-btn';
                    one.setAttribute('data-tenant-delta-one-person', String(it.r.id));
                    one.textContent = 'Jetzt anlegen';
                    c.appendChild(one);
                }
                tr.appendChild(a);
                tr.appendChild(b);
                tr.appendChild(c);
                tbody.appendChild(tr);
            });
        }
        if (sum) {
            sum.style.display = '';
            const prov = items.filter((x) => x.canProvision).length;
            sum.textContent =
                items.length +
                ' offene Position(en), davon ' +
                prov +
                ' mit Graph‑Anlage möglich (laut Schema).';
        }
    }

    window.__ms365TenantStepsRefreshMatch = function () {
        ctx.renderTenantInventoryMatchRows();
        ctx.renderTenantDeltaRows();
    };

    ctx.invTbody = document.getElementById('tenantInvMatchTbody');
    if (ctx.invTbody && !ctx.invTbody.dataset.tenantInvBound) {
        ctx.invTbody.dataset.tenantInvBound = '1';
        ctx.invTbody.addEventListener('click', (ev) => {
            const api = window.ms365TenantInventory;
            if (!api || typeof api.saveMatchLink !== 'function') return;
            const t = ev.target;
            const saveB = t.closest && t.closest('[data-tenant-inv-save]');
            const sugB = t.closest && t.closest('[data-tenant-inv-suggest]');
            const clrB = t.closest && t.closest('[data-tenant-inv-clear]');
            if (sugB) {
                const rid = sugB.getAttribute('data-tenant-inv-suggest');
                const row = (api.loadStructureState().rows || []).find((x) => String(x.id) === String(rid));
                if (!row) return;
                const gid = typeof api.suggestGroupForUnit === 'function' ? api.suggestGroupForUnit(row) : '';
                const sel = ctx.invTbody.querySelector('select[data-tenant-inv-sel="' + rid + '"]');
                if (sel && gid) sel.value = gid;
                return;
            }
            if (clrB) {
                const rid = clrB.getAttribute('data-tenant-inv-clear');
                api.saveMatchLink(rid, '', '');
                try {
                    if (typeof api.patchStructureRow === 'function') {
                        api.patchStructureRow(rid, { tenantGroupId: '', tenantMailNickname: '', syncStatus: 'Ausstehend' });
                    }
                } catch {
                    // ignore
                }
                window.__ms365TenantStepsRefreshMatch();
                return;
            }
            if (saveB) {
                const rid = saveB.getAttribute('data-tenant-inv-save');
                const sel = ctx.invTbody.querySelector('select[data-tenant-inv-sel="' + rid + '"]');
                const gid = sel && sel.value ? String(sel.value).trim() : '';
                api.saveMatchLink(rid, gid, '');
                if (typeof api.patchStructureRow === 'function') {
                    if (gid) {
                        api.patchStructureRow(rid, {
                            tenantGroupId: gid,
                            syncStatus: 'Ok',
                            letzteFehlermeldung: ''
                        });
                    } else {
                        api.patchStructureRow(rid, {
                            tenantGroupId: '',
                            tenantMailNickname: '',
                            syncStatus: 'Ausstehend',
                            letzteFehlermeldung: ''
                        });
                    }
                }
                window.__ms365TenantStepsRefreshMatch();
            }
        });
    }

    ctx.deltaTbody = document.getElementById('tenantDeltaTbody');
    if (ctx.deltaTbody && !ctx.deltaTbody.dataset.tenantDeltaBound) {
        ctx.deltaTbody.dataset.tenantDeltaBound = '1';
        ctx.deltaTbody.addEventListener('click', async (ev) => {
            const api = window.ms365TenantInventory;
            if (!api) return;
            const gBtn = ev.target.closest && ev.target.closest('[data-tenant-delta-one-group]');
            const pBtn = ev.target.closest && ev.target.closest('[data-tenant-delta-one-person]');
            const id = (gBtn && gBtn.getAttribute('data-tenant-delta-one-group')) || (pBtn && pBtn.getAttribute('data-tenant-delta-one-person'));
            if (!id) return;
            const row = (api.loadStructureState().rows || []).find((x) => String(x.id) === String(id));
            if (!row) return;
            try {
                ctx.setTenantDeltaProgress(true, 'Anlegen …', 0.2);
                if (gBtn) await api.provisionGroupRow(row);
                if (pBtn) {
                    const res = await api.provisionPersonRow(row, {});
                    if (res && res.tempPassword) {
                        await ctx.dlgAlert('Benutzer angelegt. Einmaliges Kennwort:\n\n' + res.tempPassword, {
                            title: 'Kennwort notieren',
                            okText: 'Verstanden'
                        });
                    }
                }
            } catch (e) {
                await ctx.dlgAlert('Fehler: ' + (e && e.message ? e.message : String(e)), { title: 'Fehler' });
            } finally {
                ctx.setTenantDeltaProgress(false, '', null);
                window.__ms365TenantStepsRefreshMatch();
            }
        });
    }

    ctx.runTenantInventoryAutomatch = async function runTenantInventoryAutomatch() {
        const api = window.ms365TenantInventory;
        const elSt = document.getElementById('tenantInvStatus');
        if (!api || typeof api.loadStructureState !== 'function' || typeof api.suggestGroupForUnit !== 'function') {
            if (elSt) elSt.textContent = 'Automatching: Technik nicht bereit.';
            return;
        }
        const cache = typeof api.readCache === 'function' ? api.readCache() : { rows: [] };
        const gr = cache.rows || [];
        if (!gr.length) {
            await ctx.dlgAlert('Zuerst „Tenant laden“ ausführen, damit Entra-Gruppen für das Automatching vorliegen.', {
                title: 'Automatching'
            });
            return;
        }
        const gidSet = new Set(gr.map((g) => String(g.id || '').trim()).filter(Boolean));
        const st0 = api.loadStructureState();
        const rows = st0.rows || [];
        let links = api.loadMatchLinks() || {};
        const orgRows = rows.filter((r) => {
            if (!r || r.isStructureTreeRoot) return false;
            const ty = String(r.typ || '');
            return (
                ty === 'Jahrgang' ||
                ty === 'Klasse' ||
                ty === 'Gruppe' ||
                ty === 'Kursteam' ||
                ty === 'Arbeitsgemeinschaft'
            );
        });
        let saved = 0;
        let skippedHas = 0;
        let skippedNoHit = 0;
        for (let i = 0; i < orgRows.length; i++) {
            const r = orgRows[i];
            const rid = String(r.id || '');
            const linked =
                ctx.normStr(r.tenantGroupId) ||
                (links[rid] && ctx.normStr(links[rid].tenantGroupId));
            if (linked) {
                skippedHas++;
                continue;
            }
            const gid = String(api.suggestGroupForUnit(r) || '').trim();
            if (!gid || !gidSet.has(gid)) {
                skippedNoHit++;
                continue;
            }
            try {
                api.saveMatchLink(rid, gid, 'Automatching');
                if (typeof api.patchStructureRow === 'function') {
                    api.patchStructureRow(rid, {
                        tenantGroupId: gid,
                        syncStatus: 'Ok',
                        letzteFehlermeldung: ''
                    });
                }
                saved++;
                links = api.loadMatchLinks() || links;
                r.tenantGroupId = gid;
            } catch {
                skippedNoHit++;
            }
        }
        window.__ms365TenantStepsRefreshMatch();
        const parts = ['Automatching: ' + saved + ' neu verknüpft.'];
        if (skippedHas) parts.push(skippedHas + ' bereits gesetzt.');
        if (skippedNoHit) parts.push(skippedNoHit + ' ohne Treffer.');
        if (elSt) elSt.textContent = parts.join(' ');
    }

    ctx.invAuto = document.getElementById('tenantInvAutoMatchBtn');
    if (ctx.invAuto && !ctx.invAuto.dataset.tenantInvAutoBound) {
        ctx.invAuto.dataset.tenantInvAutoBound = '1';
        ctx.invAuto.addEventListener('click', () => void ctx.runTenantInventoryAutomatch());
    }

    ctx.invBtn = document.getElementById('tenantInvRefreshBtn');
    if (ctx.invBtn) {
        ctx.invBtn.addEventListener('click', async () => {
            const st = document.getElementById('tenantInvStatus');
            const api = window.ms365TenantInventory;
            if (!api || typeof api.refresh !== 'function') {
                if (st) st.textContent = 'Schulstruktur‑Modul nicht geladen. Seite neu laden.';
                return;
            }
            ctx.invBtn.disabled = true;
            if (st) st.textContent = 'Lade Daten über Microsoft Graph …';
            try {
                await api.refresh((ev) => {
                    if (!st) return;
                    const ph = ev && ev.phase === 'users' ? 'Benutzer' : 'Gruppen/Teams';
                    const pg = ev && ev.page != null ? ' (Seite ' + ev.page + ')' : '';
                    st.textContent = ph + ' werden geladen' + pg + ' …';
                });
                const c = api.readCache();
                if (st) {
                    st.textContent =
                        'Fertig: ' +
                        (c.rows || []).length +
                        ' Gruppe(n)/Team(s), ' +
                        (c.users || []).length +
                        ' Benutzerkonto/-konten.';
                }
                ctx.renderTenantInventoryMatchRows();
                ctx.renderTenantDeltaRows();
            } catch (e) {
                if (st) st.textContent = 'Fehler: ' + (e && e.message ? e.message : String(e));
            } finally {
                ctx.invBtn.disabled = false;
            }
        });
    }
    ctx.deltaBtn = document.getElementById('tenantDeltaRefreshBtn');
    if (ctx.deltaBtn) {
        ctx.deltaBtn.addEventListener('click', () => ctx.renderTenantDeltaRows());
    }

    ctx.runBatchProvision = async function runBatchProvision(kind) {
        const api = window.ms365TenantInventory;
        if (!api) return;
        const items = ctx.collectTenantDeltaItems(api).filter((x) => x.canProvision && x.kind === kind);
        if (!items.length) {
            await ctx.dlgAlert(
                kind === 'group'
                    ? 'Keine anlegbaren Gruppen (Jahrgang, Klasse, Gruppe, Kursteam, ARGE – Schema prüfen).'
                    : 'Keine anlegbaren Personen (Name + E‑Mail).',
                { title: 'Delta-Anlage' }
            );
            return;
        }
        const n = items.length;
        const ok = await ctx.dlgConfirm(
            kind === 'group'
                ? n + ' Gruppe(n)/Team(s) (inkl. Klasse, Jahrgang, ARGE, Kursteam) wirklich in Entra anlegen?'
                : n + ' Benutzerkonto/-konten wirklich in Entra anlegen?',
            { title: 'Entra-Anlage', okText: 'Anlegen', danger: true }
        );
        if (!ok) return;
        const btnG = document.getElementById('tenantDeltaProvisionGroupsBtn');
        const btnP = document.getElementById('tenantDeltaProvisionPersonsBtn');
        if (btnG) btnG.disabled = true;
        if (btnP) btnP.disabled = true;
        let pwdNotes = [];
        for (let i = 0; i < items.length; i++) {
            const it = items[i];
            const ratio = (i + 0.35) / n;
            ctx.setTenantDeltaProgress(true, 'Anlegen ' + (i + 1) + ' / ' + n + ' …', ratio);
            try {
                if (kind === 'group') await api.provisionGroupRow(it.r);
                else {
                    const res = await api.provisionPersonRow(it.r, { skipConfirm: true });
                    if (res && res.tempPassword) pwdNotes.push((it.r.bezeichnung || it.r.personName || '') + ': ' + res.tempPassword);
                }
            } catch (e) {
                await ctx.dlgAlert('Abbruch bei Position ' + (i + 1) + ': ' + (e && e.message ? e.message : String(e)), {
                    title: 'Fehler'
                });
                break;
            }
        }
        if (btnG) btnG.disabled = false;
        if (btnP) btnP.disabled = false;
        ctx.setTenantDeltaProgress(true, 'Fertig.', 1);
        setTimeout(() => ctx.setTenantDeltaProgress(false, '', null), 1400);
        if (pwdNotes.length) {
            await ctx.dlgAlert('Temporäre Kennwörter (bitte sicher notieren):\n\n' + pwdNotes.join('\n'), {
                title: 'Kennwörter',
                okText: 'Verstanden'
            });
        }
        window.__ms365TenantStepsRefreshMatch();
    }

    ctx.btnBatchG = document.getElementById('tenantDeltaProvisionGroupsBtn');
    if (ctx.btnBatchG) ctx.btnBatchG.addEventListener('click', () => ctx.runBatchProvision('group'));
    ctx.btnBatchP = document.getElementById('tenantDeltaProvisionPersonsBtn');
    if (ctx.btnBatchP) ctx.btnBatchP.addEventListener('click', () => ctx.runBatchProvision('person'));

    ctx.renderTenantInventoryMatchRows();
    ctx.renderTenantDeltaRows();
}
