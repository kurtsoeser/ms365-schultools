/**
 * tenant-settings-ui-teachers-lifecycle.js
 * @param {Record<string, any>} ctx
 */
export function installTenantTeachersLifecycle(ctx) {
        ctx.collectTeacherEmailsFromRows = function collectTeacherEmailsFromRows(rows) {
            const out = [];
            const seen = new Set();
            (rows || []).forEach(function (r) {
                const em = ctx.normStr(r && r.email).toLowerCase();
                if (!em || em.indexOf('@') === -1 || seen.has(em)) return;
                seen.add(em);
                out.push(em);
            });
            return out;
        }

        ctx.diffTeacherEmailLists = function diffTeacherEmailLists(prev, next) {
            const prevSet = new Set(ctx.collectTeacherEmailsFromRows(prev));
            const nextList = ctx.collectTeacherEmailsFromRows(next);
            const nextSet = new Set(nextList);
            const added = nextList.filter(function (em) {
                return !prevSet.has(em);
            });
            const removed = [];
            prevSet.forEach(function (em) {
                if (!nextSet.has(em)) removed.push(em);
            });
            return { added: added, removed: removed };
        }

        /** @type {string[]} */
        ctx.lastTeacherLifecycleJoin = [];

        ctx.renderTeacherLifecyclePanel = function renderTeacherLifecyclePanel(prev, next) {
            const host = document.getElementById('tenantTeacherLifecycle');
            const listEl = document.getElementById('tenantTeacherLifecycleList');
            const metaEl = document.getElementById('tenantTeacherLifecycleMeta');
            if (!host) return;
            const gid = ctx.getLehrerSammelGroupId();
            const diff = ctx.diffTeacherEmailLists(prev, next);
            ctx.lastTeacherLifecycleJoin = gid ? diff.added.slice() : [];
            const btnApply = document.getElementById('tenantTeacherLifecycleApply');
            if (!diff.added.length) {
                host.hidden = true;
                if (listEl) listEl.replaceChildren();
                return;
            }
            host.hidden = false;
            if (metaEl) {
                if (!gid) {
                    metaEl.textContent =
                        diff.added.length === 1
                            ? 'Neue Lehrkraft in der Liste – zuerst unten die Lehrer:innen-Sammelgruppe in Microsoft 365 prüfen oder anlegen, dann Mitglieder aufnehmen.'
                            : diff.added.length +
                              ' neue Lehrkräfte in der Liste – zuerst unten die Lehrer:innen-Sammelgruppe in Microsoft 365 prüfen oder anlegen.';
                } else {
                    metaEl.textContent =
                        diff.added.length === 1
                            ? 'Eine neue Lehrkraft mit E-Mail in der Liste – in die Microsoft-365-Sammelgruppe „Lehrer:innen“ aufnehmen?'
                            : diff.added.length +
                              ' neue Lehrkräfte mit E-Mail – in die Microsoft-365-Sammelgruppe „Lehrer:innen“ aufnehmen?';
                }
            }
            if (btnApply) btnApply.disabled = !gid;
            if (listEl) {
                listEl.replaceChildren();
                diff.added.forEach(function (em) {
                    const li = document.createElement('li');
                    li.textContent = em;
                    listEl.appendChild(li);
                });
            }
        }

        ctx.applyTeacherLifecyclePreview = async function applyTeacherLifecyclePreview() {
            const gug = window.ms365GraphUnifiedGroups;
            const emails = ctx.lastTeacherLifecycleJoin.slice();
            const gid = ctx.getLehrerSammelGroupId();
            if (!gug || !gid || !emails.length) {
                ctx.setSummary('Keine neuen Lehrkräfte oder keine gematchte Sammelgruppe.', 'warn');
                return;
            }
            const btn = document.getElementById('tenantTeacherLifecycleApply');
            if (btn) btn.disabled = true;
            try {
                const token = await gug.getGraphToken();
                const r = await gug.syncEmailsToGroup(token, gid, emails, 'Lehrer', function () {});
                ctx.setSummary(
                    'Lehrer:innen-Gruppe: ' +
                        (r.ok || 0) +
                        ' aufgenommen' +
                        (r.skip ? ', ' + r.skip + ' schon Mitglied' : '') +
                        (r.fail ? ', ' + r.fail + ' Fehler' : '') +
                        '.',
                    r.fail ? 'warn' : 'ok'
                );
                const host = document.getElementById('tenantTeacherLifecycle');
                if (host) host.hidden = true;
                ctx.lastTeacherLifecycleJoin = [];
            } catch (e) {
                ctx.setSummary('Lehrer:innen-Gruppe: ' + (e.message || e), 'warn');
            } finally {
                if (btn) btn.disabled = false;
            }
        }

        ctx.syncLehrerSammelGroupMembersAdditive = async function syncLehrerSammelGroupMembersAdditive(opts) {
            const silent = opts && opts.silent;
            const gug = window.ms365GraphUnifiedGroups;
            const gid = ctx.getLehrerSammelGroupId();
            if (!gug || !gid) {
                if (!silent) ctx.setSummary('Zuerst Lehrer:innen-Sammelgruppe prüfen oder anlegen.', 'warn');
                return { ok: 0, skip: 0, fail: 0 };
            }
            const local = ctx.collectTeacherEmailsFromRows(ctx.getTeachersFromTextarea());
            if (!local.length) {
                if (!silent) ctx.setSummary('Keine Lehrkräfte mit E-Mail in der Liste.', 'warn');
                return { ok: 0, skip: 0, fail: 0 };
            }
            if (!silent) ctx.setSummary('Lehrer:innen-Gruppe: Mitglieder werden verglichen …', 'warn');
            const token = await gug.getGraphToken();
            const mem = await gug.fetchGroupMembers(token, gid);
            const items = mem && Array.isArray(mem.items) ? mem.items : [];
            const inGroup = new Set();
            items.forEach(function (m) {
                const em = ctx.normStr(m && (m.mail || m.userPrincipalName)).toLowerCase();
                if (em && em.indexOf('@') !== -1) inGroup.add(em);
            });
            const toAdd = local.filter(function (em) {
                return !inGroup.has(em);
            });
            if (!toAdd.length) {
                if (!silent) ctx.setSummary('Lehrer:innen-Gruppe: alle Listen-E-Mails sind bereits Mitglied.', 'ok');
                return { ok: 0, skip: local.length, fail: 0 };
            }
            const r = await gug.syncEmailsToGroup(token, gid, toAdd, 'Lehrer', function () {});
            if (!silent) {
                ctx.setSummary(
                    'Lehrer:innen-Gruppe: ' +
                        (r.ok || 0) +
                        ' ergänzt' +
                        (r.skip ? ', ' + r.skip + ' übersprungen' : '') +
                        (r.fail ? ', ' + r.fail + ' Fehler' : '') +
                        '.',
                    r.fail ? 'warn' : 'ok'
                );
            }
            return r;
        }

        window.ms365TenantTeachersAfterListChange = function (prevTeachers) {
            const prev = Array.isArray(prevTeachers) ? prevTeachers : [];
            const next = ctx.getTeachersFromTextarea();
            ctx.renderTeachersTableFromTextarea();
            ctx.renderTeacherLifecyclePanel(prev, next);
        };

        ctx.getSchuelerSammelGroupId = function getSchuelerSammelGroupId() {
            const api = window.ms365AppDataV2;
            const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
            const matched = setup && setup.matched ? setup.matched : {};
            return ctx.normStr(matched.schuelerGroupId || '');
        }

        ctx.patchSchuelerSammelGroupId = function patchSchuelerSammelGroupId(group) {
            const api = window.ms365AppDataV2;
            if (!api || typeof api.patchSetup !== 'function') return;
            const gid = ctx.normStr(group && group.id);
            api.patchSetup({ matched: { schuelerGroupId: gid || null } });
            if (gid && group && typeof api.upsertCatalogLink === 'function') {
                api.upsertCatalogLink({
                    kind: 'sammelgruppe',
                    code: 'schueler',
                    graphGroupId: gid,
                    displayName: ctx.normStr(group.displayName),
                    mailNickname: ctx.normStr(group.mailNickname),
                    mode: group.mode || 'matched',
                    syncStatus: ''
                });
            }
            ctx.updateSchuelerSammelGroupMatchActions();
        }

        ctx.updateSchuelerSammelGroupMatchActions = function updateSchuelerSammelGroupMatchActions() {
            const gid = ctx.getSchuelerSammelGroupId();
            if (ctx.btnUnmatchSchuelerGroup) ctx.btnUnmatchSchuelerGroup.hidden = !gid;
            if (ctx.btnOpenSchuelerGroup) ctx.btnOpenSchuelerGroup.hidden = !gid;
        }

        ctx.expectedSchuelerGroupMailNickname = function expectedSchuelerGroupMailNickname() {
            const api = window.ms365AppDataV2;
            const link = api && typeof api.getCatalogLink === 'function' ? api.getCatalogLink('sammelgruppe', 'schueler') : null;
            const fromLink = link && ctx.normStr(link.mailNickname);
            if (fromLink) return ctx.graphApi().sanitizeUnifiedGroupMailNickname(fromLink);
            const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
            const draftNick =
                setup && setup.slgDraft && ctx.normStr(setup.slgDraft.slgNewMailNick)
                    ? ctx.normStr(setup.slgDraft.slgNewMailNick)
                    : 'schueler';
            const nick = draftNick.toLowerCase().indexOf('lehrer') === 0 && draftNick.length <= 8 ? 'schueler' : draftNick;
            return ctx.graphApi().sanitizeUnifiedGroupMailNickname(nick || 'schueler');
        }

        ctx.expectedSchuelerGroupDisplayName = function expectedSchuelerGroupDisplayName() {
            const api = window.ms365AppDataV2;
            const link = api && typeof api.getCatalogLink === 'function' ? api.getCatalogLink('sammelgruppe', 'schueler') : null;
            const fromLink = link && ctx.normStr(link.displayName);
            if (fromLink) return fromLink;
            const nm = ctx.schoolNameInput ? ctx.normStr(ctx.schoolNameInput.value || '') : '';
            return nm ? 'Schüler:innen ' + nm : 'Schüler:innen';
        }

        ctx.expectedSchuelerGroupMailAddress = function expectedSchuelerGroupMailAddress() {
            const domain =
                typeof window.ms365GetSchoolDomainNoAt === 'function' ? ctx.normStr(window.ms365GetSchoolDomainNoAt()) : '';
            const nick = ctx.expectedSchuelerGroupMailNickname();
            if (!domain || !nick) return '';
            return nick + '@' + domain.replace(/^@+/, '');
        }

        ctx.restoreSchuelerSammelGroupMatchStatus = function restoreSchuelerSammelGroupMatchStatus() {
            const gid = ctx.getSchuelerSammelGroupId();
            if (gid) {
                ctx.setSingleGroupMatchStatus(
                    ctx.schuelerGroupMatchCell,
                    {
                        found: true,
                        group: {
                            id: gid,
                            displayName: ctx.expectedSchuelerGroupDisplayName(),
                            mailNickname: ctx.expectedSchuelerGroupMailNickname()
                        }
                    },
                    ctx.expectedSchuelerGroupMailNickname()
                );
            } else {
                ctx.setSingleGroupMatchStatus(ctx.schuelerGroupMatchCell, null);
            }
            ctx.updateSchuelerSammelGroupMatchActions();
        }

        ctx.resolveSchuelerSammelGroupByStoredOrExpected = async function resolveSchuelerSammelGroupByStoredOrExpected() {
            const expectedNick = ctx.expectedSchuelerGroupMailNickname();
            const expectedDn = ctx.expectedSchuelerGroupDisplayName();
            const expectedMail = ctx.expectedSchuelerGroupMailAddress();
            const token = await ctx.graphApi().getGraphToken();
            const storedId = ctx.getSchuelerSammelGroupId();
            if (storedId && typeof ctx.graphApi().fetchGroup === 'function') {
                try {
                    const g = await ctx.graphApi().fetchGroup(token, storedId);
                    if (g && g.id) return { found: true, group: g, by: 'stored' };
                } catch {
                    // ignore
                }
            }
            const queries = [expectedMail, expectedNick, expectedDn].filter(Boolean);
            let hits = [];
            for (let i = 0; i < queries.length; i++) {
                hits = await ctx.graphApi().searchUnifiedGroups(token, queries[i]);
                if (hits && hits.length) break;
            }
            const nickLc = ctx.normStr(expectedNick).toLowerCase();
            const matchByNick =
                Array.isArray(hits) &&
                hits.find(function (g) {
                    const mn = ctx.normStr(g && g.mailNickname).toLowerCase();
                    if (mn && mn === nickLc) return true;
                    const mail = ctx.normStr(g && g.mail).toLowerCase();
                    return mail && expectedMail && mail === expectedMail.toLowerCase();
                });
            if (matchByNick) return { found: true, group: matchByNick, by: 'nick' };
            return { found: false, notFound: true, hits: hits || [] };
        }

        ctx.verifySchuelerSammelGroupExistence = async function verifySchuelerSammelGroupExistence() {
            const expectedNick = ctx.expectedSchuelerGroupMailNickname();
            const expectedDn = ctx.expectedSchuelerGroupDisplayName();
            ctx.setSingleGroupMatchStatus(ctx.schuelerGroupMatchCell, { loading: true }, expectedNick);
            const resolved = await ctx.resolveSchuelerSammelGroupByStoredOrExpected();
            if (resolved && resolved.found && resolved.group) {
                ctx.patchSchuelerSammelGroupId(resolved.group);
                ctx.setSingleGroupMatchStatus(ctx.schuelerGroupMatchCell, { found: true, group: resolved.group }, expectedNick);
                ctx.setSummary('Schüler:innen-Sammelgruppe in Microsoft 365 gefunden.', 'ok');
                ctx.renderStatusOverview();
                return { found: true, group: resolved.group };
            }
            ctx.patchSchuelerSammelGroupId(null);
            ctx.setSingleGroupMatchStatus(ctx.schuelerGroupMatchCell, { notFound: true }, expectedNick);
            ctx.setSummary(
                'Schüler:innen-Sammelgruppe nicht gefunden (Suche: ' + (expectedNick || expectedDn || '–') + ').',
                'warn'
            );
            ctx.renderStatusOverview();
            return { found: false, notFound: true };
        }

        ctx.createSchuelerSammelGroupExistence = async function createSchuelerSammelGroupExistence() {
            const expectedNick = ctx.expectedSchuelerGroupMailNickname();
            const expectedDn = ctx.expectedSchuelerGroupDisplayName();
            const existing = await ctx.verifySchuelerSammelGroupExistence().catch(function () {
                return { found: false };
            });
            if (existing && existing.found && existing.group && existing.group.id) {
                ctx.setSummary('Schüler:innen-Sammelgruppe existiert bereits – Zuordnung übernommen.', 'ok');
                return existing.group;
            }
            ctx.setSingleGroupMatchStatus(ctx.schuelerGroupMatchCell, { loading: true }, expectedNick);
            const token = await ctx.graphApi().getGraphToken();
            const desc = 'Alle Schüler:innen (MS365-Schul-Tools / Schul‑Liste)';
            const created = await ctx.graphApi().createUnifiedGroup(token, expectedDn, expectedNick, desc);
            const match =
                created && created.id
                    ? {
                          id: created.id,
                          displayName: created.displayName || expectedDn,
                          mailNickname: created.mailNickname || expectedNick,
                          mode: 'created'
                      }
                    : null;
            ctx.patchSchuelerSammelGroupId(match || created);
            ctx.setSingleGroupMatchStatus(ctx.schuelerGroupMatchCell, { found: true, group: match || created }, expectedNick);
            ctx.setSummary('Schüler:innen-Sammelgruppe angelegt.', 'ok');
            ctx.renderStatusOverview();
            await ctx.syncSchuelerSammelGroupMembersAdditive({ silent: true });
            return created;
        }

        ctx.applyManualSchuelerGroupMatch = function applyManualSchuelerGroupMatch(group) {
            if (!group || !group.id) return;
            ctx.patchSchuelerSammelGroupId(group);
            ctx.setSingleGroupMatchStatus(
                ctx.schuelerGroupMatchCell,
                { found: true, group: group },
                ctx.expectedSchuelerGroupMailNickname()
            );
            if (ctx.schuelerGroupSearchResults) {
                ctx.schuelerGroupSearchResults.style.display = 'none';
                ctx.schuelerGroupSearchResults.replaceChildren();
            }
            ctx.setSummary('Schüler:innen-Sammelgruppe: bestehende Microsoft-365-Gruppe zugeordnet.', 'ok');
            ctx.renderStatusOverview();
        }

        ctx.searchSchuelerSammelGroups = async function searchSchuelerSammelGroups() {
            const q = ctx.inpSchuelerGroupSearch && ctx.inpSchuelerGroupSearch.value ? String(ctx.inpSchuelerGroupSearch.value).trim() : '';
            if (!q) {
                ctx.setSummary('Bitte einen Suchbegriff eingeben (Name, Mail oder Alias).', 'warn');
                return;
            }
            ctx.setSummary('Schüler:innen-Sammelgruppe: Suche …', 'warn');
            const token = await ctx.graphApi().getGraphToken();
            const list = await ctx.graphApi().searchUnifiedGroups(token, q);
            if (ctx.schuelerGroupSearchResults) {
                ctx.schuelerGroupSearchResults.replaceChildren();
                if (!list || !list.length) {
                    ctx.schuelerGroupSearchResults.style.display = 'none';
                } else {
                    ctx.schuelerGroupSearchResults.style.display = 'block';
                    const box = document.createElement('div');
                    box.style.border = '1px solid var(--border)';
                    box.style.borderRadius = '10px';
                    box.style.overflow = 'hidden';
                    list.forEach(function (g, idx) {
                        const row = document.createElement('div');
                        row.style.display = 'flex';
                        row.style.flexWrap = 'wrap';
                        row.style.gap = '10px';
                        row.style.alignItems = 'center';
                        row.style.justifyContent = 'space-between';
                        row.style.padding = '10px 12px';
                        if (idx > 0) row.style.borderTop = '1px solid var(--border)';
                        const dn = ctx.normStr(g && g.displayName) || '(ohne Namen)';
                        const mail = ctx.normStr(g && g.mail) || '–';
                        const nick = ctx.normStr(g && g.mailNickname) || '–';
                        const left = document.createElement('div');
                        left.style.minWidth = '0';
                        left.style.flex = '1 1 220px';
                        left.innerHTML =
                            '<div style="font-weight:700;line-height:1.25;">' +
                            ctx.escapeHtml(dn) +
                            '</div>' +
                            '<div class="muted" style="margin-top:2px;font-size:0.92em;">Alias: <code>' +
                            ctx.escapeHtml(nick) +
                            '</code> · ' +
                            ctx.escapeHtml(mail) +
                            '</div>';
                        const btn = document.createElement('button');
                        btn.type = 'button';
                        btn.className = 'btn btn-primary';
                        btn.innerHTML = '<i class="bi bi-link-45deg" aria-hidden="true"></i> Zuordnen';
                        btn.addEventListener('click', function () {
                            ctx.applyManualSchuelerGroupMatch(g);
                        });
                        row.appendChild(left);
                        row.appendChild(btn);
                        box.appendChild(row);
                    });
                    ctx.schuelerGroupSearchResults.appendChild(box);
                }
            }
            if (!list || !list.length) {
                ctx.setSummary('Keine passenden Microsoft-365-Gruppen gefunden.', 'warn');
            } else {
                ctx.setSummary(list.length + ' Gruppe(n) gefunden – bitte zuordnen.', 'ok');
            }
        }

        ctx.unmatchSchuelerSammelGroup = function unmatchSchuelerSammelGroup() {
            ctx.patchSchuelerSammelGroupId(null);
            ctx.setSingleGroupMatchStatus(ctx.schuelerGroupMatchCell, null);
            if (ctx.schuelerGroupSearchResults) {
                ctx.schuelerGroupSearchResults.style.display = 'none';
                ctx.schuelerGroupSearchResults.replaceChildren();
            }
            ctx.setSummary('Schüler:innen-Sammelgruppe: Match gelöst.', 'ok');
            ctx.renderStatusOverview();
        }

        ctx.openSchuelerSammelGroupInEntra = function openSchuelerSammelGroupInEntra() {
            const gid = ctx.getSchuelerSammelGroupId();
            if (!gid) {
                ctx.setSummary('Keine gematchte Schüler:innen-Gruppe zum Öffnen.', 'warn');
                return;
            }
            window.open(ctx.ENTRA_GROUP_DETAILS + encodeURIComponent(gid), '_blank', 'noopener');
        }

        ctx.collectStudentEmailsFromRows = function collectStudentEmailsFromRows(rows) {
            const out = [];
            const seen = new Set();
            (rows || []).forEach(function (r) {
                const em = ctx.normStr(r && r.email).toLowerCase();
                if (!em || em.indexOf('@') === -1 || seen.has(em)) return;
                seen.add(em);
                out.push(em);
            });
            return out;
        }

        ctx.diffStudentEmailLists = function diffStudentEmailLists(prev, next) {
            const prevSet = new Set(ctx.collectStudentEmailsFromRows(prev));
            const nextList = ctx.collectStudentEmailsFromRows(next);
            const nextSet = new Set(nextList);
            const added = nextList.filter(function (em) {
                return !prevSet.has(em);
            });
            const removed = [];
            prevSet.forEach(function (em) {
                if (!nextSet.has(em)) removed.push(em);
            });
            return { added: added, removed: removed };
        }

        /** @type {string[]} */
        ctx.lastStudentSammelLifecycleJoin = [];

        ctx.renderStudentSammelLifecyclePanel = function renderStudentSammelLifecyclePanel(prev, next) {
            const host = document.getElementById('tenantStudentSammelLifecycle');
            const listEl = document.getElementById('tenantStudentSammelLifecycleList');
            const metaEl = document.getElementById('tenantStudentSammelLifecycleMeta');
            if (!host) return;
            const gid = ctx.getSchuelerSammelGroupId();
            const diff = ctx.diffStudentEmailLists(prev, next);
            ctx.lastStudentSammelLifecycleJoin = gid ? diff.added.slice() : [];
            const btnApply = document.getElementById('tenantStudentSammelLifecycleApply');
            if (!diff.added.length) {
                host.hidden = true;
                if (listEl) listEl.replaceChildren();
                return;
            }
            host.hidden = false;
            if (metaEl) {
                if (!gid) {
                    metaEl.textContent =
                        diff.added.length === 1
                            ? 'Neue:r Schüler:in in der Liste – zuerst unten die Schüler:innen-Sammelgruppe in Microsoft 365 prüfen oder anlegen.'
                            : diff.added.length +
                              ' neue Schüler:innen in der Liste – zuerst unten die Schüler:innen-Sammelgruppe prüfen oder anlegen.';
                } else {
                    metaEl.textContent =
                        diff.added.length === 1
                            ? 'Ein:e neue:r Schüler:in mit E-Mail – in die Microsoft-365-Sammelgruppe „Schüler:innen“ aufnehmen?'
                            : diff.added.length +
                              ' neue Schüler:innen mit E-Mail – in die Microsoft-365-Sammelgruppe „Schüler:innen“ aufnehmen?';
                }
            }
            if (btnApply) btnApply.disabled = !gid;
            if (listEl) {
                listEl.replaceChildren();
                diff.added.forEach(function (em) {
                    const li = document.createElement('li');
                    li.textContent = em;
                    listEl.appendChild(li);
                });
            }
        }

        ctx.applyStudentSammelLifecyclePreview = async function applyStudentSammelLifecyclePreview() {
            const gug = window.ms365GraphUnifiedGroups;
            const emails = ctx.lastStudentSammelLifecycleJoin.slice();
            const gid = ctx.getSchuelerSammelGroupId();
            if (!gug || !gid || !emails.length) {
                ctx.setSummary('Keine neuen Schüler:innen oder keine gematchte Sammelgruppe.', 'warn');
                return;
            }
            const btn = document.getElementById('tenantStudentSammelLifecycleApply');
            if (btn) btn.disabled = true;
            try {
                const token = await gug.getGraphToken();
                const r = await gug.syncEmailsToGroup(token, gid, emails, 'Schüler', function () {});
                ctx.setSummary(
                    'Schüler:innen-Gruppe: ' +
                        (r.ok || 0) +
                        ' aufgenommen' +
                        (r.skip ? ', ' + r.skip + ' schon Mitglied' : '') +
                        (r.fail ? ', ' + r.fail + ' Fehler' : '') +
                        '.',
                    r.fail ? 'warn' : 'ok'
                );
                const host = document.getElementById('tenantStudentSammelLifecycle');
                if (host) host.hidden = true;
                ctx.lastStudentSammelLifecycleJoin = [];
            } catch (e) {
                ctx.setSummary('Schüler:innen-Gruppe: ' + (e.message || e), 'warn');
            } finally {
                if (btn) btn.disabled = false;
            }
        }

        ctx.syncSchuelerSammelGroupMembersAdditive = async function syncSchuelerSammelGroupMembersAdditive(opts) {
            const silent = opts && opts.silent;
            const gug = window.ms365GraphUnifiedGroups;
            const gid = ctx.getSchuelerSammelGroupId();
            if (!gug || !gid) {
                if (!silent) ctx.setSummary('Zuerst Schüler:innen-Sammelgruppe prüfen oder anlegen.', 'warn');
                return { ok: 0, skip: 0, fail: 0 };
            }
            const local = ctx.collectStudentEmailsFromRows(ctx.getStudentsFromTextarea());
            if (!local.length) {
                if (!silent) ctx.setSummary('Keine Schüler:innen mit E-Mail in der Liste.', 'warn');
                return { ok: 0, skip: 0, fail: 0 };
            }
            if (!silent) ctx.setSummary('Schüler:innen-Gruppe: Mitglieder werden verglichen …', 'warn');
            const token = await gug.getGraphToken();
            const mem = await gug.fetchGroupMembers(token, gid);
            const items = mem && Array.isArray(mem.items) ? mem.items : [];
            const inGroup = new Set();
            items.forEach(function (m) {
                const em = ctx.normStr(m && (m.mail || m.userPrincipalName)).toLowerCase();
                if (em && em.indexOf('@') !== -1) inGroup.add(em);
            });
            const toAdd = local.filter(function (em) {
                return !inGroup.has(em);
            });
            if (!toAdd.length) {
                if (!silent) ctx.setSummary('Schüler:innen-Gruppe: alle Listen-E-Mails sind bereits Mitglied.', 'ok');
                return { ok: 0, skip: local.length, fail: 0 };
            }
            const r = await gug.syncEmailsToGroup(token, gid, toAdd, 'Schüler', function () {});
            if (!silent) {
                ctx.setSummary(
                    'Schüler:innen-Gruppe: ' +
                        (r.ok || 0) +
                        ' ergänzt' +
                        (r.skip ? ', ' + r.skip + ' übersprungen' : '') +
                        (r.fail ? ', ' + r.fail + ' Fehler' : '') +
                        '.',
                    r.fail ? 'warn' : 'ok'
                );
            }
            return r;
        }

        window.ms365TenantStudentsAfterListChange = function (prevStudents) {
            const prev = Array.isArray(prevStudents) ? prevStudents : [];
            const next = ctx.getStudentsFromTextarea();
            ctx.renderStudentsTableFromTextarea();
            ctx.renderStudentSammelLifecyclePanel(prev, next);
        };

        ctx.lastLifecyclePreview = null;
        ctx.pendingSisImport = null;
        /** @type {{ incoming: object[], existing: object[], backupLines: string, result: object, diff: object, routeSuffix?: string }|null} */
        ctx.pendingTeachersWuImport = null;

}
