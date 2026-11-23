/**
 * tenant-settings-ui-school-groups-lehrer.js
 * @param {Record<string, any>} ctx
 */
export function installTenantSchoolGroupsLehrer(ctx) {
    ctx.getLehrerSammelGroupId = function getLehrerSammelGroupId() {
        const api = window.ms365AppDataV2;
        const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
        const matched = setup && setup.matched ? setup.matched : {};
        return ctx.normStr(matched.lehrerGroupId || '');
    }

    ctx.patchLehrerSammelGroupId = function patchLehrerSammelGroupId(group) {
        const api = window.ms365AppDataV2;
        if (!api || typeof api.patchSetup !== 'function') return;
        const gid = ctx.normStr(group && group.id);
        api.patchSetup({ matched: { lehrerGroupId: gid || null } });
        if (gid && group && typeof api.upsertCatalogLink === 'function') {
            api.upsertCatalogLink({
                kind: 'sammelgruppe',
                code: 'lehrer',
                graphGroupId: gid,
                displayName: ctx.normStr(group.displayName),
                mailNickname: ctx.normStr(group.mailNickname),
                mode: group.mode || 'matched',
                syncStatus: ''
            });
        }
        ctx.updateLehrerSammelGroupMatchActions();
    }

    ctx.updateLehrerSammelGroupMatchActions = function updateLehrerSammelGroupMatchActions() {
        const gid = ctx.getLehrerSammelGroupId();
        if (ctx.btnUnmatchLehrerGroup) ctx.btnUnmatchLehrerGroup.hidden = !gid;
        if (ctx.btnOpenLehrerGroup) ctx.btnOpenLehrerGroup.hidden = !gid;
    }

    ctx.expectedLehrerGroupMailNickname = function expectedLehrerGroupMailNickname() {
        const api = window.ms365AppDataV2;
        const link = api && typeof api.getCatalogLink === 'function' ? api.getCatalogLink('sammelgruppe', 'lehrer') : null;
        const fromLink = link && ctx.normStr(link.mailNickname);
        if (fromLink) return ctx.graphApi().sanitizeUnifiedGroupMailNickname(fromLink);
        const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
        const draftNick =
            setup && setup.slgDraft && ctx.normStr(setup.slgDraft.slgNewMailNick)
                ? ctx.normStr(setup.slgDraft.slgNewMailNick)
                : 'lehrer';
        return ctx.graphApi().sanitizeUnifiedGroupMailNickname(draftNick || 'lehrer');
    }

    ctx.expectedLehrerGroupDisplayName = function expectedLehrerGroupDisplayName() {
        const api = window.ms365AppDataV2;
        const link = api && typeof api.getCatalogLink === 'function' ? api.getCatalogLink('sammelgruppe', 'lehrer') : null;
        const fromLink = link && ctx.normStr(link.displayName);
        if (fromLink) return fromLink;
        const nm = ctx.schoolNameInput ? ctx.normStr(ctx.schoolNameInput.value || '') : '';
        return nm ? 'Lehrer:innen ' + nm : 'Lehrer:innen';
    }

    ctx.expectedLehrerGroupMailAddress = function expectedLehrerGroupMailAddress() {
        const domain =
            typeof window.ms365GetSchoolDomainNoAt === 'function' ? ctx.normStr(window.ms365GetSchoolDomainNoAt()) : '';
        const nick = ctx.expectedLehrerGroupMailNickname();
        if (!domain || !nick) return '';
        return nick + '@' + domain.replace(/^@+/, '');
    }

    ctx.restoreLehrerSammelGroupMatchStatus = function restoreLehrerSammelGroupMatchStatus() {
        const gid = ctx.getLehrerSammelGroupId();
        if (gid) {
            ctx.setSingleGroupMatchStatus(
                ctx.lehrerGroupMatchCell,
                {
                    found: true,
                    group: {
                        id: gid,
                        displayName: ctx.expectedLehrerGroupDisplayName(),
                        mailNickname: ctx.expectedLehrerGroupMailNickname()
                    }
                },
                ctx.expectedLehrerGroupMailNickname()
            );
        } else {
            ctx.setSingleGroupMatchStatus(ctx.lehrerGroupMatchCell, null);
        }
        ctx.updateLehrerSammelGroupMatchActions();
    }

    ctx.resolveLehrerSammelGroupByStoredOrExpected = async function resolveLehrerSammelGroupByStoredOrExpected() {
        const expectedNick = ctx.expectedLehrerGroupMailNickname();
        const expectedDn = ctx.expectedLehrerGroupDisplayName();
        const expectedMail = ctx.expectedLehrerGroupMailAddress();
        const token = await ctx.graphApi().getGraphToken();
        const storedId = ctx.getLehrerSammelGroupId();
        if (storedId && typeof ctx.graphApi().fetchGroup === 'function') {
            try {
                const g = await ctx.graphApi().fetchGroup(token, storedId);
                if (g && g.id) return { found: true, group: g, by: 'stored' };
            } catch {
                // gespeicherte ID ungültig
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

    ctx.verifyLehrerSammelGroupExistence = async function verifyLehrerSammelGroupExistence() {
        const expectedNick = ctx.expectedLehrerGroupMailNickname();
        const expectedDn = ctx.expectedLehrerGroupDisplayName();
        ctx.setSingleGroupMatchStatus(ctx.lehrerGroupMatchCell, { loading: true }, expectedNick);
        const resolved = await ctx.resolveLehrerSammelGroupByStoredOrExpected();
        if (resolved && resolved.found && resolved.group) {
            ctx.patchLehrerSammelGroupId(resolved.group);
            ctx.setSingleGroupMatchStatus(ctx.lehrerGroupMatchCell, { found: true, group: resolved.group }, expectedNick);
            ctx.setSummary('Lehrer:innen-Sammelgruppe in Microsoft 365 gefunden.', 'ok');
            ctx.renderStatusOverview();
            return { found: true, group: resolved.group };
        }
        ctx.patchLehrerSammelGroupId(null);
        ctx.setSingleGroupMatchStatus(ctx.lehrerGroupMatchCell, { notFound: true }, expectedNick);
        ctx.setSummary(
            'Lehrer:innen-Sammelgruppe nicht gefunden (Suche: ' + (expectedNick || expectedDn || '–') + ').',
            'warn'
        );
        ctx.renderStatusOverview();
        return { found: false, notFound: true };
    }

    ctx.createLehrerSammelGroupExistence = async function createLehrerSammelGroupExistence() {
        const expectedNick = ctx.expectedLehrerGroupMailNickname();
        const expectedDn = ctx.expectedLehrerGroupDisplayName();
        const existing = await ctx.verifyLehrerSammelGroupExistence().catch(function () {
            return { found: false };
        });
        if (existing && existing.found && existing.group && existing.group.id) {
            ctx.setSummary('Lehrer:innen-Sammelgruppe existiert bereits – Zuordnung übernommen.', 'ok');
            return existing.group;
        }
        ctx.setSingleGroupMatchStatus(ctx.lehrerGroupMatchCell, { loading: true }, expectedNick);
        const token = await ctx.graphApi().getGraphToken();
        const desc = 'Alle Lehrer:innen (MS365-Schul-Tools / Schul‑Liste)';
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
        ctx.patchLehrerSammelGroupId(match || created);
        ctx.setSingleGroupMatchStatus(ctx.lehrerGroupMatchCell, { found: true, group: match || created }, expectedNick);
        ctx.setSummary('Lehrer:innen-Sammelgruppe angelegt.', 'ok');
        ctx.renderStatusOverview();
        await ctx.syncLehrerSammelGroupMembersAdditive({ silent: true });
        return created;
    }

    ctx.applyManualLehrerGroupMatch = function applyManualLehrerGroupMatch(group) {
        if (!group || !group.id) return;
        ctx.patchLehrerSammelGroupId(group);
        ctx.setSingleGroupMatchStatus(
            ctx.lehrerGroupMatchCell,
            { found: true, group: group },
            ctx.expectedLehrerGroupMailNickname()
        );
        if (ctx.lehrerGroupSearchResults) {
            ctx.lehrerGroupSearchResults.style.display = 'none';
            ctx.lehrerGroupSearchResults.replaceChildren();
        }
        ctx.setSummary('Lehrer:innen-Sammelgruppe: bestehende Microsoft-365-Gruppe zugeordnet.', 'ok');
        ctx.renderStatusOverview();
    }

    ctx.searchLehrerSammelGroups = async function searchLehrerSammelGroups() {
        const q = ctx.inpLehrerGroupSearch && ctx.inpLehrerGroupSearch.value ? String(ctx.inpLehrerGroupSearch.value).trim() : '';
        if (!q) {
            ctx.setSummary('Bitte einen Suchbegriff eingeben (Name, Mail oder Alias).', 'warn');
            return;
        }
        ctx.setSummary('Lehrer:innen-Sammelgruppe: Suche …', 'warn');
        const token = await ctx.graphApi().getGraphToken();
        const list = await ctx.graphApi().searchUnifiedGroups(token, q);
        if (ctx.lehrerGroupSearchResults) {
            ctx.lehrerGroupSearchResults.replaceChildren();
            if (!list || !list.length) {
                ctx.lehrerGroupSearchResults.style.display = 'none';
            } else {
                ctx.lehrerGroupSearchResults.style.display = 'block';
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
                        ctx.applyManualLehrerGroupMatch(g);
                    });
                    row.appendChild(left);
                    row.appendChild(btn);
                    box.appendChild(row);
                });
                ctx.lehrerGroupSearchResults.appendChild(box);
            }
        }
        if (!list || !list.length) {
            ctx.setSummary('Keine passenden Microsoft-365-Gruppen gefunden.', 'warn');
        } else {
            ctx.setSummary(list.length + ' Gruppe(n) gefunden – bitte zuordnen.', 'ok');
        }
    }

    ctx.unmatchLehrerSammelGroup = function unmatchLehrerSammelGroup() {
        ctx.patchLehrerSammelGroupId(null);
        ctx.setSingleGroupMatchStatus(ctx.lehrerGroupMatchCell, null);
        if (ctx.lehrerGroupSearchResults) {
            ctx.lehrerGroupSearchResults.style.display = 'none';
            ctx.lehrerGroupSearchResults.replaceChildren();
        }
        ctx.setSummary('Lehrer:innen-Sammelgruppe: Match gelöst.', 'ok');
        ctx.renderStatusOverview();
    }

    ctx.openLehrerSammelGroupInEntra = function openLehrerSammelGroupInEntra() {
        const gid = ctx.getLehrerSammelGroupId();
        if (!gid) {
            ctx.setSummary('Keine gematchte Lehrer:innen-Gruppe zum Öffnen.', 'warn');
            return;
        }
        window.open(ctx.ENTRA_GROUP_DETAILS + encodeURIComponent(gid), '_blank', 'noopener');
    }

}
