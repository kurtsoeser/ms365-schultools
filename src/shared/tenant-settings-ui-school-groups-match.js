/**
 * tenant-settings-ui-school-groups-match.js
 * @param {Record<string, any>} ctx
 */
export function installTenantSchoolGroupsMatch(ctx) {
    ctx.setSingleGroupMatchStatus = function setSingleGroupMatchStatus(el, payload, expectedNick) {
        if (!el) return;
        el.classList.add('ts-match-pill');
        el.style.fontSize = '';
        el.style.lineHeight = '';
        el.style.background = '';
        el.style.color = '';
        el.title = '';
        const paint = function (state, text) {
            el.dataset.state = state;
            el.textContent = text;
        };
        if (!payload) {
            paint('idle', 'Noch nicht geprüft');
            ctx.updateSchoolWideGroupMatchActions();
            return;
        }
        if (payload.loading) {
            paint('loading', 'Prüfe …');
            return;
        }
        if (payload.found && payload.group) {
            const g = payload.group || {};
            const gid = ctx.normStr(g.id);
            const short = gid.length > 14 ? gid.slice(0, 12) + '…' : gid;
            const shown =
                ctx.normStr(g.displayName) ||
                ctx.normStr(g.mailNickname) ||
                (expectedNick ? expectedNick : '') ||
                'Zugeordnet';
            paint('found', shown);
            el.title =
                (g.displayName ? String(g.displayName) : '') +
                (g.mailNickname ? '\nAlias: ' + String(g.mailNickname) : '') +
                (g.mail ? '\nMail: ' + String(g.mail) : '') +
                (gid ? '\nObject-ID: ' + short : '');
            ctx.updateSchoolWideGroupMatchActions();
            return;
        }
        if (payload.notFound) {
            paint('missing', 'Nicht gefunden');
            el.title = 'Keine passende Gruppe in Microsoft 365 gefunden';
            ctx.updateSchoolWideGroupMatchActions();
            return;
        }
        if (payload.error) {
            paint('error', 'Fehler');
            el.title = String(payload.error || '');
            ctx.updateSchoolWideGroupMatchActions();
            return;
        }
        paint('idle', '–');
        ctx.updateSchoolWideGroupMatchActions();
    }

    ctx.updateSchoolWideGroupMatchActions = function updateSchoolWideGroupMatchActions() {
        const sgaId = ctx.getMatchedGroupId('sga');
        const svId = ctx.getMatchedGroupId('studentCouncil');
        if (ctx.btnUnmatchSgaGroup) ctx.btnUnmatchSgaGroup.hidden = !sgaId;
        if (ctx.btnOpenSgaGroup) ctx.btnOpenSgaGroup.hidden = !sgaId;
        if (ctx.btnUnmatchStudentCouncilGroup) ctx.btnUnmatchStudentCouncilGroup.hidden = !svId;
        if (ctx.btnOpenStudentCouncilGroup) ctx.btnOpenStudentCouncilGroup.hidden = !svId;
    }

    ctx.getMatchedGroupId = function getMatchedGroupId(kind) {
        const api = window.ms365AppDataV2;
        const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
        const matched = setup && setup.matched && typeof setup.matched === 'object' ? setup.matched : {};
        if (kind === 'sga') return ctx.normStr(matched.sgaGroupId || '');
        if (kind === 'studentCouncil') return ctx.normStr(matched.studentCouncilGroupId || '');
        return '';
    }

    ctx.patchMatchedGroupId = function patchMatchedGroupId(kind, group) {
        const api = window.ms365AppDataV2;
        if (!api || typeof api.patchSetup !== 'function') return;
        const gid = ctx.normStr(group && group.id);
        const key = kind === 'sga' ? 'sgaGroupId' : 'studentCouncilGroupId';
        api.patchSetup({ matched: { [key]: gid || null } });
        ctx.updateSchoolWideGroupMatchActions();
    }

    ctx.clearStoredSchoolWideGroupMatches = function clearStoredSchoolWideGroupMatches() {
        ctx.patchMatchedGroupId('sga', null);
        ctx.patchMatchedGroupId('studentCouncil', null);
        ctx.setSingleGroupMatchStatus(ctx.sgaGroupMatchCell, null);
        ctx.setSingleGroupMatchStatus(ctx.studentCouncilGroupMatchCell, null);
        if (ctx.sgaGroupSearchResults) {
            ctx.sgaGroupSearchResults.style.display = 'none';
            ctx.sgaGroupSearchResults.replaceChildren();
        }
        if (ctx.studentCouncilGroupSearchResults) {
            ctx.studentCouncilGroupSearchResults.style.display = 'none';
            ctx.studentCouncilGroupSearchResults.replaceChildren();
        }
        ctx.renderStatusOverview();
    }

    ctx.restoreStoredSchoolWideGroupMatchStatus = function restoreStoredSchoolWideGroupMatchStatus() {
        const sgaId = ctx.getMatchedGroupId('sga');
        const svId = ctx.getMatchedGroupId('studentCouncil');
        if (sgaId) {
            ctx.setSingleGroupMatchStatus(
                ctx.sgaGroupMatchCell,
                {
                    found: true,
                    group: {
                        id: sgaId,
                        displayName: ctx.expectedSgaGroupDisplayName(),
                        mailNickname: ctx.expectedSgaGroupMailNickname()
                    }
                },
                ctx.expectedSgaGroupMailNickname()
            );
        } else {
            ctx.setSingleGroupMatchStatus(ctx.sgaGroupMatchCell, null);
        }
        if (svId) {
            ctx.setSingleGroupMatchStatus(
                ctx.studentCouncilGroupMatchCell,
                {
                    found: true,
                    group: {
                        id: svId,
                        displayName: ctx.expectedStudentCouncilGroupDisplayName(),
                        mailNickname: ctx.expectedStudentCouncilGroupMailNickname()
                    }
                },
                ctx.expectedStudentCouncilGroupMailNickname()
            );
        } else {
            ctx.setSingleGroupMatchStatus(ctx.studentCouncilGroupMatchCell, null);
        }
        ctx.updateSchoolWideGroupMatchActions();
    }

    ctx.resolveSchoolWideGroupByStoredOrExpected = async function resolveSchoolWideGroupByStoredOrExpected(kind, expectedNick, expectedDn) {
        const token = await ctx.graphApi().getGraphToken();
        const storedId = ctx.getMatchedGroupId(kind);
        if (storedId && typeof ctx.graphApi().fetchGroup === 'function') {
            try {
                const g = await ctx.graphApi().fetchGroup(token, storedId);
                if (g && g.id) return { found: true, group: g, by: 'stored' };
            } catch {
                // gespeicherte ID ungültig → weiter mit Nick/Name-Suche
            }
        }
        const queries = [expectedNick, expectedDn].filter(Boolean);
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
                return mail && mail.startsWith(nickLc + '@');
            });
        if (matchByNick) return { found: true, group: matchByNick, by: 'nick' };
        return { found: false, notFound: true, hits: hits || [] };
    }

    ctx.renderSchoolWideGroupSearchResults = function renderSchoolWideGroupSearchResults(host, kind, list) {
        if (!host) return;
        host.replaceChildren();
        if (!list || !list.length) {
            host.style.display = 'none';
            return;
        }
        host.style.display = 'block';
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
                ctx.applyManualSchoolWideGroupMatch(kind, g);
            });
            row.appendChild(left);
            row.appendChild(btn);
            box.appendChild(row);
        });
        host.appendChild(box);
    }

    ctx.applyManualSchoolWideGroupMatch = function applyManualSchoolWideGroupMatch(kind, group) {
        if (!group || !group.id) return;
        const expectedNick =
            kind === 'sga' ? ctx.expectedSgaGroupMailNickname() : ctx.expectedStudentCouncilGroupMailNickname();
        const cell = kind === 'sga' ? ctx.sgaGroupMatchCell : ctx.studentCouncilGroupMatchCell;
        const resultsHost = kind === 'sga' ? ctx.sgaGroupSearchResults : ctx.studentCouncilGroupSearchResults;
        ctx.patchMatchedGroupId(kind, group);
        ctx.setSingleGroupMatchStatus(cell, { found: true, group: group }, expectedNick);
        if (resultsHost) {
            resultsHost.style.display = 'none';
            resultsHost.replaceChildren();
        }
        ctx.setSummary(
            (kind === 'sga' ? 'SGA' : 'Schülervertretung') + ': bestehende Microsoft-365-Gruppe zugeordnet.',
            'ok'
        );
        ctx.renderStatusOverview();
    }

    ctx.searchSchoolWideGroups = async function searchSchoolWideGroups(kind) {
        const inp = kind === 'sga' ? ctx.inpSgaGroupSearch : ctx.inpStudentCouncilGroupSearch;
        const host = kind === 'sga' ? ctx.sgaGroupSearchResults : ctx.studentCouncilGroupSearchResults;
        const q = inp && inp.value ? String(inp.value).trim() : '';
        if (!q) {
            ctx.setSummary('Bitte einen Suchbegriff eingeben (Name, Mail oder Alias).', 'warn');
            return;
        }
        ctx.setSummary((kind === 'sga' ? 'SGA' : 'Schülervertretung') + ': Suche Gruppen …', 'warn');
        const token = await ctx.graphApi().getGraphToken();
        const list = await ctx.graphApi().searchUnifiedGroups(token, q);
        ctx.renderSchoolWideGroupSearchResults(host, kind, list);
        if (!list || !list.length) {
            ctx.setSummary('Keine passenden Microsoft-365-Gruppen gefunden.', 'warn');
        } else {
            ctx.setSummary(list.length + ' Gruppe(n) gefunden – bitte zuordnen.', 'ok');
        }
    }

    ctx.unmatchSchoolWideGroup = function unmatchSchoolWideGroup(kind) {
        const cell = kind === 'sga' ? ctx.sgaGroupMatchCell : ctx.studentCouncilGroupMatchCell;
        const host = kind === 'sga' ? ctx.sgaGroupSearchResults : ctx.studentCouncilGroupSearchResults;
        ctx.patchMatchedGroupId(kind, null);
        ctx.setSingleGroupMatchStatus(cell, null);
        if (host) {
            host.style.display = 'none';
            host.replaceChildren();
        }
        ctx.setSummary((kind === 'sga' ? 'SGA' : 'Schülervertretung') + ': Match gelöst.', 'ok');
        ctx.renderStatusOverview();
    }

    ctx.openSchoolWideGroupInEntra = function openSchoolWideGroupInEntra(kind) {
        const gid = ctx.getMatchedGroupId(kind);
        if (!gid) {
            ctx.setSummary('Keine gematchte Gruppe zum Öffnen.', 'warn');
            return;
        }
        window.open(ctx.ENTRA_GROUP_DETAILS + encodeURIComponent(gid), '_blank', 'noopener');
    }

}
