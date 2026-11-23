/**
 * tenant-settings-ui-school-groups-naming.js
 * @param {Record<string, any>} ctx
 */
export function installTenantSchoolGroupsNaming(ctx) {
    ctx.schoolBaseNick = function schoolBaseNick() {
        const nm = ctx.schoolNameInput ? ctx.normStr(ctx.schoolNameInput.value || '') : '';
        return nm.replace(/[^a-zA-Z0-9]/g, '').toLowerCase().slice(0, 24);
    }

    ctx.defaultSgaGroupMailNickname = function defaultSgaGroupMailNickname() {
        const base = ctx.schoolBaseNick();
        if (!base) return '';
        return ctx.graphApi().sanitizeUnifiedGroupMailNickname('sga' + base);
    }

    ctx.defaultSgaGroupDisplayName = function defaultSgaGroupDisplayName() {
        const nm = ctx.schoolNameInput ? ctx.normStr(ctx.schoolNameInput.value || '') : '';
        return nm ? 'SGA ' + nm : '';
    }

    ctx.defaultStudentCouncilGroupMailNickname = function defaultStudentCouncilGroupMailNickname() {
        const base = ctx.schoolBaseNick();
        if (!base) return '';
        const yLbl = ctx.getDisplayedSchoolYearLabel() || ctx.currentSchoolYearLabel();
        const digits = String(yLbl || '').replace(/\D/g, '').slice(0, 6);
        return ctx.graphApi().sanitizeUnifiedGroupMailNickname('sv' + digits + base);
    }

    ctx.defaultStudentCouncilGroupDisplayName = function defaultStudentCouncilGroupDisplayName() {
        const nm = ctx.schoolNameInput ? ctx.normStr(ctx.schoolNameInput.value || '') : '';
        const yLbl = ctx.getDisplayedSchoolYearLabel() || ctx.currentSchoolYearLabel();
        if (!nm) return '';
        return 'Schülervertretung ' + nm + ' ' + yLbl;
    }

    ctx.readSgaGroupDisplayName = function readSgaGroupDisplayName() {
        const v = ctx.inpSgaNewDisplayName && ctx.normStr(ctx.inpSgaNewDisplayName.value);
        return v || ctx.defaultSgaGroupDisplayName();
    }

    ctx.readSgaGroupMailNickname = function readSgaGroupMailNickname() {
        const v = ctx.inpSgaNewMailNick && ctx.normStr(ctx.inpSgaNewMailNick.value);
        if (v) return ctx.graphApi().sanitizeUnifiedGroupMailNickname(v);
        return ctx.defaultSgaGroupMailNickname();
    }

    ctx.readStudentCouncilGroupDisplayName = function readStudentCouncilGroupDisplayName() {
        const v = ctx.inpSvNewDisplayName && ctx.normStr(ctx.inpSvNewDisplayName.value);
        return v || ctx.defaultStudentCouncilGroupDisplayName();
    }

    ctx.readStudentCouncilGroupMailNickname = function readStudentCouncilGroupMailNickname() {
        const v = ctx.inpSvNewMailNick && ctx.normStr(ctx.inpSvNewMailNick.value);
        if (v) return ctx.graphApi().sanitizeUnifiedGroupMailNickname(v);
        return ctx.defaultStudentCouncilGroupMailNickname();
    }

    ctx.expectedSgaGroupMailNickname = function expectedSgaGroupMailNickname() {
        return ctx.readSgaGroupMailNickname();
    }

    ctx.expectedSgaGroupDisplayName = function expectedSgaGroupDisplayName() {
        return ctx.readSgaGroupDisplayName();
    }

    ctx.expectedStudentCouncilGroupMailNickname = function expectedStudentCouncilGroupMailNickname() {
        return ctx.readStudentCouncilGroupMailNickname();
    }

    ctx.expectedStudentCouncilGroupDisplayName = function expectedStudentCouncilGroupDisplayName() {
        return ctx.readStudentCouncilGroupDisplayName();
    }

    ctx.updateSchoolWideGroupMailPreview = function updateSchoolWideGroupMailPreview(kind) {
        const nick =
            kind === 'sga' ? ctx.readSgaGroupMailNickname() : ctx.readStudentCouncilGroupMailNickname();
        const domain =
            typeof window.ms365GetSchoolDomainNoAt === 'function' ? ctx.normStr(window.ms365GetSchoolDomainNoAt()) : '';
        const el = document.getElementById(kind === 'sga' ? 'tenantSgaMailPreview' : 'tenantSvMailPreview');
        if (!el) return;
        if (!nick) {
            el.textContent = '—';
            return;
        }
        el.textContent = nick + '@' + (domain || 'ihre-schul-domain.at');
    }

    ctx.patchSchoolWideGroupNamingDraft = function patchSchoolWideGroupNamingDraft(kind) {
        const api = window.ms365AppDataV2;
        if (!api || typeof api.patchSetup !== 'function') return;
        if (kind === 'sga') {
            api.patchSetup({
                sgaDraft: {
                    sgaNewDisplayName: ctx.inpSgaNewDisplayName ? ctx.normStr(ctx.inpSgaNewDisplayName.value) : '',
                    sgaNewMailNick: ctx.inpSgaNewMailNick ? ctx.normStr(ctx.inpSgaNewMailNick.value) : ''
                }
            });
        } else {
            api.patchSetup({
                studentCouncilDraft: {
                    svNewDisplayName: ctx.inpSvNewDisplayName ? ctx.normStr(ctx.inpSvNewDisplayName.value) : '',
                    svNewMailNick: ctx.inpSvNewMailNick ? ctx.normStr(ctx.inpSvNewMailNick.value) : ''
                }
            });
        }
        ctx.updateSchoolWideGroupMailPreview(kind);
    }

    ctx.restoreSchoolWideGroupNamingFields = function restoreSchoolWideGroupNamingFields() {
        const api = window.ms365AppDataV2;
        const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
        const sga = setup && setup.sgaDraft ? setup.sgaDraft : {};
        const sv = setup && setup.studentCouncilDraft ? setup.studentCouncilDraft : {};
        if (ctx.inpSgaNewDisplayName) {
            ctx.inpSgaNewDisplayName.value = ctx.normStr(sga.sgaNewDisplayName);
            ctx.inpSgaNewDisplayName.placeholder = ctx.defaultSgaGroupDisplayName() || 'SGA …';
        }
        if (ctx.inpSgaNewMailNick) {
            ctx.inpSgaNewMailNick.value = ctx.normStr(sga.sgaNewMailNick);
            ctx.inpSgaNewMailNick.placeholder = ctx.defaultSgaGroupMailNickname() || 'sga…';
        }
        if (ctx.inpSvNewDisplayName) {
            ctx.inpSvNewDisplayName.value = ctx.normStr(sv.svNewDisplayName);
            ctx.inpSvNewDisplayName.placeholder = ctx.defaultStudentCouncilGroupDisplayName() || 'Schülervertretung …';
        }
        if (ctx.inpSvNewMailNick) {
            ctx.inpSvNewMailNick.value = ctx.normStr(sv.svNewMailNick);
            ctx.inpSvNewMailNick.placeholder = ctx.defaultStudentCouncilGroupMailNickname() || 'sv…';
        }
        ctx.updateSchoolWideGroupMailPreview('sga');
        ctx.updateSchoolWideGroupMailPreview('studentCouncil');
    }

    ctx.bindSchoolWideGroupNamingField = function bindSchoolWideGroupNamingField(inp, kind) {
        if (!inp || inp.dataset.tenantSchoolWideNamingBound) return;
        inp.dataset.tenantSchoolWideNamingBound = '1';
        const onChange = function () {
            ctx.patchSchoolWideGroupNamingDraft(kind);
            ctx.scheduleAutoSave();
        };
        inp.addEventListener('input', onChange);
        inp.addEventListener('change', onChange);
    }

    ctx.refreshSchoolWideGroupNamingPlaceholders = function refreshSchoolWideGroupNamingPlaceholders() {
        if (ctx.inpSgaNewDisplayName && !ctx.normStr(ctx.inpSgaNewDisplayName.value)) {
            ctx.inpSgaNewDisplayName.placeholder = ctx.defaultSgaGroupDisplayName() || 'SGA …';
        }
        if (ctx.inpSgaNewMailNick && !ctx.normStr(ctx.inpSgaNewMailNick.value)) {
            ctx.inpSgaNewMailNick.placeholder = ctx.defaultSgaGroupMailNickname() || 'sga…';
        }
        if (ctx.inpSvNewDisplayName && !ctx.normStr(ctx.inpSvNewDisplayName.value)) {
            ctx.inpSvNewDisplayName.placeholder = ctx.defaultStudentCouncilGroupDisplayName() || 'Schülervertretung …';
        }
        if (ctx.inpSvNewMailNick && !ctx.normStr(ctx.inpSvNewMailNick.value)) {
            ctx.inpSvNewMailNick.placeholder = ctx.defaultStudentCouncilGroupMailNickname() || 'sv…';
        }
        ctx.updateSchoolWideGroupMailPreview('sga');
        ctx.updateSchoolWideGroupMailPreview('studentCouncil');
    }

    ctx.verifySgaGroupExistence = async function verifySgaGroupExistence() {
        if (!ctx.btnVerifySgaGraph) {
            // btnVerifySgaGraph existiert als „Guard“ für die Seite – wenn nicht, ist die UI nicht aktiv.
            throw new Error('UI nicht bereit.');
        }
        const expectedNick = ctx.expectedSgaGroupMailNickname();
        const expectedDn = ctx.expectedSgaGroupDisplayName();
        if (!expectedNick || !expectedDn) {
            // Ohne Schulname trotzdem gespeicherte Zuordnung behalten/anzeigen
            const storedOnly = ctx.getMatchedGroupId('sga');
            if (storedOnly) {
                ctx.setSingleGroupMatchStatus(
                    ctx.sgaGroupMatchCell,
                    { found: true, group: { id: storedOnly } },
                    ''
                );
                return { found: true, group: { id: storedOnly } };
            }
            ctx.patchMatchedGroupId('sga', null);
            ctx.setSingleGroupMatchStatus(ctx.sgaGroupMatchCell, { notFound: true }, expectedNick || '');
            ctx.setSummary('SGA: Bitte zuerst „Schulname“ setzen.', 'warn');
            return { found: false, notFound: true };
        }
        ctx.setSingleGroupMatchStatus(ctx.sgaGroupMatchCell, { loading: true }, expectedNick);
        const resolved = await ctx.resolveSchoolWideGroupByStoredOrExpected('sga', expectedNick, expectedDn);
        if (resolved && resolved.found && resolved.group) {
            ctx.patchMatchedGroupId('sga', resolved.group);
            ctx.setSingleGroupMatchStatus(ctx.sgaGroupMatchCell, { found: true, group: resolved.group }, expectedNick);
            return { found: true, group: resolved.group };
        }
        ctx.patchMatchedGroupId('sga', null);
        ctx.setSingleGroupMatchStatus(ctx.sgaGroupMatchCell, { notFound: true }, expectedNick);
        return { found: false, notFound: true };
    }

    ctx.createSgaGroupExistence = async function createSgaGroupExistence() {
        const expectedNick = ctx.expectedSgaGroupMailNickname();
        const expectedDn = ctx.expectedSgaGroupDisplayName();
        if (!expectedNick || !expectedDn) {
            ctx.setSummary('SGA: Bitte zuerst „Schulname“ setzen.', 'warn');
            ctx.patchMatchedGroupId('sga', null);
            ctx.setSingleGroupMatchStatus(ctx.sgaGroupMatchCell, { notFound: true }, expectedNick || '');
            return null;
        }
        const mode = ctx.selSgaMode ? ctx.normStr(ctx.selSgaMode.value || 'group').toLowerCase() : 'group';

        // Vermeidet Dubletten: Erst prüfen, dann ggf. anlegen.
        const existing = await ctx.verifySgaGroupExistence().catch(function () {
            return { found: false };
        });
        if (existing && existing.found && existing.group && existing.group.id) return existing.group;

        ctx.setSingleGroupMatchStatus(ctx.sgaGroupMatchCell, { loading: true }, expectedNick);
        const token = await ctx.graphApi().getGraphToken();
        const created = await ctx.graphApi().createUnifiedGroup(token, expectedDn, expectedNick, 'MS365-Schul-Tools – SGA');
        if (mode === 'group') {
            try {
                await ctx.graphApi().provisionTeamForGroup(token, created && created.id ? created.id : '');
            } catch {
                // optional: Team-Provision kann warten/fehlschlagen
            }
        }
        const match = created && created.id ? { id: created.id, displayName: created.displayName || expectedDn, mailNickname: created.mailNickname || expectedNick } : null;
        ctx.patchMatchedGroupId('sga', match || created);
        ctx.setSingleGroupMatchStatus(ctx.sgaGroupMatchCell, { found: true, group: match || created }, expectedNick);
        ctx.setSummary(
            'SGA-Gruppe angelegt: „' + ctx.readSgaGroupDisplayName() + '“ (' + ctx.readSgaGroupMailNickname() + '@…).',
            'ok'
        );
        ctx.renderStatusOverview();
        return created;
    }

    ctx.verifyStudentCouncilGroupExistence = async function verifyStudentCouncilGroupExistence() {
        const expectedNick = ctx.expectedStudentCouncilGroupMailNickname();
        const expectedDn = ctx.expectedStudentCouncilGroupDisplayName();
        if (!expectedNick || !expectedDn) {
            const storedOnly = ctx.getMatchedGroupId('studentCouncil');
            if (storedOnly) {
                ctx.setSingleGroupMatchStatus(
                    ctx.studentCouncilGroupMatchCell,
                    { found: true, group: { id: storedOnly } },
                    ''
                );
                return { found: true, group: { id: storedOnly } };
            }
            ctx.patchMatchedGroupId('studentCouncil', null);
            ctx.setSingleGroupMatchStatus(ctx.studentCouncilGroupMatchCell, { notFound: true }, expectedNick || '');
            ctx.setSummary('Schülervertretung: Bitte zuerst „Schulname“ setzen.', 'warn');
            return { found: false, notFound: true };
        }
        ctx.setSingleGroupMatchStatus(ctx.studentCouncilGroupMatchCell, { loading: true }, expectedNick);
        const resolved = await ctx.resolveSchoolWideGroupByStoredOrExpected(
            'studentCouncil',
            expectedNick,
            expectedDn
        );
        if (resolved && resolved.found && resolved.group) {
            ctx.patchMatchedGroupId('studentCouncil', resolved.group);
            ctx.setSingleGroupMatchStatus(
                ctx.studentCouncilGroupMatchCell,
                { found: true, group: resolved.group },
                expectedNick
            );
            return { found: true, group: resolved.group };
        }
        ctx.patchMatchedGroupId('studentCouncil', null);
        ctx.setSingleGroupMatchStatus(ctx.studentCouncilGroupMatchCell, { notFound: true }, expectedNick);
        return { found: false, notFound: true };
    }

    ctx.createStudentCouncilGroupExistence = async function createStudentCouncilGroupExistence() {
        const expectedNick = ctx.expectedStudentCouncilGroupMailNickname();
        const expectedDn = ctx.expectedStudentCouncilGroupDisplayName();
        if (!expectedNick || !expectedDn) {
            ctx.setSummary('Schülervertretung: Bitte zuerst „Schulname“ setzen.', 'warn');
            ctx.patchMatchedGroupId('studentCouncil', null);
            ctx.setSingleGroupMatchStatus(ctx.studentCouncilGroupMatchCell, { notFound: true }, expectedNick || '');
            return null;
        }

        const existing = await ctx.verifyStudentCouncilGroupExistence().catch(function () {
            return { found: false };
        });
        if (existing && existing.found && existing.group && existing.group.id) return existing.group;

        ctx.setSingleGroupMatchStatus(ctx.studentCouncilGroupMatchCell, { loading: true }, expectedNick);
        const token = await ctx.graphApi().getGraphToken();
        const created = await ctx.graphApi().createUnifiedGroup(
            token,
            expectedDn,
            expectedNick,
            'MS365-Schul-Tools – Schülervertretung'
        );
        // Schülervertretung: Team-Provision überspringen (kein explizites Team-Zielbild in den aktuellen Anforderungen)
        const match =
            created && created.id
                ? { id: created.id, displayName: created.displayName || expectedDn, mailNickname: created.mailNickname || expectedNick }
                : null;
        ctx.patchMatchedGroupId('studentCouncil', match || created);
        ctx.setSingleGroupMatchStatus(ctx.studentCouncilGroupMatchCell, { found: true, group: match || created }, expectedNick);
        ctx.setSummary(
            'Schülervertretungs-Gruppe angelegt: „' +
                ctx.readStudentCouncilGroupDisplayName() +
                '“ (' +
                ctx.readStudentCouncilGroupMailNickname() +
                '@…).',
            'ok'
        );
        ctx.renderStatusOverview();
        return created;
    }
}
