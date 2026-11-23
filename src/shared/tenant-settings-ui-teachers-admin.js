/**
 * tenant-settings-ui-teachers-admin.js
 * @param {Record<string, any>} ctx
 */
export function installTenantTeachersAdmin(ctx) {
    ctx.teachersToLines = function teachersToLines(rows) {
        return (rows || [])
            .map((x) => `${ctx.normCode(x.code)};${ctx.normStr(x.name || '')};${ctx.normStr(x.email || '').toLowerCase()}`.trim())
            .filter(Boolean)
            .join('\n');
    }

    ctx.adminToLines = function adminToLines(rows) {
        return (rows || [])
            .map((x) => `${ctx.normStr(x.role || '')};${ctx.normStr(x.name || '')};${ctx.normStr(x.email || '').toLowerCase()}`.trim())
            .filter(Boolean)
            .join('\n');
    }

    ctx.adminRolesToLines = function adminRolesToLines(rows) {
        return (rows || [])
            .map((x) => `${ctx.normCode(x.code || '')};${ctx.normStr(x.name || '')}`.trim())
            .filter(Boolean)
            .join('\n');
    }

    /**
     * seatMode / M365-Bezug aus dem strukturierten Snapshot in geparste Gruppen übernehmen
     * (Textzeilen kennen diese Felder nicht).
     */
    ctx.mergeAdminRolePolicyFromSnapshot = function mergeAdminRolePolicyFromSnapshot(groups, snapshot) {
        const prev = Array.isArray(snapshot) ? snapshot : [];
        const byKey = new Map();
        prev.forEach(function (g) {
            const k = (ctx.normStr(g && g.code) || ctx.normStr(g && g.name)).toLowerCase();
            if (k) byKey.set(k, g);
        });
        return (Array.isArray(groups) ? groups : []).map(function (group) {
            const k = (ctx.normCode(group && group.code) || ctx.normStr(group && group.name)).toLowerCase();
            const meta = k ? byKey.get(k) : null;
            if (!meta) return group;
            const next = Object.assign({}, group);
            if (meta.seatMode && !next.seatMode) next.seatMode = meta.seatMode;
            if (meta.m365ResourceType && !next.m365ResourceType) next.m365ResourceType = meta.m365ResourceType;
            if (meta.m365ResourceId && !next.m365ResourceId) next.m365ResourceId = meta.m365ResourceId;
            if (meta.m365ResourceEmail && !next.m365ResourceEmail) next.m365ResourceEmail = meta.m365ResourceEmail;
            if (meta.m365ResourceLabel && !next.m365ResourceLabel) next.m365ResourceLabel = meta.m365ResourceLabel;
            if (meta.tier && !next.tier) next.tier = meta.tier;
            return next;
        });
    }

    ctx.getAdministrationGroups = function getAdministrationGroups() {
        if (Array.isArray(ctx.administrationGroupsSnapshot)) {
            return ctx.administrationGroupsSnapshot;
        }
        if (ctx.taAdminBundle && typeof window.ms365TenantSettingsParseAdminGroupsLines === 'function') {
            return window.ms365TenantSettingsParseAdminGroupsLines(ctx.taAdminBundle.value);
        }
        const roles =
            typeof ctx.parseLinesToAdminRoles === 'function'
                ? ctx.parseLinesToAdminRoles(ctx.taAdminRoles ? ctx.taAdminRoles.value : '')
                : [];
        const admin =
            typeof ctx.parseLinesToAdmin === 'function' ? ctx.parseLinesToAdmin(ctx.taAdmin ? ctx.taAdmin.value : '') : [];
        if (typeof window.ms365TenantSettingsAdminRolesAndAdminToGroups === 'function') {
            return window.ms365TenantSettingsAdminRolesAndAdminToGroups(roles, admin);
        }
        return [];
    }

    ctx.setAdministrationGroups = function setAdministrationGroups(groups) {
        ctx.administrationGroupsSnapshot = Array.isArray(groups) ? groups : [];
        if (ctx.taAdminBundle && typeof window.ms365TenantSettingsAdminGroupsToLines === 'function') {
            ctx.taAdminBundle.value = window.ms365TenantSettingsAdminGroupsToLines(ctx.administrationGroupsSnapshot);
            return;
        }
        if (typeof window.ms365TenantSettingsAdminRolesAndAdminToGroups !== 'function') return;
        const roles = (groups || []).map(function (group) {
            const row = {
                code: ctx.normCode(group && group.code),
                name: ctx.normStr(group && group.name),
                tier: ctx.inferAdminTierForRole(group)
            };
            if (group && group.seatMode) row.seatMode = ctx.normStr(group.seatMode);
            if (group && group.m365ResourceType) row.m365ResourceType = ctx.normStr(group.m365ResourceType);
            if (group && group.m365ResourceId) row.m365ResourceId = ctx.normStr(group.m365ResourceId);
            if (group && group.m365ResourceEmail) {
                row.m365ResourceEmail = ctx.normStr(group.m365ResourceEmail).toLowerCase();
            }
            if (group && group.m365ResourceLabel) row.m365ResourceLabel = ctx.normStr(group.m365ResourceLabel);
            return row;
        });
        const admin = [];
        (groups || []).forEach(function (group) {
            const roleName = ctx.normStr(group && group.name);
            (Array.isArray(group && group.people) ? group.people : []).forEach(function (person) {
                admin.push({
                    role: roleName,
                    name: ctx.normStr(person && person.name),
                    email: ctx.normStr(person && person.email).toLowerCase(),
                    defaultKey: ctx.normStr(person && person.defaultKey)
                });
            });
        });
        if (ctx.taAdminRoles) ctx.taAdminRoles.value = ctx.adminRolesToLines(roles);
        if (ctx.taAdmin) ctx.taAdmin.value = ctx.adminToLines(admin);
    }

    ctx.adminTierLabel = function adminTierLabel(tier) {
        const t = ctx.normalizeAdminTier(tier) || ctx.ADMIN_TIER_VERWALTUNG;
        return t === ctx.ADMIN_TIER_SCHULLEITUNG ? 'Schulleitung' : 'Verwaltung (Personal)';
    }

    ctx.syncAdministrationSnapshotFromSettings = function syncAdministrationSnapshotFromSettings(s) {
        ctx.administrationGroupsSnapshot = ctx.adminGroupsFromSettings(s || ctx.load());
    }

    ctx.syncAudienceSnapshotsFromSettings = function syncAudienceSnapshotsFromSettings(s) {
        const ensured = ctx.ensureAdminAudienceOnSettings(s || ctx.load());
        ctx.audienceGroupsSnapshot = ensured.groups;
        ctx.audienceMembershipsSnapshot = ensured.memberships;
    }

    ctx.applyAudienceMembershipPatch = function applyAudienceMembershipPatch(groups, memberships) {
        ctx.audienceGroupsSnapshot = ctx.normalizeVerwaltungAudienceGroups(groups);
        ctx.audienceMembershipsSnapshot = ctx.normalizeAdminAudienceMemberships(memberships);
        if (ctx.audienceBoardRef) ctx.audienceBoardRef.refresh();
    }

    ctx.refreshAdminAudienceBoardUi = function refreshAdminAudienceBoardUi() {
        if (ctx.audienceBoardRef) ctx.audienceBoardRef.refresh();
        if (ctx.verwaltungTreeRef) ctx.verwaltungTreeRef.refresh();
    }

    ctx.openOrganigrammTabFromVerwaltung = function openOrganigrammTabFromVerwaltung() {
        const btn = document.getElementById('tabMainOrganigramm');
        if (btn) btn.click();
    }

    ctx.patchAdministrationDisplayFieldForEmail = function patchAdministrationDisplayFieldForEmail(emailRaw, field, value, context, meta) {
        if (meta && meta.cancelled) {
            ctx.refreshAdminAudienceBoardUi();
            return;
        }
        const em = ctx.normStr(emailRaw).toLowerCase();
        if (!em) return;
        const rows = ctx.groupsToDisplayRows(ctx.getAdministrationGroups());
        let idx = rows.findIndex(function (r) {
            return ctx.normStr(r && r.email).toLowerCase() === em;
        });
        const val = ctx.normStr(value);
        if (idx < 0) {
            const ag = context && context.audienceGroupId ? String(context.audienceGroupId) : '';
            const tier =
                ag === ctx.BUILTIN_AUDIENCE_SCHULLEITUNG || ag === ctx.ADMIN_TIER_SCHULLEITUNG
                    ? ctx.ADMIN_TIER_SCHULLEITUNG
                    : ctx.ADMIN_TIER_VERWALTUNG;
            const row = {
                name: field === 'name' ? val : '',
                code: field === 'code' ? ctx.normCode(val) : '',
                personName: context && context.personName ? ctx.normStr(context.personName) : '',
                email: em,
                tier: tier
            };
            if (field === 'name' && val && !row.code && typeof window.ms365TenantSettingsAdminRoleCodeFromName === 'function') {
                row.code = ctx.normCode(window.ms365TenantSettingsAdminRoleCodeFromName(val));
            }
            rows.push(row);
        } else {
            const prevName = rows[idx].name;
            if (field === 'name') {
                rows[idx].name = val;
                if (rows[idx].code && prevName && val !== prevName && typeof window.ms365TenantSettingsAdminRoleCodeFromName === 'function') {
                    rows[idx].code = ctx.normCode(window.ms365TenantSettingsAdminRoleCodeFromName(val)) || rows[idx].code;
                } else if (!rows[idx].code && val && typeof window.ms365TenantSettingsAdminRoleCodeFromName === 'function') {
                    rows[idx].code = ctx.normCode(window.ms365TenantSettingsAdminRoleCodeFromName(val));
                }
            } else if (field === 'code') {
                rows[idx].code = ctx.normCode(val);
            }
        }
        ctx.setAdministrationGroups(ctx.displayRowsToGroups(rows));
        ctx.scheduleAutoSave();
        ctx.refreshAdminAudienceBoardUi();
    }

    if (ctx.tenantVerwaltungStructureTreeHost) {
        ctx.verwaltungTreeRef = ctx.mountVerwaltungStammdatenTree({
            host: ctx.tenantVerwaltungStructureTreeHost,
            getSettings: function () {
                const base = ctx.load();
                return Object.assign({}, base, {
                    verwaltungAudienceGroups: ctx.audienceGroupsSnapshot,
                    adminAudienceMemberships: ctx.audienceMembershipsSnapshot,
                    schulOrganigramm: base.schulOrganigramm
                });
            },
            saveSettings: function (next) {
                ctx.save(next);
            },
            onOpenOrganigramm: ctx.openOrganigrammTabFromVerwaltung,
            scheduleAutoSave: ctx.scheduleAutoSave,
            getSchoolName: function () {
                return ctx.schoolNameInput
                    ? ctx.normStr(ctx.schoolNameInput.value || '')
                    : ctx.normStr(ctx.load().schoolName || '');
            },
            setSummary: ctx.setSummary,
            onAudienceGroupsChanged: function (groups) {
                ctx.applyAudienceMembershipPatch(groups, ctx.audienceMembershipsSnapshot);
            },
            onStructureChanged: function () {
                ctx.syncAudienceSnapshotsFromSettings(ctx.load());
                ctx.syncAdministrationSnapshotFromSettings(ctx.load());
                ctx.refreshAdminAudienceBoardUi();
            }
        });
        const tabVerwaltung = document.getElementById('tabMainVerwaltung');
        if (tabVerwaltung) {
            tabVerwaltung.addEventListener('click', function () {
                if (ctx.verwaltungTreeRef) ctx.verwaltungTreeRef.refresh();
            });
        }
    }
    // Flaches Zielgruppen-Board nur noch als versteckter Legacy-Host (API für Tests/Refresh)
    if (ctx.tenantAdminAudienceBoardHost && !ctx.tenantAdminAudienceBoardHost.hidden) {
        ctx.audienceBoardRef = ctx.mountTenantAdminAudienceBoard({
            host: ctx.tenantAdminAudienceBoardHost,
            getSettings: function () {
                const base = ctx.load();
                return Object.assign({}, base, {
                    verwaltungAudienceGroups: ctx.audienceGroupsSnapshot,
                    adminAudienceMemberships: ctx.audienceMembershipsSnapshot
                });
            },
            onChange: function (patch) {
                ctx.applyAudienceMembershipPatch(patch.verwaltungAudienceGroups, patch.adminAudienceMemberships);
            },
            scheduleAutoSave: ctx.scheduleAutoSave,
            normStr: ctx.normStr,
            getSchoolName: function () {
                return ctx.schoolNameInput ? ctx.normStr(ctx.schoolNameInput.value || '') : ctx.normStr(ctx.load().schoolName || '');
            },
            setSummary: ctx.setSummary,
            onVerifyDirectory: function (email) {
                return ctx.verifyAdminDirectoryEmail(email);
            },
            onCreateEntraUser: function (email, personName) {
                return ctx.createAdminEntraUserInteractive(email, personName);
            },
            onDirectoryChanged: function () {
                if (ctx.audienceBoardRef) ctx.audienceBoardRef.refresh();
                if (ctx.verwaltungTreeRef) ctx.verwaltungTreeRef.refresh();
            },
            getAdminDisplayRows: function () {
                return ctx.groupsToDisplayRows(ctx.getAdministrationGroups());
            },
            onAdminDisplayFieldEdit: function (email, field, value, ctx, meta) {
                ctx.patchAdministrationDisplayFieldForEmail(email, field, value, ctx, meta);
            }
        });
    }

    ctx.applyAdminTierToDisplayRows = function applyAdminTierToDisplayRows(rows, roleName, tier) {
        const key = ctx.normStr(roleName).toLowerCase();
        if (!key) return;
        const next = ctx.normalizeAdminTier(tier) || ctx.ADMIN_TIER_VERWALTUNG;
        (Array.isArray(rows) ? rows : []).forEach(function (row) {
            if (ctx.normStr(row && row.name).toLowerCase() === key) row.tier = next;
        });
    }

    ctx.groupsToDisplayRows = function groupsToDisplayRows(groups) {
        const rows = [];
        (Array.isArray(groups) ? groups : []).forEach(function (group) {
            const code = ctx.normStr(group && group.code);
            const name = ctx.normStr(group && group.name);
            const tier = ctx.inferAdminTierForRole(group);
            const people = Array.isArray(group && group.people) ? group.people : [];
            if (!people.length) {
                rows.push({ code: code, name: name, personName: '', email: '', tier: tier });
                return;
            }
            people.forEach(function (person) {
                rows.push({
                    code: code,
                    name: name,
                    personName: ctx.normStr(person && person.name),
                    email: ctx.normStr(person && person.email).toLowerCase(),
                    tier: tier
                });
            });
        });
        return rows;
    }

    ctx.displayRowsToGroups = function displayRowsToGroups(rows) {
        const groupMap = new Map();
        const order = [];
        (Array.isArray(rows) ? rows : []).forEach(function (row) {
            const name = ctx.normStr(row && row.name);
            let code = ctx.normCode(row && row.code);
            if (!code && name && typeof window.ms365TenantSettingsAdminRoleCodeFromName === 'function') {
                code = window.ms365TenantSettingsAdminRoleCodeFromName(name);
            }
            const personName = ctx.normStr(row && row.personName);
            const email = ctx.normStr(row && row.email).toLowerCase();
            if (!name && !code && !personName && !email) return;
            const key = (name || code).toLowerCase();
            if (!groupMap.has(key)) {
                groupMap.set(key, {
                    code: code,
                    name: name || code,
                    people: [],
                    tier: ctx.inferAdminTierForRole({ name: name || code, code: code, tier: row && row.tier })
                });
                order.push(key);
            }
            const group = groupMap.get(key);
            if (code) group.code = code;
            if (name) group.name = name;
            if (row && row.tier) {
                group.tier = ctx.inferAdminTierForRole({ name: group.name, code: group.code, tier: row.tier });
            }
            if (personName || email) {
                group.people.push({ name: personName, email: email });
            }
        });
        return order.map(function (key) {
            return groupMap.get(key);
        });
    }

    ctx.adminGroupsFromSettings = function adminGroupsFromSettings(settings) {
        if (
            settings &&
            Array.isArray(settings.administration) &&
            settings.administration.some(function (entry) {
                return entry && Array.isArray(entry.people);
            })
        ) {
            return settings.administration;
        }
        if (typeof window.ms365TenantSettingsAdminRolesAndAdminToGroups === 'function') {
            return window.ms365TenantSettingsAdminRolesAndAdminToGroups(
                settings && settings.adminRoles ? settings.adminRoles : [],
                settings && settings.admin ? settings.admin : []
            );
        }
        return [];
    }

    ctx.getAdministrationEntries = function getAdministrationEntries() {
        return ctx.getAdministrationGroups();
    }

    ctx.getAdminFromTextarea = function getAdminFromTextarea() {
        const out = [];
        ctx.getAdministrationGroups().forEach(function (group) {
            const roleName = ctx.normStr(group && group.name);
            (Array.isArray(group && group.people) ? group.people : []).forEach(function (person) {
                const row = {
                    role: roleName,
                    name: ctx.normStr(person && person.name),
                    email: ctx.normStr(person && person.email).toLowerCase()
                };
                if (person && person.defaultKey) row.defaultKey = ctx.normStr(person.defaultKey);
                out.push(row);
            });
        });
        return out;
    }

    ctx.setAdminTextareaFromRows = function setAdminTextareaFromRows(rows) {
        const roles = ctx.getAdminRolesFromTextarea();
        if (typeof window.ms365TenantSettingsAdminRolesAndAdminToGroups === 'function') {
            ctx.setAdministrationGroups(window.ms365TenantSettingsAdminRolesAndAdminToGroups(roles, rows));
            return;
        }
        if (!ctx.taAdmin) return;
        ctx.taAdmin.value = ctx.adminToLines(rows);
    }

    ctx.getAdminRolesFromTextarea = function getAdminRolesFromTextarea() {
        return ctx.getAdministrationGroups().map(function (group) {
            const row = {
                code: ctx.normCode(group && group.code),
                name: ctx.normStr(group && group.name),
                tier: ctx.inferAdminTierForRole(group)
            };
            if (group && group.seatMode) row.seatMode = ctx.normStr(group.seatMode);
            if (group && group.m365ResourceType) row.m365ResourceType = ctx.normStr(group.m365ResourceType);
            if (group && group.m365ResourceId) row.m365ResourceId = ctx.normStr(group.m365ResourceId);
            if (group && group.m365ResourceEmail) {
                row.m365ResourceEmail = ctx.normStr(group.m365ResourceEmail).toLowerCase();
            }
            if (group && group.m365ResourceLabel) row.m365ResourceLabel = ctx.normStr(group.m365ResourceLabel);
            return row;
        });
    }

    ctx.setAdminRolesTextareaFromRows = function setAdminRolesTextareaFromRows(rows) {
        const people = ctx.getAdminFromTextarea();
        if (typeof window.ms365TenantSettingsAdminRolesAndAdminToGroups === 'function') {
            ctx.setAdministrationGroups(window.ms365TenantSettingsAdminRolesAndAdminToGroups(rows, people));
            return;
        }
        if (ctx.taAdminRoles) ctx.taAdminRoles.value = ctx.adminRolesToLines(rows);
    }

    ctx.syncAdministrationSnapshotFromSettings(ctx.load());
    ctx.syncAudienceSnapshotsFromSettings(ctx.load());

}
