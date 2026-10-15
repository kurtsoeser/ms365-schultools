/**
 * Abgleich Organigramm-Landkarte ↔ Microsoft 365
 * Vorschau + Ausführung (Graph für Gruppen/Mitglieder, PowerShell/.cmd für Postfächer).
 */
import {
    isAreaNode,
    isArgeLeiterArea,
    normalizeSchulOrganigramm,
    schulOrgGroupKindLabel,
    suggestSchulOrgMailNickname,
    collectArgeLeiterPeopleFromSettings,
    listArgeCatalogGroupRefs,
    SCHUL_ORG_SEC_ARGE_LEITER
} from './schul-organigramm-logic.js';
import { listM365PeopleForArea } from './schul-organigramm-edit.js';
import {
    areaNeedsM365Provision,
    provisionJobFromArea,
    provisionSchulOrgAreaJob
} from './schul-organigramm-provision-diff.js';
import {
    buildSharedMailboxCreateCmd,
    buildSharedMailboxCreateScript
} from '../../shared/exchange-shared-mailbox-script.js';

function normStr(v) {
    return String(v ?? '').trim();
}

function normEmail(v) {
    return normStr(v).toLowerCase();
}

/**
 * @param {object} personOrUser
 */
function emailFromGraphUser(personOrUser) {
    const u = personOrUser || {};
    return normEmail(u.mail || u.userPrincipalName || u.email || u.upn || '');
}

/**
 * @param {object[]} people listM365PeopleForArea
 * @returns {string[]}
 */
export function emailsFromAreaPeople(people) {
    const seen = new Set();
    /** @type {string[]} */
    const out = [];
    (Array.isArray(people) ? people : []).forEach(function (p) {
        const em = normEmail(p && p.email);
        if (!em || seen.has(em)) return;
        seen.add(em);
        out.push(em);
    });
    return out;
}

/**
 * Mitglieder-Diff: gewünschte Organigramm-Mails vs. aktuelle Graph-Mitglieder.
 * Nur Ergänzungen (kein Entfernen – sicherer Default).
 * @param {string[]} desiredEmails
 * @param {object[]} currentMembers Graph users
 */
export function diffGroupMembers(desiredEmails, currentMembers) {
    const desired = emailsFromAreaPeople(
        (Array.isArray(desiredEmails) ? desiredEmails : []).map(function (e) {
            return { email: e };
        })
    );
    /** @type {Set<string>} */
    const have = new Set();
    (Array.isArray(currentMembers) ? currentMembers : []).forEach(function (u) {
        const em = emailFromGraphUser(u);
        if (em) have.add(em);
    });
    const toAdd = desired.filter(function (e) {
        return !have.has(e);
    });
    const already = desired.filter(function (e) {
        return have.has(e);
    });
    return {
        desired: desired,
        toAdd: toAdd,
        already: already,
        currentCount: have.size
    };
}

/**
 * @param {object} settings
 * @param {object} areaNode
 */
function peopleEmailsForArea(settings, areaNode) {
    if (isArgeLeiterArea(areaNode)) {
        const fromHeads = collectArgeLeiterPeopleFromSettings(settings).map(function (p) {
            return p.email;
        });
        return emailsFromAreaPeople(
            fromHeads.map(function (e) {
                return { email: e };
            })
        );
    }
    return emailsFromAreaPeople(listM365PeopleForArea(settings, areaNode));
}

function getSetupFromRuntime(settings) {
    try {
        if (settings && settings.setup && typeof settings.setup === 'object') return settings.setup;
        if (
            typeof window !== 'undefined' &&
            window.ms365AppDataV2 &&
            typeof window.ms365AppDataV2.getContainer === 'function'
        ) {
            const c = window.ms365AppDataV2.getContainer();
            if (c && c.setup) return c.setup;
        }
    } catch {
        /* ignore */
    }
    return {};
}

/**
 * ARGE-Leiter-Soll: Stammdaten-Köpfe + optional Graph-Owner der ARGE-Gruppen.
 * @param {object} settings
 * @param {{
 *   token?: string,
 *   setup?: object,
 *   fetchGroupOwners?: (token: string, groupId: string) => Promise<{ items?: object[] }|object[]>
 * }} [deps]
 */
export async function resolveArgeLeiterDesiredEmails(settings, deps) {
    const d = deps || {};
    const base = peopleEmailsForArea(settings, {
        id: SCHUL_ORG_SEC_ARGE_LEITER,
        type: 'area',
        refKind: 'argeLeiter'
    });
    const seen = new Set(base);
    /** @type {string[]} */
    const out = base.slice();
    if (!d.token || typeof d.fetchGroupOwners !== 'function') {
        return { emails: out, fromHeads: base.length, fromOwners: 0 };
    }
    const setup =
        d.setup && typeof d.setup === 'object' ? d.setup : getSetupFromRuntime(settings);
    const refs = listArgeCatalogGroupRefs(setup);
    let fromOwners = 0;
    for (let i = 0; i < refs.length; i++) {
        try {
            const res = await d.fetchGroupOwners(d.token, refs[i].groupId);
            const items = res && Array.isArray(res.items) ? res.items : Array.isArray(res) ? res : [];
            items.forEach(function (u) {
                const em = emailFromGraphUser(u);
                if (!em || seen.has(em)) return;
                seen.add(em);
                out.push(em);
                fromOwners += 1;
            });
        } catch {
            /* einzelne Gruppe überspringen */
        }
    }
    return { emails: out, fromHeads: base.length, fromOwners: fromOwners };
}

/**
 * Statische Vorschau-Kandidaten (ohne Graph-Live-Daten).
 * @param {object} settings
 * @returns {{
 *   actions: Array<object>,
 *   counts: object
 * }}
 */
export function buildSchulOrgM365AbgleichDraft(settings) {
    const s = settings && typeof settings === 'object' ? settings : {};
    const layout = normalizeSchulOrganigramm(s.schulOrganigramm);
    /** @type {object[]} */
    const actions = [];

    layout.nodes.forEach(function (n) {
        if (!isAreaNode(n)) return;
        const m365 = n.m365 && typeof n.m365 === 'object' ? n.m365 : {};
        const peopleEmails = peopleEmailsForArea(s, n);
        const ownerEmail = peopleEmails[0] || '';
        const argeLeiterNeedsOwners = isArgeLeiterArea(n);

        if (areaNeedsM365Provision(n)) {
            const job = provisionJobFromArea(n);
            actions.push({
                id: 'createGroup:' + n.id,
                type: 'createGroup',
                channel: 'graph',
                nodeId: n.id,
                label: n.label || job.label,
                title: 'Gruppe anlegen',
                detail:
                    schulOrgGroupKindLabel(job.groupKind) +
                    ' · ' +
                    job.mailNickname +
                    (peopleEmails.length
                        ? ' · ' + peopleEmails.length + ' Mitglied(er)'
                        : '') +
                    (argeLeiterNeedsOwners ? ' · Owner der ARGEs werden nachgeladen' : ''),
                groupKind: job.groupKind,
                groupKindLabel: job.groupKindLabel,
                displayName: job.displayName,
                mailNickname: job.mailNickname,
                memberEmails: peopleEmails,
                ownerEmail: ownerEmail,
                selected: true,
                status: argeLeiterNeedsOwners ? 'needsGraph' : 'pending',
                needsGraph: argeLeiterNeedsOwners
            });
        } else if (normStr(m365.groupId)) {
            actions.push({
                id: 'syncMembers:' + n.id,
                type: 'syncMembers',
                channel: 'graph',
                nodeId: n.id,
                label: n.label || 'Bereich',
                title: 'Mitglieder abgleichen',
                detail:
                    'Gruppe ' +
                    (normStr(m365.groupLabel) || normStr(m365.groupMail) || m365.groupId) +
                    (peopleEmails.length
                        ? ' · Soll: ' + peopleEmails.length + ' Person(en)'
                        : ' · keine Organigramm-Personen'),
                groupId: normStr(m365.groupId),
                groupLabel: normStr(m365.groupLabel) || '',
                memberEmails: peopleEmails,
                ownerEmail: ownerEmail,
                toAdd: peopleEmails.slice(),
                already: [],
                selected: peopleEmails.length > 0,
                status: 'needsGraph',
                needsGraph: true
            });
        }

        const mailMode = normStr(m365.mailboxMode).toLowerCase();
        const mailUpn = normEmail(m365.mailboxUpn);
        if (mailMode === 'create' && mailUpn) {
            const displayName = normStr(m365.mailboxLabel) || normStr(n.label) || 'Postfach';
            const alias =
                suggestSchulOrgMailNickname(mailUpn.split('@')[0] || displayName) || 'postfach';
            actions.push({
                id: 'createMailbox:' + n.id,
                type: 'createMailbox',
                channel: 'powershell',
                nodeId: n.id,
                label: n.label || displayName,
                title: 'Freigegebenes Postfach (PowerShell)',
                detail:
                    displayName +
                    ' · ' +
                    mailUpn +
                    (peopleEmails.length
                        ? ' · Rechte für ' + peopleEmails.length + ' Person(en)'
                        : ''),
                mailbox: {
                    displayName: displayName,
                    alias: alias,
                    primarySmtp: mailUpn,
                    delegates: peopleEmails
                },
                memberEmails: peopleEmails,
                selected: true,
                status: 'pending'
            });
        } else if ((mailMode === 'match' || mailMode === 'create') && mailUpn && peopleEmails.length) {
            actions.push({
                id: 'mailboxPermissions:' + n.id,
                type: 'mailboxPermissions',
                channel: 'powershell',
                nodeId: n.id,
                label: n.label || mailUpn,
                title: 'Postfach-Rechte setzen (PowerShell)',
                detail:
                    mailUpn +
                    ' · FullAccess/SendAs für ' +
                    peopleEmails.length +
                    ' Person(en)',
                mailbox: {
                    displayName: normStr(m365.mailboxLabel) || mailUpn,
                    alias: suggestSchulOrgMailNickname(mailUpn.split('@')[0] || 'postfach'),
                    primarySmtp: mailUpn,
                    delegates: peopleEmails,
                    permissionsOnly: true
                },
                memberEmails: peopleEmails,
                selected: true,
                status: 'pending'
            });
        }
    });

    actions.sort(function (a, b) {
        const order = { createGroup: 1, syncMembers: 2, createMailbox: 3, mailboxPermissions: 4 };
        const oa = order[a.type] || 9;
        const ob = order[b.type] || 9;
        if (oa !== ob) return oa - ob;
        return String(a.label || '').localeCompare(String(b.label || ''), 'de');
    });

    return {
        actions: actions,
        counts: countAbgleichActions(actions)
    };
}

/**
 * @param {object[]} actions
 */
export function countAbgleichActions(actions) {
    const list = Array.isArray(actions) ? actions : [];
    return {
        total: list.length,
        createGroup: list.filter(function (a) {
            return a.type === 'createGroup';
        }).length,
        syncMembers: list.filter(function (a) {
            return a.type === 'syncMembers';
        }).length,
        createMailbox: list.filter(function (a) {
            return a.type === 'createMailbox';
        }).length,
        mailboxPermissions: list.filter(function (a) {
            return a.type === 'mailboxPermissions';
        }).length,
        graph: list.filter(function (a) {
            return a.channel === 'graph';
        }).length,
        powershell: list.filter(function (a) {
            return a.channel === 'powershell';
        }).length,
        selected: list.filter(function (a) {
            return a.selected;
        }).length,
        ok: list.filter(function (a) {
            return a.status === 'noop';
        }).length,
        pending: list.filter(function (a) {
            return a.status === 'pending' || a.status === 'needsGraph';
        }).length,
        error: list.filter(function (a) {
            return a.status === 'error';
        }).length
    };
}

/**
 * Vergleich Landkarte ↔ M365 für die UI (Abschnitte + Kurzstatistik).
 * @param {object[]} actions
 */
export function summarizeAbgleichCompare(actions) {
    const list = Array.isArray(actions) ? actions : [];
    /** @type {object[]} */
    const ok = [];
    /** @type {object[]} */
    const missing = [];
    /** @type {object[]} */
    const create = [];
    /** @type {object[]} */
    const mailbox = [];
    /** @type {object[]} */
    const errors = [];
    /** @type {object[]} */
    const pendingGraph = [];

    list.forEach(function (a) {
        if (!a) return;
        if (a.status === 'error') {
            errors.push(a);
            return;
        }
        if (a.type === 'createGroup') {
            create.push(a);
            return;
        }
        if (a.type === 'createMailbox' || a.type === 'mailboxPermissions') {
            mailbox.push(a);
            return;
        }
        if (a.type === 'syncMembers') {
            if (a.status === 'needsGraph') {
                pendingGraph.push(a);
                return;
            }
            if (a.status === 'noop' || (!(a.toAdd || []).length && a.status !== 'pending')) {
                ok.push(a);
                return;
            }
            missing.push(a);
            return;
        }
        if (a.status === 'noop') ok.push(a);
        else missing.push(a);
    });

    return {
        ok: ok,
        missing: missing,
        create: create,
        mailbox: mailbox,
        errors: errors,
        pendingGraph: pendingGraph,
        counts: {
            ok: ok.length,
            missing: missing.length,
            create: create.length,
            mailbox: mailbox.length,
            errors: errors.length,
            pendingGraph: pendingGraph.length,
            total: list.length
        }
    };
}

/**
 * Draft mit Live-Graph anreichern (Mitgliederstände).
 * @param {ReturnType<typeof buildSchulOrgM365AbgleichDraft>} draft
 * @param {{
 *   token: string,
 *   fetchGroupMembers: (token: string, groupId: string) => Promise<{ items?: object[] }>
 * }} deps
 * @param {(ev: object) => void} [onProgress]
 */
export async function enrichSchulOrgM365AbgleichPlan(draft, deps, onProgress) {
    const d = deps || {};
    const base = draft && typeof draft === 'object' ? draft : { actions: [] };
    const actions = (Array.isArray(base.actions) ? base.actions : []).map(function (a) {
        return Object.assign({}, a);
    });
    if (!d.token || typeof d.fetchGroupMembers !== 'function') {
        return { actions: actions, counts: countAbgleichActions(actions), enriched: false };
    }

    // ARGE-Leiter: Soll-Mitglieder = Stammdaten-Köpfe + Owner der ARGE-Teams
    for (let j = 0; j < actions.length; j++) {
        const act = actions[j];
        if (
            !act ||
            (act.type !== 'syncMembers' && act.type !== 'createGroup') ||
            (act.nodeId !== SCHUL_ORG_SEC_ARGE_LEITER &&
                !/arge[\s_-]*leiter/i.test(String(act.label || '')))
        ) {
            continue;
        }
        try {
            const resolved = await resolveArgeLeiterDesiredEmails(d.settings || {}, {
                token: d.token,
                fetchGroupOwners:
                    typeof d.fetchGroupOwners === 'function' ? d.fetchGroupOwners : null,
                setup: d.setup
            });
            act.memberEmails = resolved.emails;
            act.toAdd = resolved.emails.slice();
            act.detail =
                (act.label || 'ARGE-Leiter') +
                ' · Soll ' +
                resolved.emails.length +
                ' (Köpfe ' +
                resolved.fromHeads +
                (resolved.fromOwners ? ' + Owner +' + resolved.fromOwners : '') +
                ')';
        } catch {
            /* Stammdaten-Köpfe bleiben */
        }
    }

    const syncJobs = actions.filter(function (a) {
        return a.type === 'syncMembers' && a.groupId;
    });
    let done = 0;
    for (let i = 0; i < actions.length; i++) {
        const a = actions[i];
        if (a.type !== 'syncMembers' || !a.groupId) continue;
        done += 1;
        if (typeof onProgress === 'function') {
            onProgress({
                done: done,
                total: syncJobs.length,
                index: done - 1,
                action: a,
                label: a.label || a.groupLabel || a.groupId
            });
        }
        try {
            const res = await d.fetchGroupMembers(d.token, a.groupId);
            const items = res && Array.isArray(res.items) ? res.items : Array.isArray(res) ? res : [];
            const diff = diffGroupMembers(a.memberEmails || [], items);
            a.toAdd = diff.toAdd;
            a.already = diff.already;
            a.desired = diff.desired;
            a.currentCount = diff.currentCount;
            a.needsGraph = false;
            a.status = diff.toAdd.length ? 'pending' : 'noop';
            a.selected = diff.toAdd.length > 0;
            a.detail =
                (a.groupLabel || a.groupId) +
                ' · Landkarte ' +
                diff.desired.length +
                ' · in M365 ' +
                diff.currentCount +
                ' · fehlt +' +
                diff.toAdd.length +
                (diff.already.length ? ' · passt ' + diff.already.length : '');
            if (!diff.toAdd.length) {
                a.detail =
                    (a.groupLabel || a.groupId) +
                    ' · Mitglieder synchron (Landkarte ' +
                    diff.desired.length +
                    ' · M365 ' +
                    diff.currentCount +
                    ')';
            }
        } catch (e) {
            a.status = 'error';
            a.selected = false;
            a.error = e && e.message ? String(e.message) : String(e);
            a.detail = 'Graph-Fehler: ' + a.error;
        }
    }

    return {
        actions: actions,
        counts: countAbgleichActions(actions),
        compare: summarizeAbgleichCompare(actions),
        enriched: true
    };
}

/**
 * @param {object[]} mailboxActions createMailbox | mailboxPermissions
 * @param {string} [suggestedAdminUpn]
 * @returns {{ script: string, cmd: string, count: number }}
 */
export function buildAbgleichMailboxLauncher(mailboxActions, suggestedAdminUpn) {
    const list = Array.isArray(mailboxActions) ? mailboxActions : [];
    if (!list.length) {
        return { script: '', cmd: '', count: 0 };
    }
    // Ein Postfach: bestehende Builder nutzen
    if (list.length === 1) {
        const mb = list[0].mailbox || {};
        const opts = {
            displayName: mb.displayName || list[0].label,
            alias: mb.alias || suggestSchulOrgMailNickname(mb.primarySmtp || 'postfach'),
            primarySmtp: mb.primarySmtp,
            description: 'Schul-Organigramm Abgleich: ' + (mb.displayName || list[0].label),
            delegates: mb.delegates || list[0].memberEmails || [],
            rights: ['FullAccess', 'SendAs'],
            autoMapping: true,
            suggestedAdminUpn: suggestedAdminUpn || ''
        };
        if (list[0].type === 'mailboxPermissions' || mb.permissionsOnly) {
            // Rechte-only: Create-Script würde New-Mailbox versuchen → eigene Variante
            const script = buildMailboxPermissionGuidedScript(opts);
            const cmd = wrapPsInCmd(script);
            return { script: script, cmd: cmd, count: 1 };
        }
        return {
            script: buildSharedMailboxCreateScript(opts),
            cmd: buildSharedMailboxCreateCmd(opts),
            count: 1
        };
    }

    // Mehrere: geführter Batch (ein Connect, nacheinander)
    const batchItems = list.map(function (a) {
        const mb = a.mailbox || {};
        return {
            type: a.type,
            displayName: mb.displayName || a.label,
            alias: mb.alias || suggestSchulOrgMailNickname(mb.primarySmtp || 'postfach'),
            primarySmtp: mb.primarySmtp,
            delegates: mb.delegates || a.memberEmails || [],
            permissionsOnly: a.type === 'mailboxPermissions' || !!mb.permissionsOnly
        };
    });
    const script = buildMailboxBatchGuidedScript(batchItems, suggestedAdminUpn || '');
    return { script: script, cmd: wrapPsInCmd(script), count: batchItems.length };
}

/**
 * @param {{
 *   displayName: string,
 *   alias: string,
 *   primarySmtp: string,
 *   delegates: string[],
 *   suggestedAdminUpn?: string
 * }} opts
 */
function buildMailboxPermissionGuidedScript(opts) {
    const o = opts || {};
    // Nutzt denselben Assistenten-Rahmen über Create-Script, aber wir erzeugen
    // ein schlankes Rechte-only Skript inline.
    const lines = [];
    const q = function (v) {
        return "'" + String(v == null ? '' : v).replace(/'/g, "''") + "'";
    };
    lines.push('#Requires -Version 5.1');
    lines.push('# Postfach-Rechte setzen (Abgleich) – MS365-Schul-Tools');
    lines.push('$ErrorActionPreference = "Stop"');
    lines.push('try { [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false) } catch {}');
    lines.push('$Host.UI.RawUI.WindowTitle = "Postfach-Rechte Abgleich"');
    lines.push('function Pause-End { Read-Host "Eingabetaste zum Beenden" | Out-Null }');
    lines.push('$MailboxIdentity = ' + q(o.primarySmtp));
    lines.push(
        '$Delegates = @(' +
            (o.delegates || [])
                .map(function (e) {
                    return q(e);
                })
                .join(', ') +
            ')'
    );
    lines.push('$SuggestedAdminUpn = ' + q(o.suggestedAdminUpn || ''));
    lines.push('Write-Host "Postfach-Rechte: $MailboxIdentity" -ForegroundColor Cyan');
    lines.push('Write-Host ("Berechtigte: " + ($Delegates -join ", "))');
    lines.push(
        '$AdminUpn = Read-Host "Admin-UPN (leer = interaktiv' +
            (o.suggestedAdminUpn ? ', Enter = ' + o.suggestedAdminUpn : '') +
            ')"'
    );
    lines.push(
        'if ([string]::IsNullOrWhiteSpace($AdminUpn) -and $SuggestedAdminUpn) { $AdminUpn = $SuggestedAdminUpn }'
    );
    lines.push('$c = Read-Host "Rechte jetzt setzen? (J/N)"');
    lines.push('if ($c -notmatch "^[jJyY]") { Write-Host "Abbruch."; Pause-End; return }');
    lines.push(
        'if (-not (Get-Module -ListAvailable ExchangeOnlineManagement)) { Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber }'
    );
    lines.push('Import-Module ExchangeOnlineManagement');
    lines.push(
        'if ($AdminUpn) { Connect-ExchangeOnline -UserPrincipalName $AdminUpn -ShowBanner:$false | Out-Null } else { Connect-ExchangeOnline -ShowBanner:$false | Out-Null }'
    );
    lines.push('try {');
    lines.push('  foreach ($d in $Delegates) {');
    lines.push(
        '    Add-MailboxPermission -Identity $MailboxIdentity -User $d -AccessRights FullAccess -InheritanceType All -AutoMapping:$true -ErrorAction Continue | Out-Null'
    );
    lines.push(
        '    Add-RecipientPermission -Identity $MailboxIdentity -Trustee $d -AccessRights SendAs -Confirm:$false -ErrorAction Continue | Out-Null'
    );
    lines.push('    Write-Host ("OK " + $d) -ForegroundColor Green');
    lines.push('  }');
    lines.push(
        '  Get-MailboxPermission -Identity $MailboxIdentity | Where-Object { $_.IsInherited -eq $false -and $_.User -notlike "NT AUTHORITY*" } | Format-Table -AutoSize'
    );
    lines.push('} finally {');
    lines.push('  try { Disconnect-ExchangeOnline -Confirm:$false | Out-Null } catch {}');
    lines.push('  Pause-End');
    lines.push('}');
    return lines.join('\n');
}

/**
 * @param {Array<{
 *   type: string,
 *   displayName: string,
 *   alias: string,
 *   primarySmtp: string,
 *   delegates: string[],
 *   permissionsOnly?: boolean
 * }>} items
 * @param {string} suggestedAdminUpn
 */
function buildMailboxBatchGuidedScript(items, suggestedAdminUpn) {
    const q = function (v) {
        return "'" + String(v == null ? '' : v).replace(/'/g, "''") + "'";
    };
    const lines = [];
    lines.push('#Requires -Version 5.1');
    lines.push('# Batch Postfächer / Rechte – MS365-Schul-Tools Abgleich');
    lines.push('$ErrorActionPreference = "Stop"');
    lines.push('try { [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false) } catch {}');
    lines.push('$Host.UI.RawUI.WindowTitle = "Abgleich Postfächer"');
    lines.push('function Pause-End { Read-Host "Eingabetaste zum Beenden" | Out-Null }');
    lines.push('$SuggestedAdminUpn = ' + q(suggestedAdminUpn || ''));
    lines.push('$Items = @(');
    items.forEach(function (it, idx) {
        lines.push('  [pscustomobject]@{');
        lines.push('    DisplayName = ' + q(it.displayName));
        lines.push('    Alias = ' + q(it.alias));
        lines.push('    PrimarySmtp = ' + q(it.primarySmtp));
        lines.push('    PermissionsOnly = $' + (it.permissionsOnly ? 'true' : 'false'));
        lines.push(
            '    Delegates = @(' +
                (it.delegates || [])
                    .map(function (e) {
                        return q(e);
                    })
                    .join(', ') +
                ')'
        );
        lines.push('  }' + (idx < items.length - 1 ? ',' : ''));
    });
    lines.push(')');
    lines.push('Write-Host ("Geplante Postfach-Aktionen: " + $Items.Count) -ForegroundColor Cyan');
    lines.push('$Items | ForEach-Object { Write-Host (" - " + $_.DisplayName + " <" + $_.PrimarySmtp + ">" + $(if ($_.PermissionsOnly) { " [nur Rechte]" } else { " [anlegen]" })) }');
    lines.push(
        '$AdminUpn = Read-Host "Admin-UPN der richtigen Schule/Tenant (leer = interaktiv' +
            (suggestedAdminUpn ? ', Enter = ' + suggestedAdminUpn : '') +
            ')"'
    );
    lines.push(
        'if ([string]::IsNullOrWhiteSpace($AdminUpn) -and $SuggestedAdminUpn) { $AdminUpn = $SuggestedAdminUpn }'
    );
    lines.push('$c = Read-Host "Jetzt ausführen? (J/N)"');
    lines.push('if ($c -notmatch "^[jJyY]") { Write-Host "Abbruch."; Pause-End; return }');
    lines.push(
        'if (-not (Get-Module -ListAvailable ExchangeOnlineManagement)) { Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber }'
    );
    lines.push('Import-Module ExchangeOnlineManagement');
    lines.push(
        'if ($AdminUpn) { Connect-ExchangeOnline -UserPrincipalName $AdminUpn -ShowBanner:$false | Out-Null } else { Connect-ExchangeOnline -ShowBanner:$false | Out-Null }'
    );
    lines.push('try {');
    lines.push('  foreach ($it in $Items) {');
    lines.push('    Write-Host ""');
    lines.push('    Write-Host ("=== " + $it.DisplayName + " ===") -ForegroundColor Cyan');
    lines.push('    $id = $it.PrimarySmtp');
    lines.push('    if (-not $it.PermissionsOnly) {');
    lines.push('      $params = @{ Shared = $true; Name = $it.DisplayName; DisplayName = $it.DisplayName; Alias = $it.Alias; PrimarySmtpAddress = $it.PrimarySmtp }');
    lines.push('      New-Mailbox @params | Out-Null');
    lines.push('      $mb = $null');
    lines.push('      for ($i = 1; $i -le 24; $i++) {');
    lines.push('        try { $mb = Get-Mailbox -Identity $id -ErrorAction Stop; break } catch { Start-Sleep -Seconds 5 }');
    lines.push('      }');
    lines.push('      if (-not $mb) { Write-Host "Noch nicht abrufbar – Rechte ggf. später." -ForegroundColor Yellow; continue }');
    lines.push('      Write-Host "Postfach angelegt." -ForegroundColor Green');
    lines.push('    }');
    lines.push('    foreach ($d in $it.Delegates) {');
    lines.push(
        '      Add-MailboxPermission -Identity $id -User $d -AccessRights FullAccess -InheritanceType All -AutoMapping:$true -ErrorAction Continue | Out-Null'
    );
    lines.push(
        '      Add-RecipientPermission -Identity $id -Trustee $d -AccessRights SendAs -Confirm:$false -ErrorAction Continue | Out-Null'
    );
    lines.push('      Write-Host ("  Rechte → " + $d) -ForegroundColor Green');
    lines.push('    }');
    lines.push('  }');
    lines.push('  Write-Host "Batch fertig." -ForegroundColor Green');
    lines.push('} finally {');
    lines.push('  try { Disconnect-ExchangeOnline -Confirm:$false | Out-Null } catch {}');
    lines.push('  Pause-End');
    lines.push('}');
    return lines.join('\n');
}

const CMD_PS_MARKER = '___MS365_ABGLEICH_MAILBOX_PS___';

/**
 * @param {string} ps
 */
function wrapPsInCmd(ps) {
    const withEnv =
        '@echo off\r\n' +
        'setlocal EnableExtensions\r\n' +
        'chcp 65001 >nul\r\n' +
        'title MS365 Abgleich Postfaecher\r\n' +
        'cd /d "%~dp0"\r\n' +
        'set "MS365_CMD_PATH=%~f0"\r\n' +
        'echo.\r\n' +
        'echo  Starte gefuehrten Postfach-Abgleich ...\r\n' +
        'echo.\r\n' +
        'powershell.exe -NoProfile -ExecutionPolicy Bypass -NoLogo -Command ^\r\n' +
        '  "$ErrorActionPreference=\'Stop\';' +
        ' $path=$env:MS365_CMD_PATH;' +
        ' $lines=Get-Content -LiteralPath $path -Encoding UTF8;' +
        " $marker='" +
        CMD_PS_MARKER +
        "';" +
        ' $i=0; while($i -lt $lines.Count -and $lines[$i].Trim() -ne $marker){ $i++ };' +
        ' if($i -ge $lines.Count){ throw \'PS-Teil fehlt.\' };' +
        ' $script=($lines[($i+1)..($lines.Count-1)] -join [Environment]::NewLine);' +
        " $ps1=Join-Path $env:TEMP ('ms365-abgleich-mbx-' + [guid]::NewGuid().ToString('N') + '.ps1');" +
        ' Set-Content -LiteralPath $ps1 -Value $script -Encoding UTF8;' +
        ' try { & powershell.exe -NoProfile -ExecutionPolicy Bypass -NoLogo -File $ps1; exit $LASTEXITCODE }' +
        ' finally { Remove-Item -LiteralPath $ps1 -Force -ErrorAction SilentlyContinue }"\r\n' +
        'set "EC=%ERRORLEVEL%"\r\n' +
        'endlocal & exit /b %EC%\r\n' +
        '\r\n' +
        CMD_PS_MARKER +
        '\r\n';
    return withEnv + String(ps || '').replace(/\r?\n/g, '\r\n') + '\r\n';
}

/**
 * Graph-Aktionen aus dem Plan ausführen.
 * @param {object} settings
 * @param {object[]} actions selected actions
 * @param {object} deps
 * @param {(ev: object) => void} [onProgress]
 */
export async function executeSchulOrgGraphAbgleich(settings, actions, deps, onProgress) {
    const d = deps || {};
    let layout = normalizeSchulOrganigramm(settings && settings.schulOrganigramm);
    const list = (Array.isArray(actions) ? actions : []).filter(function (a) {
        return a && a.selected && a.channel === 'graph' && a.status !== 'noop' && a.status !== 'error';
    });
    /** @type {object[]} */
    const results = [];

    for (let i = 0; i < list.length; i++) {
        const a = list[i];
        try {
            if (a.type === 'createGroup') {
                if (typeof d.createGroupByKind !== 'function') {
                    throw new Error('createGroupByKind fehlt');
                }
                const done = await provisionSchulOrgAreaJob(
                    layout,
                    {
                        nodeId: a.nodeId,
                        label: a.label,
                        groupKind: a.groupKind || 'unified',
                        displayName: a.displayName || a.label,
                        mailNickname: a.mailNickname,
                        memberEmails: a.memberEmails || [],
                        ownerEmail: a.ownerEmail || ''
                    },
                    d
                );
                layout = done.layout;
                results.push({ id: a.id, ok: true, type: a.type, groupId: done.groupId });
                if (typeof onProgress === 'function') {
                    onProgress({
                        index: i,
                        total: list.length,
                        action: a,
                        ok: true,
                        message: 'Gruppe angelegt: ' + done.groupId
                    });
                }
            } else if (a.type === 'syncMembers') {
                if (typeof d.syncEmailsToGroup !== 'function') {
                    throw new Error('syncEmailsToGroup fehlt');
                }
                const toAdd = Array.isArray(a.toAdd) ? a.toAdd : a.memberEmails || [];
                if (!toAdd.length) {
                    results.push({ id: a.id, ok: true, type: a.type, skipped: true });
                    if (typeof onProgress === 'function') {
                        onProgress({
                            index: i,
                            total: list.length,
                            action: a,
                            ok: true,
                            message: 'Keine neuen Mitglieder'
                        });
                    }
                    continue;
                }
                const sync = await d.syncEmailsToGroup(
                    d.token,
                    a.groupId,
                    toAdd,
                    a.label || a.groupLabel || ''
                );
                results.push({
                    id: a.id,
                    ok: true,
                    type: a.type,
                    added: (sync && sync.ok) || 0,
                    fail: (sync && sync.fail) || 0
                });
                if (typeof onProgress === 'function') {
                    onProgress({
                        index: i,
                        total: list.length,
                        action: a,
                        ok: true,
                        message:
                            '+' +
                            ((sync && sync.ok) || 0) +
                            ' Mitglied(er)' +
                            (sync && sync.fail ? ', ' + sync.fail + ' Fehler' : '')
                    });
                }
            } else {
                results.push({ id: a.id, ok: false, error: 'Unbekannter Typ' });
            }
        } catch (e) {
            const msg = e && e.message ? String(e.message) : String(e);
            results.push({ id: a.id, ok: false, type: a.type, error: msg });
            if (typeof onProgress === 'function') {
                onProgress({
                    index: i,
                    total: list.length,
                    action: a,
                    ok: false,
                    message: msg
                });
            }
        }
    }

    const okCount = results.filter(function (r) {
        return r.ok;
    }).length;
    return {
        layout: layout,
        results: results,
        okCount: okCount,
        failCount: results.length - okCount
    };
}

export default {
    emailsFromAreaPeople,
    diffGroupMembers,
    resolveArgeLeiterDesiredEmails,
    buildSchulOrgM365AbgleichDraft,
    enrichSchulOrgM365AbgleichPlan,
    countAbgleichActions,
    summarizeAbgleichCompare,
    buildAbgleichMailboxLauncher,
    executeSchulOrgGraphAbgleich
};
