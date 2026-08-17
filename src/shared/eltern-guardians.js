/**
 * Eltern / Erziehungsberechtigte – reine Hilfslogik (Nicknamen, SOLL-Listen, Exchange-PS).
 * Stammdaten liegen in app-data-v2 (years.byLabel.*.guardians / students.guardianIds / parentLists).
 */
(function () {
    'use strict';

    function normStr(v) {
        return String(v ?? '').trim();
    }

    function normEmail(v) {
        return normStr(v).toLowerCase();
    }

    function normCode(v) {
        return normStr(v).toUpperCase();
    }

    function psQuote(s) {
        return "'" + String(s ?? '').replace(/'/g, "''") + "'";
    }

    function mailNickSanitize(raw, maxLen) {
        if (typeof window.ms365AppDataV2?.mailNicknamePrefixSanitize === 'function') {
            return window.ms365AppDataV2.mailNicknamePrefixSanitize(raw, maxLen || 60);
        }
        const lim = typeof maxLen === 'number' && maxLen > 0 ? maxLen : 60;
        let out = '';
        const s = String(raw ?? '')
            .trim()
            .toLowerCase();
        for (let i = 0; i < s.length; i++) {
            const c = s.charCodeAt(i);
            if (c < 32 || c === 127 || c > 127) continue;
            const ch = s.charAt(i);
            if (/[@()[\]\\";:<>,\s]/.test(ch)) continue;
            out += ch;
        }
        return out.length > lim ? out.slice(0, lim) : out;
    }

    /** Bausteine wie Kursteams: text | klasse | year */
    function defaultClassAliasPattern() {
        return [
            { type: 'text', value: 'eltern' },
            { type: 'klasse' }
        ];
    }

    function defaultClassDisplayPattern() {
        return [
            { type: 'text', value: 'Eltern ' },
            { type: 'klasse' }
        ];
    }

    function defaultYearAliasPattern() {
        return [
            { type: 'text', value: 'elternjg' },
            { type: 'year' }
        ];
    }

    function defaultYearDisplayPattern() {
        return [
            { type: 'text', value: 'Eltern JG ' },
            { type: 'year' }
        ];
    }

    function normalizeNamePattern(pattern, fallback) {
        const arr = Array.isArray(pattern) ? pattern : [];
        const out = [];
        arr.forEach(function (p) {
            if (!p || typeof p !== 'object') return;
            const type = String(p.type || '').trim();
            if (type === 'text') out.push({ type: 'text', value: String(p.value ?? '') });
            else if (type === 'klasse' || type === 'year') out.push({ type: type });
        });
        if (out.length) return out;
        return Array.isArray(fallback) ? fallback.slice() : [];
    }

    function tokenLabel(t) {
        if (!t) return '';
        if (t.type === 'klasse') return 'Klasse';
        if (t.type === 'year') return 'Abschlussjahr';
        if (t.type === 'text') return 'Text';
        return String(t.type || '');
    }

    /**
     * @param {array} pattern
     * @param {{ klasse?: string, year?: string, forAlias?: boolean }} ctx
     */
    function buildNameFromPattern(pattern, ctx) {
        const c = ctx && typeof ctx === 'object' ? ctx : {};
        const forAlias = !!c.forAlias;
        const klasseRaw = normStr(c.klasse);
        const yearRaw = normStr(c.year);
        const parts = [];
        normalizeNamePattern(pattern, []).forEach(function (p) {
            if (p.type === 'text') {
                parts.push(String(p.value ?? ''));
            } else if (p.type === 'klasse') {
                parts.push(forAlias ? String(klasseRaw).toLowerCase() : klasseRaw ? normCode(klasseRaw) : '');
            } else if (p.type === 'year') {
                parts.push(yearRaw);
            }
        });
        const joined = parts.join('');
        if (forAlias) {
            const nick = mailNickSanitize(joined, 60);
            return nick || 'eltern';
        }
        return joined.trim() || 'Eltern';
    }

    function namingFromSetup(setup) {
        const s = setup && typeof setup === 'object' ? setup : {};
        return {
            classAliasPattern: normalizeNamePattern(s.elternClassAliasPattern, defaultClassAliasPattern()),
            classDisplayPattern: normalizeNamePattern(s.elternClassDisplayPattern, defaultClassDisplayPattern()),
            yearAliasPattern: normalizeNamePattern(s.elternYearAliasPattern, defaultYearAliasPattern()),
            yearDisplayPattern: normalizeNamePattern(s.elternYearDisplayPattern, defaultYearDisplayPattern())
        };
    }

    function getNaming() {
        let setup = {};
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
                setup = window.ms365AppDataV2.getSetup() || {};
            }
        } catch {
            setup = {};
        }
        return namingFromSetup(setup);
    }

    function deriveClassParentDisplayName(classCode, year, naming) {
        const n = naming || getNaming();
        return buildNameFromPattern(n.classDisplayPattern, {
            klasse: classCode,
            year: year,
            forAlias: false
        });
    }

    function deriveClassParentNick(classCode, year, naming) {
        const n = naming || getNaming();
        return buildNameFromPattern(n.classAliasPattern, {
            klasse: classCode,
            year: year,
            forAlias: true
        });
    }

    function deriveYearParentDisplayName(year, naming) {
        const n = naming || getNaming();
        return buildNameFromPattern(n.yearDisplayPattern, { year: year, forAlias: false });
    }

    function deriveYearParentNick(year, naming) {
        const n = naming || getNaming();
        return buildNameFromPattern(n.yearAliasPattern, { year: year, forAlias: true });
    }

    function contactAliasFromEmail(email) {
        const em = normEmail(email);
        const local = em.split('@')[0] || 'eltern';
        const base = mailNickSanitize('el-' + local.replace(/[^a-z0-9._-]/gi, ''), 40) || 'el-contact';
        return base;
    }

    /**
     * @param {{ students?: array, classes?: array, guardians?: array, parentLists?: array }} yearBucket
     * @param {{ naming?: object }} [opts]
     * @returns {array} SOLL-Zeilen pro Klasse
     */
    function buildClassParentSoll(yearBucket, opts) {
        const y = yearBucket && typeof yearBucket === 'object' ? yearBucket : {};
        const naming = (opts && opts.naming) || getNaming();
        const guardians = Array.isArray(y.guardians) ? y.guardians : [];
        const byId = new Map();
        guardians.forEach(function (g) {
            if (g && g.id) byId.set(String(g.id), g);
        });
        const students = Array.isArray(y.students) ? y.students : [];
        const classes = Array.isArray(y.classes) ? y.classes : [];
        const classMeta = new Map();
        classes.forEach(function (c) {
            const code = normCode(c && c.code);
            if (!code) return;
            classMeta.set(code, c);
        });

        const byClass = new Map();
        function ensure(code) {
            if (!byClass.has(code)) {
                byClass.set(code, {
                    scope: 'class',
                    code: code,
                    studentCount: 0,
                    guardianEmails: new Map(),
                    abschlussJahr: ''
                });
            }
            return byClass.get(code);
        }

        classMeta.forEach(function (_c, code) {
            const row = ensure(code);
            const meta = classMeta.get(code);
            const yr = normStr(meta && meta.year);
            if (/^\d{4}$/.test(yr)) row.abschlussJahr = yr;
        });

        students.forEach(function (s) {
            const code = normCode(s && s.klasse);
            if (!code) return;
            const row = ensure(code);
            row.studentCount += 1;
            const ids = Array.isArray(s.guardianIds) ? s.guardianIds : [];
            ids.forEach(function (gid) {
                const g = byId.get(String(gid));
                if (!g || !g.email) return;
                const em = normEmail(g.email);
                if (!em || em.indexOf('@') === -1) return;
                if (!row.guardianEmails.has(em)) {
                    row.guardianEmails.set(em, {
                        email: em,
                        name: normStr(g.name) || em,
                        id: String(g.id)
                    });
                }
            });
        });

        const parentLists = Array.isArray(y.parentLists) ? y.parentLists : [];
        const linkByCode = new Map();
        parentLists.forEach(function (p) {
            if (!p || p.scope !== 'class') return;
            const code = normCode(p.code);
            if (code) linkByCode.set(code, p);
        });

        return Array.from(byClass.values())
            .map(function (row) {
                const link = linkByCode.get(row.code) || null;
                const emails = Array.from(row.guardianEmails.values()).sort(function (a, b) {
                    return a.email.localeCompare(b.email, 'de');
                });
                return {
                    scope: 'class',
                    code: row.code,
                    displayName: deriveClassParentDisplayName(row.code, row.abschlussJahr, naming),
                    mailNickname: deriveClassParentNick(row.code, row.abschlussJahr, naming),
                    graphGroupId: link && link.graphGroupId ? String(link.graphGroupId) : '',
                    lastExportAt: link && link.lastExportAt ? String(link.lastExportAt) : '',
                    studentCount: row.studentCount,
                    guardianCount: emails.length,
                    guardians: emails,
                    abschlussJahr: row.abschlussJahr
                };
            })
            .sort(function (a, b) {
                return String(a.code).localeCompare(String(b.code), 'de', { numeric: true });
            });
    }

    /**
     * Jahrgangs-Elternlisten nach Abschlussjahr der Klassen.
     * @param {object} yearBucket
     * @param {{ naming?: object }} [opts]
     */
    function buildYearParentSoll(yearBucket, opts) {
        const naming = (opts && opts.naming) || getNaming();
        const classRows = buildClassParentSoll(yearBucket, { naming: naming });
        const byYear = new Map();
        classRows.forEach(function (cr) {
            const yr = normStr(cr.abschlussJahr);
            if (!/^\d{4}$/.test(yr)) return;
            if (!byYear.has(yr)) {
                byYear.set(yr, {
                    scope: 'year',
                    code: yr,
                    classCodes: [],
                    guardianEmails: new Map()
                });
            }
            const row = byYear.get(yr);
            row.classCodes.push(cr.code);
            (cr.guardians || []).forEach(function (g) {
                if (!row.guardianEmails.has(g.email)) row.guardianEmails.set(g.email, g);
            });
        });

        const y = yearBucket && typeof yearBucket === 'object' ? yearBucket : {};
        const parentLists = Array.isArray(y.parentLists) ? y.parentLists : [];
        const linkByYear = new Map();
        parentLists.forEach(function (p) {
            if (!p || p.scope !== 'year') return;
            const code = normStr(p.code);
            if (/^\d{4}$/.test(code)) linkByYear.set(code, p);
        });

        return Array.from(byYear.values())
            .map(function (row) {
                const link = linkByYear.get(row.code) || null;
                const emails = Array.from(row.guardianEmails.values()).sort(function (a, b) {
                    return a.email.localeCompare(b.email, 'de');
                });
                return {
                    scope: 'year',
                    code: row.code,
                    displayName: deriveYearParentDisplayName(row.code, naming),
                    mailNickname: deriveYearParentNick(row.code, naming),
                    graphGroupId: link && link.graphGroupId ? String(link.graphGroupId) : '',
                    lastExportAt: link && link.lastExportAt ? String(link.lastExportAt) : '',
                    classCodes: row.classCodes.slice().sort(function (a, b) {
                        return String(a).localeCompare(String(b), 'de', { numeric: true });
                    }),
                    guardianCount: emails.length,
                    guardians: emails
                };
            })
            .sort(function (a, b) {
                return String(b.code).localeCompare(String(a.code));
            });
    }

    function psHeader(lines) {
        lines.push('#Requires -Modules ExchangeOnlineManagement');
        lines.push('# Eltern-Verteiler: Mail Contacts (GAL-versteckt) + Distribution Groups');
        lines.push('# - DL in GAL sichtbar, Mitgliedschaft versteckt (HiddenGroupMembershipEnabled)');
        lines.push('# - Einzelne Elternmails nicht in der GAL (HiddenFromAddressListsEnabled am Contact)');
        lines.push('# Vor dem Ausführen: Connect-ExchangeOnline');
        lines.push('$ErrorActionPreference = "Continue"');
        lines.push('');
    }

    function psFooter(lines) {
        lines.push('');
        lines.push('Write-Host "Fertig." -ForegroundColor Cyan');
    }

    /**
     * Vollständiges Sync-Skript für eine oder mehrere Eltern-Listen.
     * @param {{ lists: array, domain?: string, managedBy?: string[] }} opts
     * lists[].guardians = [{email, name}]
     */
    function buildElternSyncScript(opts) {
        const o = opts && typeof opts === 'object' ? opts : {};
        const lists = Array.isArray(o.lists) ? o.lists : [];
        const domain = normStr(o.domain).replace(/^@+/, '');
        const managedBy = Array.isArray(o.managedBy) ? o.managedBy.map(normEmail).filter(Boolean) : [];
        const lines = [];
        psHeader(lines);

        if (!lists.length) {
            lines.push('Write-Host "Keine Listen ausgewählt." -ForegroundColor Yellow');
            psFooter(lines);
            return lines.join('\n');
        }

        const allContacts = new Map();
        lists.forEach(function (list) {
            (list.guardians || []).forEach(function (g) {
                const em = normEmail(g && g.email);
                if (!em || em.indexOf('@') === -1) return;
                if (!allContacts.has(em)) {
                    allContacts.set(em, {
                        email: em,
                        name: normStr(g.name) || em,
                        alias: contactAliasFromEmail(em)
                    });
                }
            });
        });

        lines.push('# === 1) Mail Contacts anlegen/aktualisieren (nicht in GAL) ===');
        lines.push('$Contacts = @(');
        Array.from(allContacts.values()).forEach(function (c, idx, arr) {
            const comma = idx < arr.length - 1 ? ',' : '';
            lines.push(
                '  [pscustomobject]@{ Email = ' +
                    psQuote(c.email) +
                    '; Name = ' +
                    psQuote(c.name) +
                    '; Alias = ' +
                    psQuote(c.alias) +
                    ' }' +
                    comma
            );
        });
        lines.push(')');
        lines.push('foreach ($c in $Contacts) {');
        lines.push('  $existing = Get-MailContact -Identity $c.Email -ErrorAction SilentlyContinue');
        lines.push('  if (-not $existing) {');
        lines.push('    try {');
        lines.push(
            '      New-MailContact -Name $c.Name -ExternalEmailAddress $c.Email -Alias $c.Alias -ErrorAction Stop | Out-Null'
        );
        lines.push('      Write-Host ("Contact angelegt: " + $c.Email) -ForegroundColor Green');
        lines.push('    } catch {');
        lines.push('      Write-Host ("FEHLER Contact " + $c.Email + ": " + $_.Exception.Message) -ForegroundColor Red');
        lines.push('      continue');
        lines.push('    }');
        lines.push('  }');
        lines.push('  try {');
        lines.push('    Set-MailContact -Identity $c.Email -HiddenFromAddressListsEnabled $true -ErrorAction Stop');
        lines.push('  } catch {');
        lines.push('    Write-Host ("Warnung Hide GAL " + $c.Email + ": " + $_.Exception.Message) -ForegroundColor Yellow');
        lines.push('  }');
        lines.push('}');
        lines.push('');

        lists.forEach(function (list, listIdx) {
            const displayName = normStr(list.displayName) || 'Eltern';
            const alias = mailNickSanitize(list.mailNickname || '', 60) || ('eltern' + String(listIdx + 1));
            let smtp = normStr(list.primarySmtp);
            if (!smtp && domain) smtp = alias + '@' + domain;
            const members = (list.guardians || [])
                .map(function (g) {
                    return normEmail(g && g.email);
                })
                .filter(function (em) {
                    return em && em.indexOf('@') !== -1;
                });
            const uniq = Array.from(new Set(members));

            lines.push('# === Liste: ' + displayName + ' (' + alias + ') ===');
            lines.push('$DisplayName = ' + psQuote(displayName));
            lines.push('$Alias = ' + psQuote(alias));
            if (smtp) lines.push('$Smtp = ' + psQuote(smtp));
            lines.push('$Wanted = @(' + uniq.map(psQuote).join(', ') + ')');
            lines.push('');
            lines.push('$dg = Get-DistributionGroup -Identity $Alias -ErrorAction SilentlyContinue');
            lines.push('if (-not $dg) {');
            lines.push('  $dg = Get-DistributionGroup -Identity $DisplayName -ErrorAction SilentlyContinue');
            lines.push('}');
            lines.push('if (-not $dg) {');
            const newArgs = ['-Name $DisplayName', '-DisplayName $DisplayName', '-Alias $Alias', '-Type Distribution', '-HiddenGroupMembershipEnabled'];
            if (smtp) newArgs.push('-PrimarySmtpAddress $Smtp');
            lines.push('  New-DistributionGroup ' + newArgs.join(' ') + ' | Out-Null');
            lines.push('  Write-Host ("DL angelegt: " + $DisplayName) -ForegroundColor Green');
            lines.push('  $dg = Get-DistributionGroup -Identity $Alias -ErrorAction Stop');
            lines.push('} else {');
            lines.push('  Write-Host ("DL vorhanden: " + $dg.DisplayName) -ForegroundColor Cyan');
            lines.push('}');
            lines.push('');
            lines.push('# GAL sichtbar, Mitgliedschaft versteckt, keine externen Absender');
            lines.push(
                'Set-DistributionGroup -Identity $dg.Identity -HiddenFromAddressListsEnabled $false -HiddenGroupMembershipEnabled:$true -RequireSenderAuthenticationEnabled $true -ErrorAction SilentlyContinue'
            );
            if (managedBy.length) {
                lines.push('$Owners = @(' + managedBy.map(psQuote).join(', ') + ')');
                lines.push('Set-DistributionGroup -Identity $dg.Identity -ManagedBy $Owners -ErrorAction SilentlyContinue');
            }
            lines.push('');
            lines.push('$current = @()');
            lines.push(
                'try { $current = @(Get-DistributionGroupMember -Identity $dg.Identity -ResultSize Unlimited | ForEach-Object { ($_.PrimarySmtpAddress | Out-String).Trim().ToLower() }) } catch {}'
            );
            lines.push('foreach ($m in $Wanted) {');
            lines.push('  if ($current -contains $m) { continue }');
            lines.push('  try {');
            lines.push('    Add-DistributionGroupMember -Identity $dg.Identity -Member $m -ErrorAction Stop');
            lines.push('    Write-Host ("  + " + $m) -ForegroundColor Green');
            lines.push('  } catch {');
            lines.push('    Write-Host ("  FEHLER + " + $m + ": " + $_.Exception.Message) -ForegroundColor Red');
            lines.push('  }');
            lines.push('}');
            lines.push('foreach ($m in $current) {');
            lines.push('  if (-not $m) { continue }');
            lines.push('  if ($Wanted -contains $m) { continue }');
            lines.push('  try {');
            lines.push('    Remove-DistributionGroupMember -Identity $dg.Identity -Member $m -Confirm:$false -ErrorAction Stop');
            lines.push('    Write-Host ("  - " + $m) -ForegroundColor Yellow');
            lines.push('  } catch {');
            lines.push('    Write-Host ("  FEHLER - " + $m + ": " + $_.Exception.Message) -ForegroundColor Red');
            lines.push('  }');
            lines.push('}');
            lines.push(
                'Get-DistributionGroup -Identity $dg.Identity | Format-List DisplayName,PrimarySmtpAddress,Alias,HiddenFromAddressListsEnabled,HiddenGroupMembershipEnabled'
            );
            lines.push('');
        });

        psFooter(lines);
        return lines.join('\n');
    }

    /**
     * Lokale Diagnose: Naming, leere Listen, fehlende Mails, Alias-Kollisionen, letzter Export.
     * @param {object} bucket years.byLabel[current]
     * @param {object} [naming]
     * @param {string} [domain]
     */
    function buildElternDiagnoseReport(bucket, naming, domain) {
        const b = bucket && typeof bucket === 'object' ? bucket : {};
        const n = naming || getNaming();
        const dom = normStr(domain).toLowerCase();
        const issues = [];
        const lists = [];
        const classSoll = buildClassParentSoll(b, { naming: n });
        const yearSoll = buildYearParentSoll(b, { naming: n });
        const allSoll = classSoll.concat(yearSoll);
        const nickSeen = new Map();
        const smtpSeen = new Map();
        let withParents = 0;
        let emptyLists = 0;

        allSoll.forEach(function (list) {
            const nick = mailNickSanitize((list && list.mailNickname) || '', 60);
            const smtp = nick && dom ? nick + '@' + dom : '';
            const members = Array.isArray(list.guardians) ? list.guardians : [];
            const valid = members.filter(function (g) {
                const em = normEmail(g && g.email);
                return em && em.indexOf('@') !== -1;
            });
            if (valid.length) withParents += 1;
            else emptyLists += 1;
            if (nick) {
                if (nickSeen.has(nick)) {
                    issues.push({
                        level: 'warn',
                        code: 'alias-collision',
                        summary: 'Alias „' + nick + '“ mehrfach: ' + nickSeen.get(nick) + ' und ' + (list.displayName || nick)
                    });
                } else nickSeen.set(nick, list.displayName || nick);
            }
            if (smtp) {
                if (smtpSeen.has(smtp)) {
                    issues.push({
                        level: 'warn',
                        code: 'smtp-collision',
                        summary: 'SMTP „' + smtp + '“ mehrfach vergeben.'
                    });
                } else smtpSeen.set(smtp, true);
            }
            lists.push({
                displayName: list.displayName || '',
                mailNickname: nick,
                primarySmtp: smtp,
                memberCount: valid.length,
                lastExportAt: list.lastExportAt || '',
                galVisible: true,
                contactsHidden: true
            });
        });

        const exported = (Array.isArray(b.parentLists) ? b.parentLists : []).filter(function (p) {
            return p && p.lastExportAt;
        }).length;
        if (!allSoll.length) {
            issues.push({
                level: 'info',
                code: 'no-lists',
                summary: 'Noch keine Klassen oder Abschlussjahre in den Stammdaten – zuerst Schüler/Klassen importieren.'
            });
        } else if (!withParents) {
            issues.push({
                level: 'warn',
                code: 'no-parents',
                summary: 'Listen vorhanden, aber keine Elternmails. Import aus Sokrates/WebUntis oder MS365-Vorlage nachziehen.'
            });
        }
        if (emptyLists && withParents) {
            issues.push({
                level: 'info',
                code: 'empty-lists',
                summary: String(emptyLists) + ' Liste(n) ohne Elternmails (werden im Skript übersprungen).'
            });
        }

        return {
            ok: issues.filter(function (i) {
                return i.level === 'warn';
            }).length === 0,
            counts: {
                lists: allSoll.length,
                withParents: withParents,
                emptyLists: emptyLists,
                exported: exported
            },
            hints: {
                gal:
                    'Verteilerlisten sollen in der GAL sichtbar sein (HiddenFromAddressListsEnabled $false). Das Sync-Skript setzt das.',
                contacts:
                    'Einzelne Elternmails werden als Mail Contacts angelegt und aus der GAL ausgeblendet (HiddenFromAddressListsEnabled $true).',
                naming:
                    'Alias nur ASCII, ohne Leerzeichen. Kollisionen oben prüfen, bevor Sie das Exchange-Skript ausführen.'
            },
            issues: issues,
            lists: lists
        };
    }

    function buildElternDiagnoseScript(lists, domain) {
        const rows = Array.isArray(lists) ? lists : [];
        const lines = [];
        psHeader(lines);
        lines.push('# Diagnose: GAL-Sichtbarkeit der Verteiler + Hide-from-GAL der Contacts');
        lines.push('# Keine Änderungen – nur Get-*');
        lines.push('');
        rows.forEach(function (list) {
            const alias = mailNickSanitize((list && list.mailNickname) || '', 60);
            const display = normStr(list && list.displayName);
            if (!alias && !display) return;
            const ident = alias || display;
            lines.push('Write-Host "=== ' + ident.replace(/"/g, '') + ' ===" -ForegroundColor Cyan');
            lines.push(
                '$dg = Get-DistributionGroup -Identity ' +
                    psQuote(ident) +
                    ' -ErrorAction SilentlyContinue'
            );
            lines.push('if ($dg) {');
            lines.push(
                '  $dg | Format-List DisplayName,PrimarySmtpAddress,Alias,HiddenFromAddressListsEnabled,HiddenGroupMembershipEnabled'
            );
            lines.push('} else {');
            lines.push('  Write-Host "  Verteiler nicht gefunden (noch nicht angelegt?)." -ForegroundColor Yellow');
            lines.push('}');
            lines.push('');
        });
        lines.push('# Stichprobe Contacts: erste 20 MailContacts mit HiddenFromAddressListsEnabled');
        lines.push(
            'Get-MailContact -ResultSize 20 | Select-Object DisplayName,PrimarySmtpAddress,HiddenFromAddressListsEnabled | Format-Table -AutoSize'
        );
        if (domain) {
            lines.push('# Domain-Hinweis: ' + String(domain));
        }
        psFooter(lines);
        return lines.join('\n');
    }

    /**
     * Optionale Elternspalten aus Schüler-Paste-Zeile extrahieren.
     * Format: Klasse;Name;E-Mail[;Eltern1;Eltern1Mail[;Eltern2;Eltern2Mail...]]
     */
    function extractParentPairsFromParts(parts) {
        const out = [];
        if (!Array.isArray(parts) || parts.length <= 3) return out;
        for (let i = 3; i < parts.length; i += 2) {
            const name = normStr(parts[i] || '');
            const email = normEmail(parts[i + 1] || '');
            if (!name && !email) continue;
            if (email && email.indexOf('@') === -1) {
                // Toleranz: nur Mail ohne Namen
                if (name.indexOf('@') !== -1) {
                    out.push({ name: '', email: normEmail(name) });
                }
                continue;
            }
            if (!email) continue;
            out.push({ name: name, email: email });
        }
        return out;
    }

    window.ms365ElternGuardians = {
        normEmail,
        defaultClassAliasPattern,
        defaultClassDisplayPattern,
        defaultYearAliasPattern,
        defaultYearDisplayPattern,
        normalizeNamePattern,
        buildNameFromPattern,
        tokenLabel,
        namingFromSetup,
        getNaming,
        deriveClassParentDisplayName,
        deriveClassParentNick,
        deriveYearParentDisplayName,
        deriveYearParentNick,
        contactAliasFromEmail,
        buildClassParentSoll,
        buildYearParentSoll,
        buildElternSyncScript,
        buildElternDiagnoseReport,
        buildElternDiagnoseScript,
        extractParentPairsFromParts,
        mailNickSanitize
    };
})();
