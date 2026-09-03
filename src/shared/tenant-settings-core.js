(function () {
    'use strict';

    const STORAGE_KEY = 'ms365-tenant-settings-v1';
    const CURRENT_VERSION = 2;

    function normStr(v) {
        return String(v ?? '').trim();
    }

    function normCode(v) {
        return normStr(v).toUpperCase();
    }

    const DEFAULT_ADMIN_ROLE_NAMES = [
        'Direktion',
        'Sekretariat',
        'Administration',
        'Schularzt',
        'Schulwart',
        'IT-Support',
        'Bibliothek'
    ];

    function adminRoleCodeFromName(name) {
        const c = normStr(name)
            .toUpperCase()
            .replace(/\s+/g, '')
            .replace(/[^A-Z0-9ÄÖÜß-]/g, '')
            .slice(0, 24);
        return c;
    }

    function uniqueAdminRoleCode(desired, used) {
        let code = adminRoleCodeFromName(desired) || 'ROLLE';
        const usedSet = used instanceof Set ? used : new Set();
        if (!usedSet.has(code.toLowerCase())) return code;
        let i = 2;
        while (usedSet.has((code + String(i)).toLowerCase())) i += 1;
        return (code + String(i)).slice(0, 24);
    }

    function defaultAdminRoleCatalog() {
        const used = new Set();
        return DEFAULT_ADMIN_ROLE_NAMES.map(function (name) {
            const code = uniqueAdminRoleCode(name, used);
            used.add(code.toLowerCase());
            return { code: code, name: name };
        });
    }

    function personMatchesAdminRole(row, role) {
        if (!row || !role) return false;
        const r = normStr(row.role).toLowerCase();
        const dk = normStr(row.defaultKey).toLowerCase();
        const n = normStr(role.name).toLowerCase();
        const c = normStr(role.code).toLowerCase();
        if (n && (r === n || dk === n)) return true;
        if (c && (r === c || dk === c)) return true;
        return false;
    }

    function normalizeAdminRoleCatalog(rolesIn, adminIn) {
        const seen = new Set();
        const roles = [];
        (Array.isArray(rolesIn) ? rolesIn : []).forEach(function (raw) {
            const name = normStr(raw && (raw.name || raw.role || raw.bezeichnung));
            let code = normCode(raw && raw.code);
            if (!code && name) code = adminRoleCodeFromName(name);
            if (!code && !name) return;
            if (!code) code = uniqueAdminRoleCode(name, seen);
            const key = code.toLowerCase();
            if (seen.has(key)) return;
            seen.add(key);
            roles.push({ code: code, name: name || code });
        });
        (Array.isArray(adminIn) ? adminIn : []).forEach(function (a) {
            const name = normStr(a && (a.role || a.rolle || a.title));
            if (!name) return;
            const exists = roles.some(function (r) {
                return r.name.toLowerCase() === name.toLowerCase() || r.code.toLowerCase() === name.toLowerCase();
            });
            if (exists) return;
            const code = uniqueAdminRoleCode(name, seen);
            seen.add(code.toLowerCase());
            roles.push({ code: code, name: name });
        });
        return roles;
    }

    function renameAdminRole(roles, admin, fromName, toName) {
        const groups = adminRolesAndAdminToGroups(roles, admin);
        const nextGroups = renameAdminRoleInGroups(groups, fromName, toName);
        return {
            roles: deriveAdminRolesFromGroups(nextGroups),
            admin: deriveAdminPeopleFromGroups(nextGroups)
        };
    }

    function isAdministrationGrouped(entries) {
        return (
            Array.isArray(entries) &&
            entries.some(function (entry) {
                return entry && Array.isArray(entry.people);
            })
        );
    }

    function normalizeAdminPersonEntry(row) {
        const person = {
            name: normStr(row && row.name),
            email: normStr(row && row.email).toLowerCase()
        };
        const defaultKey = normStr(row && row.defaultKey);
        if (defaultKey) person.defaultKey = defaultKey;
        return person;
    }

    function deriveAdminPeopleFromGroups(groups) {
        const out = [];
        (Array.isArray(groups) ? groups : []).forEach(function (group) {
            const roleName = normStr(group && group.name);
            (Array.isArray(group && group.people) ? group.people : []).forEach(function (personRow) {
                const person = normalizeAdminPersonEntry(personRow);
                if (!roleName && !person.name && !person.email) return;
                const row = { role: roleName, name: person.name, email: person.email };
                if (person.defaultKey) row.defaultKey = person.defaultKey;
                out.push(row);
            });
        });
        return out;
    }

    function deriveAdminRolesFromGroups(groups) {
        const out = [];
        const seen = new Set();
        (Array.isArray(groups) ? groups : []).forEach(function (group) {
            const name = normStr(group && group.name);
            let code = normCode(group && group.code);
            if (!code && name) code = adminRoleCodeFromName(name);
            if (!code && !name) return;
            const key = (name || code).toLowerCase();
            if (seen.has(key)) return;
            seen.add(key);
            out.push({ code: code || adminRoleCodeFromName(name), name: name || code });
        });
        return out;
    }

    function adminRolesAndAdminToGroups(rolesIn, adminIn) {
        const roles = normalizeAdminRoleCatalog(rolesIn, adminIn);
        const groupMap = new Map();
        roles.forEach(function (role) {
            const name = normStr(role && role.name);
            const code = normCode(role && role.code) || adminRoleCodeFromName(name);
            if (!name && !code) return;
            groupMap.set(name.toLowerCase(), { code: code, name: name || code, people: [] });
        });
        (Array.isArray(adminIn) ? adminIn : []).forEach(function (row) {
            const roleName = normStr(row && (row.role || row.rolle || row.title));
            if (!roleName) return;
            const key = roleName.toLowerCase();
            if (!groupMap.has(key)) {
                groupMap.set(key, {
                    code: adminRoleCodeFromName(roleName),
                    name: roleName,
                    people: []
                });
            }
            const person = normalizeAdminPersonEntry(row);
            if (!person.name && !person.email) return;
            groupMap.get(key).people.push(person);
        });
        return Array.from(groupMap.values());
    }

    function flatKindEntriesToGroups(entries) {
        const administration = Array.isArray(entries) ? entries : [];
        const roles = normalizeAdminRoleCatalog(
            administration
                .filter(function (entry) {
                    return entry && entry.kind === 'role';
                })
                .map(function (entry) {
                    return { code: normCode(entry.code), name: normStr(entry.name) };
                }),
            administration
                .filter(function (entry) {
                    return entry && entry.kind === 'person';
                })
                .map(function (entry) {
                    return {
                        role: normStr(entry.role),
                        name: normStr(entry.name),
                        email: normStr(entry.email).toLowerCase(),
                        defaultKey: normStr(entry.defaultKey)
                    };
                })
        );
        const admin = administration
            .filter(function (entry) {
                return entry && entry.kind === 'person';
            })
            .map(function (entry) {
                const row = {
                    role: normStr(entry.role),
                    name: normStr(entry.name),
                    email: normStr(entry.email).toLowerCase()
                };
                const defaultKey = normStr(entry.defaultKey);
                if (defaultKey) row.defaultKey = defaultKey;
                return row;
            });
        return adminRolesAndAdminToGroups(roles, admin);
    }

    function normalizeGroupedAdministrationEntries(groupsIn) {
        const out = [];
        const seen = new Set();
        (Array.isArray(groupsIn) ? groupsIn : []).forEach(function (group) {
            const name = normStr(group && group.name);
            const code = normCode(group && group.code) || adminRoleCodeFromName(name);
            if (!name && !code) return;
            const key = (name || code).toLowerCase();
            let target = out.find(function (entry) {
                return (entry.name || entry.code).toLowerCase() === key;
            });
            if (!target) {
                target = { code: code, name: name || code, people: [] };
                out.push(target);
                seen.add(key);
            }
            const peopleSeen = new Set(
                target.people.map(function (person) {
                    return [person.name, person.email].join('\u0001').toLowerCase();
                })
            );
            (Array.isArray(group && group.people) ? group.people : []).forEach(function (personRow) {
                const person = normalizeAdminPersonEntry(personRow);
                if (!person.name && !person.email) return;
                const personKey = [person.name, person.email].join('\u0001').toLowerCase();
                if (peopleSeen.has(personKey)) return;
                peopleSeen.add(personKey);
                target.people.push(person);
            });
        });
        return out;
    }

    function renameAdminRoleInGroups(groups, fromName, toName) {
        const fromL = normStr(fromName).toLowerCase();
        const to = normStr(toName);
        return (Array.isArray(groups) ? groups : []).map(function (group) {
            const next = Object.assign({}, group, {
                people: (Array.isArray(group.people) ? group.people : []).map(function (person) {
                    return Object.assign({}, person);
                })
            });
            if (normStr(next.name).toLowerCase() === fromL || normStr(next.code).toLowerCase() === fromL) {
                next.name = to || next.name;
                if (to && (!next.code || normStr(next.code).toLowerCase() === fromL)) {
                    next.code = adminRoleCodeFromName(to) || next.code;
                }
            }
            return next;
        });
    }

    function normalizeAdministrationEntries(entriesIn, rolesIn, adminIn) {
        if (isAdministrationGrouped(entriesIn)) {
            return normalizeGroupedAdministrationEntries(entriesIn);
        }
        if (
            Array.isArray(entriesIn) &&
            entriesIn.some(function (entry) {
                return entry && (entry.kind === 'role' || entry.kind === 'person');
            })
        ) {
            return flatKindEntriesToGroups(entriesIn);
        }
        return adminRolesAndAdminToGroups(rolesIn, adminIn);
    }

    function deriveAdminRolesFromAdministration(entries) {
        return deriveAdminRolesFromGroups(normalizeAdministrationEntries(entries, [], []));
    }

    function deriveAdminPeopleFromAdministration(entries) {
        return deriveAdminPeopleFromGroups(normalizeAdministrationEntries(entries, [], []));
    }

    function parseLinesToAdminRoles(text) {
        const out = [];
        parseDelimitedLines(text).forEach((parts) => {
            if (parts.length >= 2) {
                const code = normCode(parts[0] || '');
                const name = normStr(parts.slice(1).join(' '));
                if (!code && !name) return;
                out.push({ code: code || adminRoleCodeFromName(name), name: name || code });
                return;
            }
            const name = normStr(parts[0] || '');
            if (!name) return;
            out.push({ code: adminRoleCodeFromName(name), name: name });
        });
        return out;
    }

    /** Stabiler Mail-Nickname für Klassen-M365-Gruppe: jg{YYYY}{codeAlphaNum} (Kursteam/Umbenennen). */
    /** localStorage-Key für das konfigurierbare Alias-Schema der Klassengruppen */
    const CLASS_NICK_SCHEMA_KEY = 'ms365-class-nick-schema-v1';

    /** Gibt das aktuell gespeicherte Alias-Schema zurück (oder den Default). */
    function getClassNickSchema() {
        try {
            const raw = localStorage.getItem(CLASS_NICK_SCHEMA_KEY);
            if (raw) {
                const obj = JSON.parse(raw);
                return {
                    prefix: normStr(obj.prefix || 'jg'),
                    pattern: normStr(obj.pattern || '{prefix}{year}-{suffix}'),
                    upper: !!obj.upper
                };
            }
        } catch { /* ignore */ }
        return { prefix: 'jg', pattern: '{prefix}{year}-{suffix}', upper: false };
    }

    /** Speichert das Alias-Schema. */
    function saveClassNickSchema(schema) {
        try {
            localStorage.setItem(CLASS_NICK_SCHEMA_KEY, JSON.stringify({
                prefix: normStr(schema.prefix || 'jg') || 'jg',
                pattern: normStr(schema.pattern || '{prefix}{year}-{suffix}') || '{prefix}{year}-{suffix}',
                upper: !!schema.upper
            }));
        } catch { /* ignore */ }
    }

    /**
     * Baut den Mail-Nickname für eine Klasse nach dem konfigurierten Schema.
     * Platzhalter: {prefix}, {year}, {klasse}, {suffix}, {kv}
     * @param {string} yearRaw   Abschlussjahr (4-stellig)
     * @param {string} codeRaw   Klassencode (z. B. "1AK")
     * @param {object} [rowExtra] Optional: { headName, headEmail } für {kv}
     */
    function deriveClassStableMailNickname(yearRaw, codeRaw, rowExtra) {
        const schema = getClassNickSchema();
        const y = normStr(yearRaw);
        const yy = /^\d{4}$/.test(y) ? y : '';
        const code = normCode(codeRaw);
        if (!code) return '';

        // {suffix} = Buchstaben-Teil (z.B. AK aus 1AK)
        const suffixMatch = code.match(/^[0-9]*([A-Za-z]+)$/);
        const suffixRaw = suffixMatch ? suffixMatch[1] : code;
        const suffix = schema.upper ? suffixRaw.toUpperCase() : suffixRaw.toLowerCase();
        const klasseToken = schema.upper ? code.replace(/[^A-Za-z0-9]/g,'').toUpperCase() : code.replace(/[^A-Za-z0-9]/g,'').toLowerCase();
        const prefixToken = (schema.prefix || 'jg').toLowerCase().replace(/[^a-z0-9]/g,'');

        // {kv} = Kürzel des Klassenvorstands
        let kvToken = '';
        if (rowExtra) {
            const name = normStr(rowExtra.headName || '');
            const bracketMatch = name.match(/\(([^)]+)\)\s*$/);
            if (bracketMatch) {
                kvToken = bracketMatch[1].trim();
            } else {
                const lastWord = name.split(/\s+/).pop();
                if (lastWord && lastWord.length <= 6) kvToken = lastWord;
            }
            if (!kvToken) {
                const mail = normStr(rowExtra.headEmail || '');
                if (mail.indexOf('@') !== -1) kvToken = mail.split('@')[0].replace(/[^A-Za-z]/g,'').slice(0,8);
            }
        }
        kvToken = schema.upper ? kvToken.toUpperCase() : kvToken.toLowerCase();

        const raw = (schema.pattern || '{prefix}{year}-{suffix}')
            .replaceAll('{prefix}', prefixToken)
            .replaceAll('{year}', yy)
            .replaceAll('{klasse}', klasseToken)
            .replaceAll('{suffix}', suffix)
            .replaceAll('{kv}', kvToken);

        const sanitized = raw
            .trim()
            .replace(/\s+/g, '-')
            .replace(/[^a-zA-Z0-9-]/g, '')
            .replace(/-+/g, '-')
            .replace(/^-|-$/g, '')
            .toLowerCase()
            .slice(0, 60);

        if (sanitized) return sanitized;
        // Fallback: klassischer Stil
        const tail = code.replace(/[^0-9A-Za-z]/g,'').toLowerCase().slice(0,24);
        return (prefixToken + (yy || '') + tail).slice(0,60);
    }

    function safeJsonParse(s) {
        try {
            return JSON.parse(String(s));
        } catch {
            return null;
        }
    }

    function loadRaw() {
        try {
            const raw = localStorage.getItem(STORAGE_KEY);
            if (!raw) return null;
            return safeJsonParse(raw);
        } catch {
            return null;
        }
    }

    function normalizeSettings(obj) {
        const o = obj && typeof obj === 'object' ? obj : {};
        const schoolName = normStr(o.schoolName || o.name || o.school);
        let domain = normStr(o.domain).replace(/^@+/, '');
        if (!domain && typeof window.ms365GetSchoolDomainNoAt === 'function') {
            domain = normStr(window.ms365GetSchoolDomainNoAt()).replace(/^@+/, '');
        }

        const subjectsIn = Array.isArray(o.subjects) ? o.subjects : [];
        const argesIn = Array.isArray(o.arges) ? o.arges : [];
        const teachersIn = Array.isArray(o.teachers) ? o.teachers : [];
        const administrationIn = Array.isArray(o.administration) ? o.administration : [];
        const adminIn = Array.isArray(o.admin) ? o.admin : [];
        const adminRolesIn = Array.isArray(o.adminRoles) ? o.adminRoles : (Array.isArray(o.verwaltungRollen) ? o.verwaltungRollen : []);
        const studentsIn = Array.isArray(o.students) ? o.students : [];
        const studentCouncilIn = Array.isArray(o.studentCouncil) ? o.studentCouncil : [];
        const classesIn = Array.isArray(o.classes) ? o.classes : [];
        const sgaIn = Array.isArray(o.sga) ? o.sga : [];
        const sgaModeIn = normStr(o.sgaMode).toLowerCase();

        const subjectsSeen = new Set();
        const subjects = [];
        subjectsIn.forEach((s) => {
            const code = normCode(s?.code);
            const name = normStr(s?.name);
            if (!code) return;
            const key = code.toLowerCase();
            if (subjectsSeen.has(key)) return;
            subjectsSeen.add(key);
            subjects.push({ code, name });
        });

        const argesSeen = new Set();
        const arges = [];
        argesIn.forEach((a) => {
            const code = normCode(a?.code);
            const name = normStr(a?.name);
            const subjectsRaw = Array.isArray(a?.subjects) ? a.subjects : Array.isArray(a?.faecher) ? a.faecher : [];
            const subjects = (subjectsRaw || [])
                .map((x) => normCode(x))
                .filter(Boolean);
            if (!code) return;
            const key = code.toLowerCase();
            if (argesSeen.has(key)) return;
            argesSeen.add(key);
            arges.push({ code, name, subjects });
        });

        const teachersSeen = new Set();
        const teachers = [];
        teachersIn.forEach((t) => {
            const code = normCode(t?.code);
            const name = normStr(t?.name);
            const email = normStr(t?.email).toLowerCase();
            if (!code) return;
            const key = code.toLowerCase();
            if (teachersSeen.has(key)) return;
            teachersSeen.add(key);
            teachers.push({ code, name, email });
        });

        const administration = normalizeAdministrationEntries(administrationIn, adminRolesIn, adminIn);
        const admin = deriveAdminPeopleFromGroups(administration);
        const adminRoles = deriveAdminRolesFromGroups(administration);

        const students = [];
        studentsIn.forEach((s) => {
            const klasse = normStr(s?.klasse || s?.class || s?.group || s?.Klassse || s?.Klasse);
            const name = normStr(s?.name);
            const email = normStr(s?.email).toLowerCase();
            if (!klasse && !name && !email) return;
            const row = { klasse, name, email };
            if (s?.id) row.id = normStr(s.id);
            if (Array.isArray(s?.guardianIds)) row.guardianIds = s.guardianIds.slice();
            if (Array.isArray(s?.parentPairs) && s.parentPairs.length) {
                row.parentPairs = s.parentPairs
                    .map((p) => ({
                        name: normStr(p?.name),
                        email: normStr(p?.email).toLowerCase()
                    }))
                    .filter((p) => p.email && p.email.includes('@'));
            }
            students.push(row);
        });

        const studentCouncil = [];
        studentCouncilIn.forEach((s) => {
            const klasse = normStr(s?.klasse || s?.class || s?.group || s?.Klassse || s?.Klasse);
            const name = normStr(s?.name);
            const email = normStr(s?.email).toLowerCase();
            if (!klasse && !name && !email) return;
            studentCouncil.push({ klasse, name, email });
        });

        const classesSeen = new Set();
        const classes = [];
        classesIn.forEach((c) => {
            const code = normCode(c?.code);
            const name = normStr(c?.name || c?.klasse || c?.Klasse);
            const yearRaw = normStr(c?.year || c?.abschlussjahr || c?.Abschlussjahr || c?.graduationYear || '');
            const year = /^\d{4}$/.test(yearRaw) ? yearRaw : '';
            const headName = normStr(c?.headName || c?.klassenvorstandName || c?.kvName);
            const headEmail = normStr(c?.headEmail || c?.klassenvorstandEmail || c?.kvEmail).toLowerCase();
            let stableMailNickname = normStr(c?.stableMailNickname || '')
                .replace(/[^a-zA-Z0-9-]/g, '')
                .replace(/-+/g, '-')
                .replace(/^-|-$/g, '')
                .toLowerCase()
                .slice(0, 60);
            if (!stableMailNickname && year && code) {
                stableMailNickname = deriveClassStableMailNickname(year, code);
            }
            if (!code && !name && !year && !headName && !headEmail) return;
            const key = (code || name).toLowerCase();
            if (classesSeen.has(key)) return;
            classesSeen.add(key);
            classes.push({ code, name, year, headName, headEmail, stableMailNickname });
        });

        const sga = [];
        const sgaSeen = new Set();
        sgaIn.forEach((row) => {
            const scopeRaw = normStr(row?.scope || row?.type || row?.gruppe || row?.rolle).toLowerCase();
            const scope =
                scopeRaw === 'lehrer' || scopeRaw === 'teacher'
                    ? 'teacher'
                    : scopeRaw === 'schueler' || scopeRaw === 'schüler' || scopeRaw === 'student'
                      ? 'student'
                      : scopeRaw === 'extern' || scopeRaw === 'external'
                        ? 'external'
                        : '';
            const name = normStr(row?.name);
            const email = normStr(row?.email).toLowerCase();
            if (!scope && !name && !email) return;
            const key = [scope, name, email].join('\u0001').toLowerCase();
            if (sgaSeen.has(key)) return;
            sgaSeen.add(key);
            sga.push({ scope, name, email });
        });

        return {
            version: CURRENT_VERSION,
            schoolName,
            domain: normStr(domain),
            subjects,
            arges,
            teachers,
            administration,
            admin,
            adminRoles,
            sgaMode: sgaModeIn === 'distribution' ? 'distribution' : 'group',
            sga,
            students,
            studentCouncil,
            classes
        };
    }

    function save(settings) {
        const normalized = normalizeSettings(settings);
        try {
            localStorage.setItem(STORAGE_KEY, JSON.stringify(normalized));
        } catch {
            // ignore
        }
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.setCoreFromTenantSettings === 'function') {
                window.ms365AppDataV2.setCoreFromTenantSettings(normalized);
            }
        } catch {
            // ignore
        }
        if (typeof window.ms365SetSchoolDomainNoAt === 'function' && normalized.domain) {
            window.ms365SetSchoolDomainNoAt(normalized.domain);
        }
        return normalized;
    }

    function load() {
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function') {
                const c = window.ms365AppDataV2.getContainer();
                if (c && c.core && c.years) {
                    const cur = String((c.years && c.years.current) || '');
                    const y = (c.years && c.years.byLabel && cur && c.years.byLabel[cur]) ? c.years.byLabel[cur] : { students: [], classes: [] };
                    return normalizeSettings({
                        schoolName: c.core.schoolName,
                        domain: c.core.domain,
                        subjects: c.core.subjects,
                        arges: c.core.arges,
                        teachers: c.core.teachers,
                        administration: c.core.administration,
                        admin: c.core.admin,
                        adminRoles: c.core.adminRoles,
                        students: (y.students || []).map(function (s) {
                            const row = {
                                id: s.id,
                                klasse: s.klasse,
                                name: s.name,
                                email: s.email,
                                guardianIds: Array.isArray(s.guardianIds) ? s.guardianIds.slice() : []
                            };
                            if (Array.isArray(y.guardians) && row.guardianIds.length) {
                                const byId = new Map(
                                    y.guardians.map(function (g) {
                                        return [g.id, g];
                                    })
                                );
                                row.parentPairs = row.guardianIds
                                    .map(function (gid) {
                                        const g = byId.get(gid);
                                        return g ? { name: g.name || '', email: g.email || '' } : null;
                                    })
                                    .filter(Boolean);
                            }
                            return row;
                        }),
                        studentCouncil: Array.isArray(y.studentCouncil)
                            ? y.studentCouncil.map(function (s) {
                                  return {
                                      klasse: s.klasse,
                                      name: s.name,
                                      email: s.email
                                  };
                              })
                            : [],
                        sgaMode: c.core.sgaMode,
                        sga: c.core.sga,
                        classes: y.classes
                    });
                }
            }
        } catch {
            // ignore
        }
        const raw = loadRaw();
        return normalizeSettings(raw || {});
    }

    function getTeacherEmailMap() {
        const s = load();
        const map = {};
        s.teachers.forEach((t) => {
            if (t.code && t.email) map[t.code] = t.email;
        });
        return map;
    }

    function parseDelimitedLines(text) {
        const lines = String(text || '').split(/\r\n|\n|\r/);
        const out = [];
        lines.forEach((line) => {
            const t = normStr(line);
            if (!t || t.startsWith('#')) return;
            const parts = t
                .split(/[;\t,|]/)
                .map((x) => normStr(x))
                .filter(Boolean);
            if (!parts.length) return;
            out.push(parts);
        });
        return out;
    }

    function parseLinesToSubjects(text) {
        const out = [];
        parseDelimitedLines(text).forEach((parts) => {
            const code = normCode(parts[0] || '');
            const name = normStr(parts.slice(1).join(' '));
            if (!code) return;
            out.push({ code, name });
        });
        return out;
    }

    function parseLinesToTeachers(text) {
        const out = [];
        parseDelimitedLines(text).forEach((parts) => {
            const code = normCode(parts[0] || '');
            const name = normStr(parts[1] || '');
            const email = normStr(parts[2] || '').toLowerCase();
            if (!code) return;
            out.push({ code, name, email });
        });
        return out;
    }

    function parseLinesToStudents(text) {
        const out = [];
        const extractPairs =
            window.ms365ElternGuardians && typeof window.ms365ElternGuardians.extractParentPairsFromParts === 'function'
                ? window.ms365ElternGuardians.extractParentPairsFromParts
                : null;
        String(text || '')
            .split(/\r\n|\n|\r/)
            .forEach((line) => {
                const t = normStr(line);
                if (!t || t.startsWith('#')) return;
                const parts = t.split(/[;\t,|]/).map((x) => normStr(x));
                while (parts.length < 3) parts.push('');
                const klasse = parts[0] || '';
                const name = parts[1] || '';
                const email = (parts[2] || '').toLowerCase();
                if (!klasse && !name && !email) return;
                const row = { klasse, name, email };
                const pairs = extractPairs ? extractPairs(parts) : [];
                if (pairs.length) row.parentPairs = pairs;
                out.push(row);
            });
        return out;
    }

    function parseLinesToStudentCouncil(text) {
        const out = [];
        String(text || '')
            .split(/\r\n|\n|\r/)
            .forEach((line) => {
                const t = normStr(line);
                if (!t || t.startsWith('#')) return;
                const parts = t.split(/[;\t,|]/).map((x) => normStr(x));
                while (parts.length < 3) parts.push('');
                const klasse = parts[0] || '';
                const name = parts[1] || '';
                const email = (parts[2] || '').toLowerCase();
                if (!klasse && !name && !email) return;
                out.push({ klasse, name, email });
            });
        return out;
    }

    function parseLinesToClasses(text) {
        const out = [];
        parseDelimitedLines(text).forEach((parts) => {
            const code = normCode(parts[0] || '');
            // Unterstützte Formate:
            // - code;name;headName;headEmail (alt)
            // - code;year;name;headName;headEmail (neu)
            // - code;name;year;headName;headEmail (tolerant)
            let year = '';
            let name = '';
            let headName = '';
            let headEmail = '';

            if (parts.length >= 2 && /^\d{4}$/.test(normStr(parts[1] || ''))) {
                year = normStr(parts[1] || '');
                name = normStr(parts[2] || '');
                headName = normStr(parts[3] || '');
                headEmail = normStr(parts[4] || '').toLowerCase();
            } else if (parts.length >= 3 && /^\d{4}$/.test(normStr(parts[2] || ''))) {
                name = normStr(parts[1] || '');
                year = normStr(parts[2] || '');
                headName = normStr(parts[3] || '');
                headEmail = normStr(parts[4] || '').toLowerCase();
            } else {
                name = normStr(parts[1] || '');
                headName = normStr(parts[2] || '');
                headEmail = normStr(parts[3] || '').toLowerCase();
            }

            const y = /^\d{4}$/.test(year) ? year : '';
            if (!code && !name && !y && !headName && !headEmail) return;
            const stableMailNickname = deriveClassStableMailNickname(y, code);
            out.push({ code, name, year: y, headName, headEmail, stableMailNickname });
        });
        return out;
    }

    function parseLinesToSga(text) {
        const out = [];
        parseDelimitedLines(text).forEach((parts) => {
            const scopeRaw = normStr(parts[0] || '').toLowerCase();
            const scope =
                scopeRaw === 'lehrer' || scopeRaw === 'teacher'
                    ? 'teacher'
                    : scopeRaw === 'schueler' || scopeRaw === 'schüler' || scopeRaw === 'student'
                      ? 'student'
                      : scopeRaw === 'extern' || scopeRaw === 'external'
                        ? 'external'
                        : '';
            const name = normStr(parts[1] || '');
            const email = normStr(parts[2] || '').toLowerCase();
            if (!scope && !name && !email) return;
            out.push({ scope, name, email });
        });
        return out;
    }

    // Public API (kompatibel zu bisher)
    window.ms365TenantSettingsLoad = load;
    window.ms365TenantSettingsSave = save;
    window.ms365TenantSettingsGetTeacherEmailMap = getTeacherEmailMap;
    window.ms365TenantSettingsParseSubjectsLines = parseLinesToSubjects;
    window.ms365TenantSettingsParseArgesLines = function (text) {
        const out = [];
        parseDelimitedLines(text).forEach((parts) => {
            const code = normCode(parts[0] || '');
            const name = normStr(parts[1] || '');
            const subjRaw = normStr(parts.slice(2).join(' '));
            const subjects = subjRaw
                ? subjRaw
                      .split(/[,\s|]+/)
                      .map((x) => normCode(x))
                      .filter(Boolean)
                : [];
            if (!code) return;
            out.push({ code, name, subjects });
        });
        return out;
    };
    window.ms365TenantSettingsParseTeachersLines = parseLinesToTeachers;
    window.ms365TenantSettingsParseAdminLines = function (text) {
        const out = [];
        parseDelimitedLines(text).forEach((parts) => {
            const role = normStr(parts[0] || '');
            const name = normStr(parts[1] || '');
            const email = normStr(parts[2] || '').toLowerCase();
            if (!role && !name && !email) return;
            out.push({ role, name, email });
        });
        return out;
    };
    window.ms365TenantSettingsParseAdminGroupsLines = function (text) {
        const raw = String(text || '');
        const low = raw.toLowerCase();
        if (low.includes('# rollen') || low.includes('# personen') || low.includes('[rollen]') || low.includes('[personen]')) {
            const roleLines = [];
            const peopleLines = [];
            let mode = 'roles';
            raw.split(/\r\n|\n|\r/).forEach(function (line) {
                const t = normStr(line);
                if (!t) return;
                const lineLow = t.toLowerCase();
                if (lineLow === '# rollen' || lineLow === '[rollen]' || lineLow === 'rollen:') {
                    mode = 'roles';
                    return;
                }
                if (lineLow === '# personen' || lineLow === '[personen]' || lineLow === 'personen:') {
                    mode = 'people';
                    return;
                }
                if (t.startsWith('#')) return;
                const parts = t.split(/[;\t,|]/).map(function (x) {
                    return normStr(x);
                });
                const looksLikePerson = parts.length >= 3 || t.includes('@');
                if (mode === 'people' || looksLikePerson) peopleLines.push(t);
                else roleLines.push(t);
            });
            return adminRolesAndAdminToGroups(parseLinesToAdminRoles(roleLines.join('\n')), window.ms365TenantSettingsParseAdminLines(peopleLines.join('\n')));
        }
        const groupMap = new Map();
        parseDelimitedLines(text).forEach(function (parts) {
            let code = '';
            let name = '';
            let personName = '';
            let email = '';
            if (parts.length >= 4) {
                const first = normStr(parts[0] || '');
                const second = normStr(parts[1] || '');
                const firstLooksLikeCode = !!first && first === normCode(first) && !/\s/.test(first);
                const secondLooksLikeCode = !!second && second === normCode(second) && !/\s/.test(second);
                if (firstLooksLikeCode && !secondLooksLikeCode) {
                    code = normCode(first);
                    name = second;
                } else {
                    name = first;
                    code = normCode(second);
                }
                personName = normStr(parts[2] || '');
                email = normStr(parts[3] || '').toLowerCase();
            } else if (parts.length === 3) {
                name = normStr(parts[0] || '');
                personName = normStr(parts[1] || '');
                email = normStr(parts[2] || '').toLowerCase();
                code = adminRoleCodeFromName(name);
            } else if (parts.length === 2) {
                code = normCode(parts[0] || '');
                name = normStr(parts[1] || '');
            } else {
                name = normStr(parts[0] || '');
                code = adminRoleCodeFromName(name);
            }
            if (!name && !code && !personName && !email) return;
            const key = (name || code).toLowerCase();
            if (!groupMap.has(key)) {
                groupMap.set(key, {
                    code: code || adminRoleCodeFromName(name),
                    name: name || code,
                    people: []
                });
            }
            const group = groupMap.get(key);
            if (code && !group.code) group.code = code;
            if (name && !group.name) group.name = name;
            if (personName || email) {
                group.people.push({ name: personName, email: email });
            }
        });
        return normalizeGroupedAdministrationEntries(Array.from(groupMap.values()));
    };
    window.ms365TenantSettingsAdminGroupsToLines = function (groups) {
        const lines = [];
        (Array.isArray(groups) ? groups : []).forEach(function (group) {
            const code = normStr(group && group.code);
            const name = normStr(group && group.name);
            const people = Array.isArray(group && group.people) ? group.people : [];
            if (!people.length) {
                lines.push([name, code, '', ''].join(';'));
                return;
            }
            people.forEach(function (person) {
                lines.push(
                    [name, code, normStr(person && person.name), normStr(person && person.email).toLowerCase()].join(';')
                );
            });
        });
        return lines.join('\n').trim();
    };
    window.ms365TenantSettingsAdminRolesAndAdminToGroups = adminRolesAndAdminToGroups;
    window.ms365TenantSettingsParseAdminRolesLines = parseLinesToAdminRoles;
    window.ms365TenantSettingsDefaultAdminRoles = defaultAdminRoleCatalog;
    window.ms365TenantSettingsNormalizeAdminRoles = normalizeAdminRoleCatalog;
    window.ms365TenantSettingsRenameAdminRole = renameAdminRole;
    window.ms365TenantSettingsPersonMatchesAdminRole = personMatchesAdminRole;
    window.ms365TenantSettingsAdminRoleCodeFromName = adminRoleCodeFromName;
    window.ms365TenantSettingsParseStudentsLines = parseLinesToStudents;
    window.ms365TenantSettingsParseStudentCouncilLines = parseLinesToStudentCouncil;
    window.ms365TenantSettingsParseClassesLines = parseLinesToClasses;
    window.ms365TenantSettingsParseSgaLines = parseLinesToSga;
    window.ms365DeriveClassStableMailNickname = deriveClassStableMailNickname;
    window.ms365GetClassNickSchema = getClassNickSchema;
    window.ms365SaveClassNickSchema = saveClassNickSchema;
})();

