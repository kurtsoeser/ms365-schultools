/**
 * WebUntis-Stammdaten-Exporte: Student_*.xls, LegalGuardian_*.xls, TeacherSalary_*.xls
 * (klassische OLE-XLS, von SheetJS lesbar).
 */
(function () {
    'use strict';

    function normStr(v) {
        return String(v == null ? '' : v).trim();
    }

    function normEmail(v) {
        return normStr(v).toLowerCase();
    }

    function normHeaderKey(k) {
        return String(k ?? '')
            .trim()
            .toLowerCase()
            .replace(/\s+/g, '')
            .replace(/ä/g, 'ae')
            .replace(/ö/g, 'oe')
            .replace(/ü/g, 'ue')
            .replace(/ß/g, 'ss')
            .replace(/[^a-z0-9.]/g, '');
    }

    function headerIndexMap(headerRow) {
        const first = new Map();
        const all = new Map();
        (headerRow || []).forEach(function (h, i) {
            const key = normHeaderKey(h);
            if (!key) return;
            if (!first.has(key)) first.set(key, i);
            if (!all.has(key)) all.set(key, []);
            all.get(key).push(i);
        });
        return { first: first, all: all };
    }

    function findIdx(map, candidates) {
        for (let c = 0; c < candidates.length; c++) {
            const key = normHeaderKey(candidates[c]);
            if (map.first.has(key)) return map.first.get(key);
        }
        return -1;
    }

    function cell(row, idx) {
        if (idx < 0 || !row) return '';
        return normStr(row[idx]);
    }

    function parseDateDe(s) {
        const t = normStr(s);
        if (!t) return null;
        const m = t.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
        if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
        const d = new Date(t);
        return isNaN(d.getTime()) ? null : d;
    }

    function isActiveExitDate(exitDateStr, asOf) {
        const d = parseDateDe(exitDateStr);
        if (!d) return true;
        const ref = asOf instanceof Date ? asOf : new Date();
        return d.getTime() >= ref.getTime();
    }

    function normalizeAddressKey(street, postCode, city) {
        const s = normStr(street)
            .toLowerCase()
            .replace(/straße/g, 'strasse')
            .replace(/str\./g, 'strasse')
            .replace(/[^a-z0-9]/g, '');
        const p = normStr(postCode).replace(/\s+/g, '');
        const c = normStr(city)
            .toLowerCase()
            .replace(/[^a-z0-9]/g, '');
        if (!s && !p && !c) return '';
        return [p, c, s].join('|');
    }

    function normalizePersonNameKey(fore, last) {
        return [normStr(fore), normStr(last)]
            .join(' ')
            .toLowerCase()
            .replace(/\s+/g, ' ')
            .trim();
    }

    /**
     * @param {string[]} headers
     * @returns {'student'|'guardian'|'teacher'|''}
     */
    function detectExportKindFromHeaders(headers) {
        const keys = (headers || []).map(normHeaderKey).filter(Boolean);
        const set = new Set(keys);
        const has = function () {
            for (let i = 0; i < arguments.length; i++) {
                if (set.has(normHeaderKey(arguments[i]))) return true;
            }
            return false;
        };

        if (has('studentinternalid', 'studentexternalid') || (has('studentlastname') && has('studentfirstname'))) {
            return 'guardian';
        }
        if (has('lehrkraft') && (has('familienname') || has('personalnummer'))) {
            return 'teacher';
        }
        if (
            (has('longname') && has('forename') && (has('klasse.name') || has('klassename') || has('klasse'))) ||
            (has('externkey') && has('forename') && has('longname'))
        ) {
            return 'student';
        }
        return '';
    }

    function detectExportKindFromAoa(aoa) {
        if (!aoa || !aoa.length) return '';
        return detectExportKindFromHeaders(aoa[0]);
    }

    function parseStudentsAoa(aoa, opts) {
        const o = opts && typeof opts === 'object' ? opts : {};
        const includeExited = !!o.includeExited;
        const asOf = o.asOf || new Date();
        const rows = Array.isArray(aoa) ? aoa : [];
        if (rows.length < 2) return [];
        const map = headerIndexMap(rows[0]);

        const iName = findIdx(map, ['name']);
        const iLong = findIdx(map, ['longname', 'familienname', 'lastname', 'nachname']);
        const iFore = findIdx(map, ['forename', 'vorname', 'firstname']);
        const iKlasse = findIdx(map, ['klasse.name', 'klassename', 'klasse', 'class']);
        const iId = findIdx(map, ['id']);
        const iExt = findIdx(map, ['externkey', 'externalkey', 'schluessel', 'schlüssel']);
        const iExit = findIdx(map, ['exitdate', 'austrittsdatum']);
        const iEntry = findIdx(map, ['entrydate', 'eintrittsdatum']);
        const iEmail = findIdx(map, ['address.email', 'email', 'mail']);
        const iStreet = findIdx(map, ['address.street', 'street', 'strasse', 'straße']);
        const iCity = findIdx(map, ['address.city', 'city', 'ort']);
        const iPost = findIdx(map, ['address.postcode', 'postcode', 'plz']);
        const iPhone = findIdx(map, ['address.phone', 'phone', 'telefon']);
        const iMobile = findIdx(map, ['address.mobile', 'mobile', 'mobil']);

        const out = [];
        for (let r = 1; r < rows.length; r++) {
            const row = rows[r];
            if (!row || !row.length) continue;
            const longName = cell(row, iLong);
            const foreName = cell(row, iFore);
            const klasse = cell(row, iKlasse);
            const untisInternalId = cell(row, iId);
            const untisExternKey = cell(row, iExt);
            const exitDate = cell(row, iExit);
            if (!includeExited && !isActiveExitDate(exitDate, asOf)) continue;
            const name = foreName && longName ? foreName + ' ' + longName : foreName || longName || cell(row, iName);
            const email = normEmail(cell(row, iEmail));
            const street = cell(row, iStreet);
            const city = cell(row, iCity);
            const postCode = cell(row, iPost);
            if (!klasse && !name && !untisInternalId && !untisExternKey) continue;
            out.push({
                klasse: klasse,
                name: name,
                email: email,
                externalId: untisExternKey || untisInternalId,
                untisInternalId: untisInternalId,
                untisExternKey: untisExternKey,
                foreName: foreName,
                longName: longName,
                givenName: foreName,
                surname: longName,
                exitDate: exitDate,
                entryDate: cell(row, iEntry),
                phone: cell(row, iMobile) || cell(row, iPhone),
                address: { street: street, city: city, postCode: postCode },
                addressKey: normalizeAddressKey(street, postCode, city),
                parentPairs: [],
                active: isActiveExitDate(exitDate, asOf)
            });
        }
        return out;
    }

    function parseGuardiansAoa(aoa) {
        const rows = Array.isArray(aoa) ? aoa : [];
        if (rows.length < 2) return [];
        const map = headerIndexMap(rows[0]);

        const iLast = findIdx(map, ['lastname', 'familienname', 'nachname']);
        const iFirst = findIdx(map, ['firstname', 'vorname', 'forename']);
        const iEmail = findIdx(map, ['email', 'mail', 'mailadresse']);
        const iPhone = findIdx(map, ['phone', 'telefon']);
        const iMobile = findIdx(map, ['mobile', 'mobil']);
        const iStuLast = findIdx(map, ['studentlastname']);
        const iStuFirst = findIdx(map, ['studentfirstname']);
        const iStuShort = findIdx(map, ['studentshortname']);
        const iStuInt = findIdx(map, ['studentinternalid']);
        const iStuExt = findIdx(map, ['studentexternalid']);
        const iStreet = findIdx(map, ['addressstreet', 'street']);
        const iCity = findIdx(map, ['addresscity', 'city']);
        const iPost = findIdx(map, ['addresspostcode', 'postcode']);

        const out = [];
        for (let r = 1; r < rows.length; r++) {
            const row = rows[r];
            if (!row || !row.length) continue;
            const firstName = cell(row, iFirst);
            const lastName = cell(row, iLast);
            const name = firstName && lastName ? firstName + ' ' + lastName : firstName || lastName;
            const email = normEmail(cell(row, iEmail));
            const phone = cell(row, iMobile) || cell(row, iPhone);
            const street = cell(row, iStreet);
            const city = cell(row, iCity);
            const postCode = cell(row, iPost);
            const studentInternalId = cell(row, iStuInt);
            const studentExternalId = cell(row, iStuExt);
            const studentFirstName = cell(row, iStuFirst);
            const studentLastName = cell(row, iStuLast);
            if (!name && !email && !studentInternalId && !studentExternalId) continue;
            out.push({
                name: name,
                email: email,
                phone: phone,
                studentInternalId: studentInternalId,
                studentExternalId: studentExternalId,
                studentFirstName: studentFirstName,
                studentLastName: studentLastName,
                studentShortName: cell(row, iStuShort),
                studentNameKey: normalizePersonNameKey(studentFirstName, studentLastName),
                address: { street: street, city: city, postCode: postCode },
                addressKey: normalizeAddressKey(street, postCode, city)
            });
        }
        return out;
    }

    /**
     * TeacherSalary: Primärzeilen mit Kürzel; Folgeseilen (nur Soll/Woche) ignorieren.
     */
    function parseTeachersAoa(aoa, opts) {
        const o = opts && typeof opts === 'object' ? opts : {};
        const includeExited = !!o.includeExited;
        const rows = Array.isArray(aoa) ? aoa : [];
        if (rows.length < 2) return [];
        const map = headerIndexMap(rows[0]);

        const iCode = findIdx(map, ['lehrkraft', 'kuerzel', 'kürzel', 'code']);
        const iLast = findIdx(map, ['familienname', 'nachname', 'lastname']);
        const iFirst = findIdx(map, ['vorname', 'forename', 'firstname']);
        const iTitle = findIdx(map, ['titel', 'title']);
        const iPers = findIdx(map, ['personalnummer', 'personalnr', 'personnelnumber']);
        const iExit = findIdx(map, ['austrittsdatum', 'exitdate']);
        const iEntry = findIdx(map, ['eintrittsdatum', 'entrydate']);
        const iStatus = findIdx(map, ['lehrkraftstatus', 'status']);

        const out = [];
        for (let r = 1; r < rows.length; r++) {
            const row = rows[r];
            if (!row || !row.length) continue;
            const code = normStr(cell(row, iCode)).toUpperCase();
            if (!code) continue;
            const lastName = cell(row, iLast);
            const firstName = cell(row, iFirst);
            const exitDate = cell(row, iExit);
            if (!includeExited && exitDate && !isActiveExitDate(exitDate, o.asOf || new Date())) continue;
            const name = firstName && lastName ? firstName + ' ' + lastName : firstName || lastName;
            out.push({
                code: code,
                name: name,
                email: '',
                externalId: cell(row, iPers),
                foreName: firstName,
                longName: lastName,
                givenName: firstName,
                surname: lastName,
                title: cell(row, iTitle),
                entryDate: cell(row, iEntry),
                exitDate: exitDate,
                status: cell(row, iStatus)
            });
        }
        return out;
    }

    function pushParent(list, name, email, phone) {
        const em = normEmail(email);
        if (!em || em.indexOf('@') === -1) return;
        const nm = normStr(name);
        const ph = normStr(phone);
        for (let i = 0; i < list.length; i++) {
            if (list[i].email === em) {
                if (nm && !list[i].name) list[i].name = nm;
                if (ph && !list[i].phone) list[i].phone = ph;
                return;
            }
        }
        list.push({ name: nm, email: em, phone: ph });
    }

    /**
     * Verknüpft Eltern mit Schüler:innen.
     * Priorität: studentInternalId → studentExternalId → Name+Adresse (nur eindeutig).
     */
    function joinStudentsAndGuardians(students, guardians) {
        const stu = (students || []).map(function (s) {
            return Object.assign({}, s, { parentPairs: Array.isArray(s.parentPairs) ? s.parentPairs.slice() : [] });
        });
        const byInt = new Map();
        const byExt = new Map();
        const byName = new Map();
        const byAddr = new Map();

        stu.forEach(function (s, idx) {
            const iid = normStr(s.untisInternalId);
            const eid = normStr(s.untisExternKey || s.externalId);
            if (iid) byInt.set(iid, idx);
            if (eid) byExt.set(eid, idx);
            const nk = normalizePersonNameKey(s.foreName || s.givenName, s.longName || s.surname);
            if (nk) {
                if (!byName.has(nk)) byName.set(nk, []);
                byName.get(nk).push(idx);
            }
            if (s.addressKey) {
                if (!byAddr.has(s.addressKey)) byAddr.set(s.addressKey, []);
                byAddr.get(s.addressKey).push(idx);
            }
        });

        const unmatched = [];
        const matchCounts = { byInternalId: 0, byExternalId: 0, byNameAddress: 0, unmatched: 0 };

        (guardians || []).forEach(function (g) {
            let idx = -1;
            let method = '';
            const gid = normStr(g.studentInternalId);
            const gex = normStr(g.studentExternalId);
            if (gid && byInt.has(gid)) {
                idx = byInt.get(gid);
                method = 'internalId';
                matchCounts.byInternalId++;
            } else if (gex && byExt.has(gex)) {
                idx = byExt.get(gex);
                method = 'externalId';
                matchCounts.byExternalId++;
            } else {
                const nk = normStr(g.studentNameKey) || normalizePersonNameKey(g.studentFirstName, g.studentLastName);
                let candidates = nk && byName.has(nk) ? byName.get(nk).slice() : [];
                if (g.addressKey && byAddr.has(g.addressKey)) {
                    const addrHits = byAddr.get(g.addressKey);
                    if (candidates.length) {
                        candidates = candidates.filter(function (i) {
                            return addrHits.indexOf(i) !== -1;
                        });
                    } else if (nk) {
                        // Name bekannt, Adresse als Verstärker – ohne Namens-Treffer nicht nur Adresse
                        candidates = [];
                    }
                }
                if (candidates.length === 1) {
                    idx = candidates[0];
                    method = 'nameAddress';
                    matchCounts.byNameAddress++;
                }
            }

            if (idx < 0) {
                matchCounts.unmatched++;
                unmatched.push(g);
                return;
            }
            pushParent(stu[idx].parentPairs, g.name, g.email, g.phone);
            if (!stu[idx]._parentMatchMethods) stu[idx]._parentMatchMethods = [];
            stu[idx]._parentMatchMethods.push(method);
        });

        return {
            students: stu,
            unmatchedGuardians: unmatched,
            matchCounts: matchCounts
        };
    }

    function teachersToSemicolonLines(teachers) {
        return (teachers || [])
            .map(function (t) {
                return [normStr(t.code).toUpperCase(), normStr(t.name), normEmail(t.email)].filter(Boolean).join(';');
            })
            .filter(Boolean)
            .join('\n');
    }

    function applyPersonEmails(records, opts) {
        const api = window.ms365PersonEmailFromName;
        if (!api || typeof api.assignEmails !== 'function') {
            return { records: records || [], generated: 0, conflicts: 0 };
        }
        const assigned = api.assignEmails(records || [], opts || {});
        let generated = 0;
        let conflicts = 0;
        assigned.forEach(function (r) {
            if (r.emailMeta && r.emailMeta.generated) generated++;
            if (r.emailMeta && r.emailMeta.conflict) conflicts++;
        });
        return { records: assigned, generated: generated, conflicts: conflicts };
    }

    /**
     * Kombiniert Student- + LegalGuardian-AOAs zu SIS-Records.
     * @param {{ studentAoa?: any[][], guardianAoa?: any[][], domain?: string, pattern?: string, firstNameMode?: string, includeExited?: boolean, applyEmails?: boolean }} input
     */
    function importStudentsFromWebuntis(input) {
        const inp = input && typeof input === 'object' ? input : {};
        const studentsRaw = parseStudentsAoa(inp.studentAoa || [], { includeExited: inp.includeExited });
        const guardians = parseGuardiansAoa(inp.guardianAoa || []);
        const joined = joinStudentsAndGuardians(studentsRaw, guardians);
        let records = joined.students;
        let emailMeta = { generated: 0, conflicts: 0 };

        if (inp.applyEmails !== false && inp.domain) {
            const em = applyPersonEmails(records, {
                domain: inp.domain,
                pattern: inp.pattern || 'vorname.nachname',
                firstNameMode: inp.firstNameMode || 'first'
            });
            records = em.records;
            emailMeta = { generated: em.generated, conflicts: em.conflicts };
        }

        const withParents = records.filter(function (r) {
            return (r.parentPairs || []).length > 0;
        }).length;

        let lines = '';
        if (window.ms365SchoolSisImport && typeof window.ms365SchoolSisImport.recordsToSemicolonLines === 'function') {
            lines = window.ms365SchoolSisImport.recordsToSemicolonLines(records);
        }

        return {
            source: 'webuntis',
            records: records,
            lines: lines,
            unmatchedGuardians: joined.unmatchedGuardians,
            matchCounts: joined.matchCounts,
            emailMeta: emailMeta,
            meta: {
                studentCount: records.length,
                withParents: withParents,
                parentMails: records.reduce(function (n, r) {
                    return n + (r.parentPairs || []).length;
                }, 0),
                unmatchedGuardians: (joined.unmatchedGuardians || []).length,
                emailsGenerated: emailMeta.generated,
                emailConflicts: emailMeta.conflicts
            }
        };
    }

    /**
     * @param {{ teacherAoa?: any[][], domain?: string, pattern?: string, firstNameMode?: string, applyEmails?: boolean, includeExited?: boolean }} input
     */
    function importTeachersFromWebuntis(input) {
        const inp = input && typeof input === 'object' ? input : {};
        let teachers = parseTeachersAoa(inp.teacherAoa || [], { includeExited: inp.includeExited });
        let emailMeta = { generated: 0, conflicts: 0 };
        if (inp.applyEmails !== false && inp.domain) {
            const em = applyPersonEmails(teachers, {
                domain: inp.domain,
                pattern: inp.pattern || 'vorname.nachname',
                firstNameMode: inp.firstNameMode || 'first'
            });
            teachers = em.records;
            emailMeta = { generated: em.generated, conflicts: em.conflicts };
        }
        return {
            source: 'webuntis',
            teachers: teachers,
            lines: teachersToSemicolonLines(teachers),
            emailMeta: emailMeta,
            meta: {
                teacherCount: teachers.length,
                emailsGenerated: emailMeta.generated,
                emailConflicts: emailMeta.conflicts
            }
        };
    }

    /**
     * Liest mehrere Dateien (bereits als AOA) und sortiert sie nach Typ.
     * @param {{ aoa: any[][], name?: string }[]} sheets
     */
    function classifySheets(sheets) {
        const out = { studentAoa: null, guardianAoa: null, teacherAoa: null, unknown: [] };
        (sheets || []).forEach(function (sh) {
            const kind = detectExportKindFromAoa(sh && sh.aoa);
            if (kind === 'student' && !out.studentAoa) out.studentAoa = sh.aoa;
            else if (kind === 'guardian' && !out.guardianAoa) out.guardianAoa = sh.aoa;
            else if (kind === 'teacher' && !out.teacherAoa) out.teacherAoa = sh.aoa;
            else out.unknown.push({ name: sh && sh.name, kind: kind || 'unknown' });
        });
        return out;
    }

    /**
     * @param {{ sheets: { aoa: any[][], name?: string }[], domain?: string, pattern?: string, includeExited?: boolean }} input
     */
    function importFromSheets(input) {
        const inp = input && typeof input === 'object' ? input : {};
        const classified = classifySheets(inp.sheets || []);
        const result = { classified: classified, students: null, teachers: null };
        if (classified.studentAoa) {
            result.students = importStudentsFromWebuntis({
                studentAoa: classified.studentAoa,
                guardianAoa: classified.guardianAoa || [],
                domain: inp.domain,
                pattern: inp.pattern,
                includeExited: inp.includeExited,
                applyEmails: inp.applyEmails
            });
        } else if (classified.guardianAoa && !classified.studentAoa) {
            result.students = {
                source: 'webuntis',
                records: [],
                lines: '',
                unmatchedGuardians: parseGuardiansAoa(classified.guardianAoa),
                matchCounts: { byInternalId: 0, byExternalId: 0, byNameAddress: 0, unmatched: classified.guardianAoa.length - 1 },
                emailMeta: { generated: 0, conflicts: 0 },
                meta: {
                    studentCount: 0,
                    withParents: 0,
                    parentMails: 0,
                    unmatchedGuardians: Math.max(0, (classified.guardianAoa || []).length - 1),
                    emailsGenerated: 0,
                    emailConflicts: 0,
                    error: 'LegalGuardian-Export ohne Student-Export – bitte beide Dateien wählen.'
                }
            };
        }
        if (classified.teacherAoa) {
            result.teachers = importTeachersFromWebuntis({
                teacherAoa: classified.teacherAoa,
                domain: inp.domain,
                pattern: inp.pattern,
                includeExited: inp.includeExited,
                applyEmails: inp.applyEmails
            });
        }
        return result;
    }

    /**
     * WebUntis Klassen-PDF (PD-Bericht „Klassen“): Spalten Kurzname, Langname, Klassenlehrkraft, Text.
     * Text `.hak` / `.has` = Zweig; Klassenlehrkraft = Lehrkraft-Kürzel.
     */

    function isClassCodeToken(t) {
        const s = normStr(t);
        if (!s) return false;
        if (/^\.\w+$/i.test(s)) return false;
        // 1AK, 2BS, 5EK, FS_BAFEP, …
        if (/^\d+[A-Za-z][A-Za-z0-9_]*$/.test(s)) return true;
        if (/^[A-Za-z][A-Za-z0-9_]{1,20}$/.test(s) && !/^(Schuljahr|Klassen|Seite|WebUntis|Untis|von)$/i.test(s)) {
            return true;
        }
        return false;
    }

    function isTeacherCodeToken(t) {
        const s = normStr(t).toUpperCase();
        if (!s || s.length < 2 || s.length > 10) return false;
        if (/^\d/.test(s)) return false;
        if (/^\./.test(s)) return false;
        // typische Untis-Kürzel: nur Buchstaben
        return /^[A-ZÄÖÜ]{2,10}$/.test(s);
    }

    function isDeptTextToken(t) {
        return /^\.\w+$/i.test(normStr(t));
    }

    function parseSchoolYearFromClassPdfText(text) {
        const m = String(text || '').match(/Schuljahr\s*:?\s*(\d{4})\s*\/\s*(\d{4})/i);
        if (!m) return { label: '', startYear: 0, endYear: 0 };
        return { label: m[1] + '/' + m[2], startYear: Number(m[1]), endYear: Number(m[2]) };
    }

    /**
     * Abschlussjahr aus Schulstufe + Zweig (.hak≈5 Jahre, .has≈3 Jahre).
     * Formel: Endjahr des aktuellen Schuljahrs + (Dauer − Stufe).
     * 5. Klassen (HAK/K): Abschluss = Schuljahres-Endjahr (nicht Startjahr).
     */
    function inferGraduationYear(code, deptText, schoolYearEnd) {
        const end = Number(schoolYearEnd) || 0;
        if (!end) return '';
        // Bei kombinierten Codes (5AK5BK) erste Stufe verwenden
        const m = String(code || '').match(/^(\d+)/);
        if (!m) return '';
        const grade = Number(m[1]);
        if (!grade || grade > 8) return '';
        const dept = normStr(deptText).toLowerCase();
        const codeUpper = String(code || '').toUpperCase();
        // Buchstabenanteil der ersten Klasse (5AK… → AK, 5AS → AS)
        const restMatch = codeUpper.match(/^\d+([A-Z]+)/);
        const rest = restMatch ? restMatch[1] : codeUpper.replace(/^\d+/, '');

        let duration = 5;
        if (dept === '.has' || dept === 'has') duration = 3;
        else if (dept === '.hak' || dept === 'hak') duration = 5;
        else if (/K/.test(rest)) {
            // K im Kürzel → HAK (5 Jahre), auch wenn PDF-Zweig fehlt
            duration = 5;
        } else if (/S$/.test(rest) && !/K/.test(rest)) {
            duration = 3;
        }

        // Abwehr: fälschlich HAS bei 5. HAK-Klasse (K im Code)
        if (grade >= 5 && /K/.test(rest)) duration = 5;

        if (grade > duration) return '';
        // Abschlussjahrgang: Endjahr + (Dauer − Stufe); 5. HAK → Endjahr
        return String(end + (duration - grade));
    }

    /**
     * Parst den Klartext eines WebUntis-Klassen-PDFs (Zeilen: Kurzname, Langname, [.zweig], [Kürzel]).
     * @param {string} text
     * @param {{ skipWithoutTeacher?: boolean, inferYear?: boolean }} [opts]
     */
    function parseClassesFromPdfText(text, opts) {
        const o = opts && typeof opts === 'object' ? opts : {};
        const skipWithoutTeacher = !!o.skipWithoutTeacher;
        const inferYear = o.inferYear !== false;
        const sy = parseSchoolYearFromClassPdfText(text);
        const lines = String(text || '')
            .split(/\r\n|\n|\r/)
            .map(function (l) {
                return normStr(l);
            })
            .filter(Boolean);

        let start = 0;
        for (let i = 0; i < lines.length; i++) {
            if (/^Klassenlehrkraft$/i.test(lines[i]) || (/Kurzname/i.test(lines[i]) && /Langname/i.test(lines[i]))) {
                start = i + 1;
            }
            if (/^Schuljahr/i.test(lines[i])) start = Math.max(start, i + 1);
        }

        const tokens = [];
        for (let i = start; i < lines.length; i++) {
            const l = lines[i];
            if (/^WebUntis/i.test(l) || /^Untis GmbH/i.test(l) || /^Seite\s+\d/i.test(l)) break;
            if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(l)) continue;
            if (/^\d{1,2}:\d{2}$/.test(l)) continue;
            if (/^admin_/i.test(l)) continue;
            tokens.push(l);
        }

        const out = [];
        let i = 0;
        while (i < tokens.length) {
            const kurz = tokens[i];
            if (!isClassCodeToken(kurz)) {
                i++;
                continue;
            }
            i++;
            let lang = '';
            let dept = '';
            let teacherCode = '';
            if (i < tokens.length && (isClassCodeToken(tokens[i]) || /^\d{3,}$/.test(tokens[i]))) {
                lang = tokens[i];
                i++;
            }
            // Klassenlehrkraft nur nach Text-Spalte (.hak / .has)
            if (i < tokens.length && isDeptTextToken(tokens[i])) {
                dept = tokens[i];
                i++;
                if (i < tokens.length && isTeacherCodeToken(tokens[i]) && !/^\d/.test(tokens[i])) {
                    teacherCode = tokens[i].toUpperCase();
                    i++;
                }
            }
            if (skipWithoutTeacher && !teacherCode) continue;

            const code = normStr(kurz).toUpperCase();
            const name = normStr(lang) || code;
            const year = inferYear ? inferGraduationYear(code, dept, sy.endYear) : '';
            out.push({
                code: code,
                name: name,
                year: year,
                headCode: teacherCode,
                headName: '',
                headEmail: '',
                deptText: dept,
                schoolYear: sy.label
            });
        }
        return {
            classes: out,
            schoolYear: sy,
            meta: {
                classCount: out.length,
                withTeacher: out.filter(function (c) {
                    return !!c.headCode;
                }).length
            }
        };
    }

    /**
     * Positionierte PDF-Wörter (pdf.js / pymupdf): { str|text, x, y }.
     */
    function parseClassesFromPdfWords(words, opts) {
        const o = opts && typeof opts === 'object' ? opts : {};
        const list = Array.isArray(words) ? words : [];
        const rows = new Map();
        list.forEach(function (w) {
            const text = normStr(w.str != null ? w.str : w.text);
            if (!text) return;
            const x = Number(w.x != null ? w.x : w.x0);
            const y = Number(w.y != null ? w.y : w.y0);
            if (!isFinite(x) || !isFinite(y)) return;
            if (y < 130 || y > 760) return;
            const yk = Math.round(y);
            if (!rows.has(yk)) rows.set(yk, []);
            rows.get(yk).push({ x: x, text: text });
        });

        // Schuljahr aus allen Wörtern
        let schoolYearEnd = 0;
        let schoolYearLabel = '';
        const allText = list
            .map(function (w) {
                return normStr(w.str != null ? w.str : w.text);
            })
            .join(' ');
        const sy = parseSchoolYearFromClassPdfText(allText);
        schoolYearEnd = sy.endYear;
        schoolYearLabel = sy.label;

        const out = [];
        Array.from(rows.keys())
            .sort(function (a, b) {
                return a - b;
            })
            .forEach(function (yk) {
                const cells = rows.get(yk).slice().sort(function (a, b) {
                    return a.x - b.x;
                });
                let kurz = '';
                let lang = '';
                let teacherCode = '';
                let dept = '';
                cells.forEach(function (c) {
                    if (c.x < 70) kurz = c.text;
                    else if (c.x < 160) lang = c.text;
                    else if (c.x < 300) teacherCode = c.text;
                    else dept = c.text;
                });
                if (!isClassCodeToken(kurz) && !isClassCodeToken(lang)) return;
                const code = normStr(kurz || lang).toUpperCase();
                if (o.skipWithoutTeacher && !isTeacherCodeToken(teacherCode)) return;
                const tCode = isTeacherCodeToken(teacherCode) ? teacherCode.toUpperCase() : '';
                const d = isDeptTextToken(dept) ? dept : '';
                // Sonderzeile FS_BAFEP: Text=BAFEP ohne Lehrer
                out.push({
                    code: code,
                    name: normStr(lang) || code,
                    year: o.inferYear === false ? '' : inferGraduationYear(code, d, schoolYearEnd),
                    headCode: tCode,
                    headName: '',
                    headEmail: '',
                    deptText: d || (isDeptTextToken(dept) ? dept : normStr(dept)),
                    schoolYear: schoolYearLabel
                });
            });

        return {
            classes: out,
            schoolYear: sy,
            meta: { classCount: out.length, withTeacher: out.filter(function (c) { return !!c.headCode; }).length }
        };
    }

    /**
     * Reichert Klassen mit Name/E-Mail aus der Lehrerliste an (Match über Kürzel).
     * @param {array} classes
     * @param {array} teachers [{code,name,email}]
     */
    function enrichClassesWithTeachers(classes, teachers) {
        const byCode = new Map();
        (teachers || []).forEach(function (t) {
            const c = normStr(t && t.code).toUpperCase();
            if (!c) return;
            byCode.set(c, t);
        });
        let matched = 0;
        let missing = 0;
        const out = (classes || []).map(function (cl) {
            const rec = Object.assign({}, cl);
            const code = normStr(rec.headCode).toUpperCase();
            if (!code) return rec;
            const t = byCode.get(code);
            if (t) {
                matched++;
                if (!rec.headName) rec.headName = normStr(t.name);
                if (!rec.headEmail) rec.headEmail = normEmail(t.email);
            } else {
                missing++;
                if (!rec.headName) rec.headName = code;
            }
            return rec;
        });
        return { classes: out, meta: { matched: matched, missingTeacher: missing } };
    }

    function classesToSemicolonLines(classes) {
        return (classes || [])
            .map(function (c) {
                return [
                    normStr(c.code).toUpperCase(),
                    normStr(c.year),
                    normStr(c.name) || normStr(c.code).toUpperCase(),
                    normStr(c.headName),
                    normEmail(c.headEmail)
                ].join(';');
            })
            .filter(Boolean)
            .join('\n');
    }

    /**
     * @param {{ text?: string, words?: array, teachers?: array, skipWithoutTeacher?: boolean, inferYear?: boolean }} input
     */
    function importClassesFromWebuntisPdf(input) {
        const inp = input && typeof input === 'object' ? input : {};
        let parsed;
        if (inp.words && inp.words.length) {
            parsed = parseClassesFromPdfWords(inp.words, {
                skipWithoutTeacher: inp.skipWithoutTeacher,
                inferYear: inp.inferYear
            });
        } else {
            parsed = parseClassesFromPdfText(inp.text || '', {
                skipWithoutTeacher: inp.skipWithoutTeacher,
                inferYear: inp.inferYear
            });
        }
        const enriched = enrichClassesWithTeachers(parsed.classes, inp.teachers || []);
        return {
            source: 'webuntis-class-pdf',
            classes: enriched.classes,
            lines: classesToSemicolonLines(enriched.classes),
            schoolYear: parsed.schoolYear,
            meta: Object.assign({}, parsed.meta, enriched.meta)
        };
    }

    window.ms365WebuntisExportImport = {
        normHeaderKey: normHeaderKey,
        detectExportKindFromHeaders: detectExportKindFromHeaders,
        detectExportKindFromAoa: detectExportKindFromAoa,
        parseStudentsAoa: parseStudentsAoa,
        parseGuardiansAoa: parseGuardiansAoa,
        parseTeachersAoa: parseTeachersAoa,
        joinStudentsAndGuardians: joinStudentsAndGuardians,
        importStudentsFromWebuntis: importStudentsFromWebuntis,
        importTeachersFromWebuntis: importTeachersFromWebuntis,
        teachersToSemicolonLines: teachersToSemicolonLines,
        classifySheets: classifySheets,
        importFromSheets: importFromSheets,
        applyPersonEmails: applyPersonEmails,
        normalizeAddressKey: normalizeAddressKey,
        isActiveExitDate: isActiveExitDate,
        parseClassesFromPdfText: parseClassesFromPdfText,
        parseClassesFromPdfWords: parseClassesFromPdfWords,
        enrichClassesWithTeachers: enrichClassesWithTeachers,
        inferGraduationYear: inferGraduationYear,
        parseSchoolYearFromClassPdfText: parseSchoolYearFromClassPdfText,
        classesToSemicolonLines: classesToSemicolonLines,
        importClassesFromWebuntisPdf: importClassesFromWebuntisPdf
    };
})();
