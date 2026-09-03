/**
 * Import Schüler + Erziehungsberechtigte aus Schulinformationssystemen (SIS)
 * und der MS365-Eigenvorlage (CSV/XLSX).
 *
 * Bekannte Quellen (Dokumentation / Praxis):
 * - Sokrates: Dynamische Suche Standard 111 / Elternabfrage → CSV/XLSX
 *   (u. a. digbi.net, eduFLOW-Wiki, Untis-Elternimport aus Sokrates)
 * - WebUntis: Elternstammdaten-CSV mit Schüler-ID + Eltern-E-Mail
 *   (Untis PDF „Import Elternstammdaten“)
 * - MS365-Vorlage: Klasse;Name;E-Mail;Eltern1;Eltern1Mail;…
 */
(function () {
    'use strict';

    function normStr(v) {
        return String(v ?? '').trim();
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
            .replace(/[^a-z0-9]/g, '');
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

    function findIdx(map, candidates, nth) {
        const n = typeof nth === 'number' ? nth : 0;
        for (let c = 0; c < candidates.length; c++) {
            const key = normHeaderKey(candidates[c]);
            const arr = map.all.get(key);
            if (arr && arr.length > n) return arr[n];
            if (n === 0 && map.first.has(key)) return map.first.get(key);
        }
        return -1;
    }

    function cell(row, idx) {
        if (idx < 0 || !row) return '';
        return normStr(row[idx]);
    }

    function joinPersonName(vorname, familienname) {
        const v = normStr(vorname);
        const f = normStr(familienname);
        if (v && f) return v + ' ' + f;
        return v || f;
    }

    function getFieldFromObject(row, candidates) {
        if (!row || typeof row !== 'object' || Array.isArray(row)) return '';
        const map = new Map();
        Object.keys(row).forEach(function (k) {
            map.set(normHeaderKey(k), row[k]);
        });
        for (let i = 0; i < candidates.length; i++) {
            const v = map.get(normHeaderKey(candidates[i]));
            if (v != null && String(v).trim() !== '') return String(v).trim();
        }
        return '';
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

    function studentKey(rec) {
        const em = normEmail(rec.email);
        if (em) return 'e:' + em;
        const ext = normStr(rec.externalId);
        if (ext) return 'x:' + ext.toLowerCase();
        return (
            'n:' +
            String(rec.klasse || '')
                .toLowerCase() +
            '|' +
            String(rec.name || '')
                .toLowerCase()
        );
    }

    function mergeStudentRecords(list) {
        const by = new Map();
        (list || []).forEach(function (raw) {
            if (!raw) return;
            const klasse = normStr(raw.klasse);
            const name = normStr(raw.name);
            const email = normEmail(raw.email);
            if (!klasse && !name && !email) return;
            const rec = {
                klasse: klasse,
                name: name,
                email: email,
                externalId: normStr(raw.externalId),
                parentPairs: []
            };
            (Array.isArray(raw.parentPairs) ? raw.parentPairs : []).forEach(function (p) {
                pushParent(rec.parentPairs, p && p.name, p && p.email, p && p.phone);
            });
            const key = studentKey(rec);
            if (!by.has(key)) {
                by.set(key, rec);
                return;
            }
            const cur = by.get(key);
            if (!cur.klasse && rec.klasse) cur.klasse = rec.klasse;
            if (!cur.name && rec.name) cur.name = rec.name;
            if (!cur.email && rec.email) cur.email = rec.email;
            if (!cur.externalId && rec.externalId) cur.externalId = rec.externalId;
            rec.parentPairs.forEach(function (p) {
                pushParent(cur.parentPairs, p.name, p.email, p.phone);
            });
        });
        return Array.from(by.values());
    }

    function recordsToSemicolonLines(records) {
        return mergeStudentRecords(records)
            .map(function (r) {
                const parts = [r.klasse || '', r.name || '', r.email || ''];
                (r.parentPairs || []).forEach(function (p) {
                    parts.push(p.name || '');
                    parts.push(p.email || '');
                });
                return parts.join(';');
            })
            .filter(Boolean)
            .join('\n');
    }

    /** MS365 / universelle Vorlage (Objektzeilen aus sheet_to_json) */
    function mapMs365ObjectRows(rows) {
        const out = [];
        (rows || []).forEach(function (r) {
            const klasse = getFieldFromObject(r, ['klasse', 'class', 'zug', 'gruppe', 'k']);
            let name = getFieldFromObject(r, ['name', 'schueler', 'schüler', 'vollname', 'displayname', 'schuelername']);
            let email = getFieldFromObject(r, ['e-mail', 'email', 'mail', 'upn', 'schuelermail', 'schülermail']);
            if (name.includes('@') && (!email || !email.includes('@'))) {
                email = name;
                name = '';
            }
            const parentPairs = [];
            pushParent(
                parentPairs,
                getFieldFromObject(r, ['eltern1', 'erziehungsberechtigte1', 'mutter', 'parent1', 'guardian1', 'sorge1name']),
                getFieldFromObject(r, [
                    'eltern1mail',
                    'eltern1-mail',
                    'eltern1email',
                    'parent1email',
                    'guardian1email',
                    'muttermail',
                    'sorge1mail',
                    'sorge1email'
                ])
            );
            pushParent(
                parentPairs,
                getFieldFromObject(r, ['eltern2', 'erziehungsberechtigte2', 'vater', 'parent2', 'guardian2', 'sorge2name']),
                getFieldFromObject(r, [
                    'eltern2mail',
                    'eltern2-mail',
                    'eltern2email',
                    'parent2email',
                    'guardian2email',
                    'vatermail',
                    'sorge2mail',
                    'sorge2email'
                ])
            );
            // WebUntis-ähnliche Einzeilen-Felder
            pushParent(
                parentPairs,
                joinPersonName(
                    getFieldFromObject(r, ['elternvorname', 'erziehungsberechtigtervorname', 'guardianfirstname']),
                    getFieldFromObject(r, ['elternfamilienname', 'elternnachname', 'erziehungsberechtigterfamilienname', 'guardianlastname'])
                ),
                getFieldFromObject(r, ['elternmail', 'elternemail', 'erziehungsberechtigtermail', 'guardianemail', 'mailadresse'])
            );
            if (!klasse && !name && !email && !parentPairs.length) return;
            out.push({
                klasse: klasse,
                name: name,
                email: normEmail(email),
                externalId: getFieldFromObject(r, ['schuelerkennzahl', 'schülerkennzahl', 'schluessel', 'schlüssel', 'externalid', 'schuelerid', 'id']),
                parentPairs: parentPairs
            });
        });
        return mergeStudentRecords(out);
    }

    /**
     * Sokrates / WebUntis-Elternabfrage als AOA (Headerzeile + Daten).
     * Doppelte Spaltennamen (Familienname/Vorname Schüler vs. Eltern) über Index.
     */
    function mapSokratesOrUntisAoa(aoa) {
        const rows = Array.isArray(aoa) ? aoa : [];
        if (rows.length < 2) return [];
        const headers = rows[0];
        const map = headerIndexMap(headers);

        const iKlasse = findIdx(map, ['klasse', 'class', 'klassenname', 'zug']);
        const iKennzahl = findIdx(map, [
            'schuelerkennzahl',
            'schülerkennzahl',
            'schluessel',
            'schlüssel',
            'schluesselextern',
            'schlüssel(extern,schüler)',
            'schluessel(extern,schueler)',
            'externeid',
            'studentid',
            'schuelerid'
        ]);
        const iStudFam = findIdx(map, ['familienname', 'nachname', 'lastname'], 0);
        const iStudVor = findIdx(map, ['vorname', 'vornamen', 'rufname', 'firstname'], 0);
        const iStudMail = findIdx(map, ['schuelermail', 'schülermail', 'schueleremail', 'emailschueler', 'mailschueler', 'upn']);

        const iParFam = findIdx(map, ['familienname', 'nachname', 'lastname'], 1);
        const iParVor = findIdx(map, ['vorname', 'vornamen', 'rufname', 'firstname'], 1);
        const iParMail = findIdx(map, [
            'mailadresse',
            'mail',
            'email',
            'e-mail',
            'elternmail',
            'elternemail',
            'erziehungsberechtigtermail'
        ]);
        const iParMail2 = findIdx(map, ['mailadresse', 'mail', 'email', 'e-mail'], 1);
        const iParPhone = findIdx(map, ['mobiltelefon', 'mobil', 'telefon', 'telefonnr2', 'telefonnr', 'phone']);

        // Fallback: nur eine Namensspalte „Name“ für Schüler
        const iStudName = findIdx(map, ['name', 'schueler', 'schüler', 'schuelername', 'displayname']);

        const out = [];
        for (let r = 1; r < rows.length; r++) {
            const row = rows[r];
            if (!row || !row.length) continue;
            const klasse = cell(row, iKlasse);
            const externalId = cell(row, iKennzahl);
            let name = joinPersonName(cell(row, iStudVor), cell(row, iStudFam));
            if (!name) name = cell(row, iStudName);
            const email = normEmail(cell(row, iStudMail));

            const parentPairs = [];
            const pName = joinPersonName(
                cell(row, iParVor >= 0 ? iParVor : -1),
                cell(row, iParFam >= 0 ? iParFam : -1)
            );
            const pMail = cell(row, iParMail);
            const pMailAlt = cell(row, iParMail2);
            pushParent(parentPairs, pName, pMail, cell(row, iParPhone));
            if (pMailAlt && normEmail(pMailAlt) !== normEmail(pMail)) {
                pushParent(parentPairs, '', pMailAlt, '');
            }

            // Wide format: Eltern1/Eltern2 already unique headers → also try object-like from header names
            // (handled better via mapMs365 when sheet_to_json used)

            if (!klasse && !name && !email && !externalId && !parentPairs.length) continue;
            out.push({ klasse: klasse, name: name, email: email, externalId: externalId, parentPairs: parentPairs });
        }
        return mergeStudentRecords(out);
    }

    function detectSourceFromHeaders(headers) {
        const keys = (headers || []).map(normHeaderKey).filter(Boolean);
        const set = new Set(keys);
        const has = function () {
            for (let i = 0; i < arguments.length; i++) {
                if (set.has(normHeaderKey(arguments[i]))) return true;
            }
            return false;
        };
        const famCount = keys.filter(function (k) {
            return k === 'familienname' || k === 'nachname';
        }).length;
        const vorCount = keys.filter(function (k) {
            return k === 'vorname' || k === 'vornamen' || k === 'rufname';
        }).length;

        if (has('eltern1mail', 'eltern1email', 'parent1email') || (has('klasse') && has('name') && has('e-mail', 'email'))) {
            if (has('eltern1', 'eltern1mail', 'eltern2mail')) return 'ms365';
        }
        if (has('schuelerkennzahl', 'schülerkennzahl') || (has('mailadresse') && (famCount >= 1 || has('klasse')))) {
            return 'sokrates';
        }
        if (has('schluessel', 'schlüssel', 'schluesselextern') || has('guardianemail', 'elternmail')) {
            return 'webuntis';
        }
        if (famCount >= 2 || vorCount >= 2) return 'sokrates';
        if (has('klasse') && (has('name') || has('familienname'))) return 'ms365';
        return 'auto';
    }

    function detectSourceFromAoa(aoa) {
        if (!aoa || !aoa.length) return 'auto';
        return detectSourceFromHeaders(aoa[0]);
    }

    /**
     * @param {{ aoa?: any[][], objectRows?: object[], source?: string }} input
     * @returns {{ source: string, records: array, lines: string, meta: object }}
     */
    function importStudentsAndGuardians(input) {
        const inp = input && typeof input === 'object' ? input : {};
        let source = String(inp.source || 'auto').trim().toLowerCase();
        const aoa = Array.isArray(inp.aoa) ? inp.aoa : null;
        const objectRows = Array.isArray(inp.objectRows) ? inp.objectRows : null;

        if (source === 'auto') {
            if (aoa && aoa.length) source = detectSourceFromAoa(aoa);
            else if (objectRows && objectRows.length) source = detectSourceFromHeaders(Object.keys(objectRows[0] || {}));
        }

        let records = [];
        if (source === 'sokrates' || source === 'webuntis') {
            if (aoa && aoa.length) records = mapSokratesOrUntisAoa(aoa);
            else if (objectRows) records = mapMs365ObjectRows(objectRows);
        } else {
            if (objectRows) records = mapMs365ObjectRows(objectRows);
            else if (aoa && aoa.length) {
                // AOA → Objekte über Header
                const headers = aoa[0] || [];
                const objs = [];
                for (let i = 1; i < aoa.length; i++) {
                    const row = aoa[i] || [];
                    const o = {};
                    headers.forEach(function (h, idx) {
                        const key = normStr(h);
                        if (!key) return;
                        // bei Duplikaten: erste gewinnt für Objektmodus; Sokrates besser über AOA-Parser
                        if (o[key] == null || o[key] === '') o[key] = row[idx];
                    });
                    objs.push(o);
                }
                const maybeSok = detectSourceFromAoa(aoa);
                if (maybeSok === 'sokrates' || maybeSok === 'webuntis') {
                    source = maybeSok;
                    records = mapSokratesOrUntisAoa(aoa);
                } else {
                    records = mapMs365ObjectRows(objs);
                    source = 'ms365';
                }
            }
        }

        records = mergeStudentRecords(records);
        const withParents = records.filter(function (r) {
            return (r.parentPairs || []).length > 0;
        }).length;
        return {
            source: source,
            records: records,
            lines: recordsToSemicolonLines(records),
            meta: {
                studentCount: records.length,
                withParents: withParents,
                parentMails: records.reduce(function (n, r) {
                    return n + (r.parentPairs || []).length;
                }, 0)
            }
        };
    }

    function ms365TemplateAoa() {
        return [
            ['Klasse', 'Name', 'E-Mail', 'Eltern1', 'Eltern1Mail', 'Eltern2', 'Eltern2Mail'],
            [
                '1A',
                'Anna Beispiel',
                'anna.beispiel@schule.at',
                'Maria Beispiel',
                'maria.beispiel@mail.com',
                'Thomas Beispiel',
                'thomas.beispiel@mail.com'
            ],
            ['1A', 'Ben Demo', 'ben.demo@schule.at', 'Eva Demo', 'eva.demo@mail.com', '', ''],
            ['1B', 'Dave Grohl', 'dave.grohl@schule.at', 'Jane Grohl', 'jane@mail.com', 'John Grohl', 'john@mail.com']
        ];
    }

    function anleitungAoa() {
        return [
            ['Thema', 'Hinweis'],
            ['MS365-Vorlage', 'Blatt „Schueler_Eltern“: Klasse; Name; Schüler-E-Mail; optional Eltern1/2 mit Mail. UTF-8 CSV oder XLSX.'],
            [
                'Sokrates',
                'Laufendes Schuljahr → Dynamische Suche → Standard → 111 Aktive Schüler (ErzBer) bzw. Elternabfrage. Export CSV/XLSX. Typische Spalten: Klasse, Schülerkennzahl, Familienname, Vorname, … Mailadresse (Eltern).'
            ],
            [
                'WebUntis',
                'Elternstammdaten-CSV: Schüler-ID (Schlüssel extern) + Eltern Vor-/Familienname + E-Mail. Schüler sollten in der App bereits Klasse/Name/Mail haben oder in derselben Datei stehen.'
            ],
            ['Datenschutz', 'Elternmails bleiben lokal; Exchange nur als GAL-versteckte Mail Contacts + DL mit versteckter Mitgliedschaft (Eltern-Verteiler).'],
            ['Trenner CSV', 'Semikolon (;) bevorzugt; Komma wird ebenfalls erkannt. BOM/UTF-8 empfohlen.']
        ];
    }

    function sokratesBeispielAoa() {
        return [
            [
                'Klasse',
                'Schülerkennzahl',
                'Familienname',
                'Vorname',
                'Titel',
                'Akad. Grad',
                'Vorname',
                'Familienname',
                'Mailadresse',
                'Mobiltelefon'
            ],
            ['1A', '10001', 'Beispiel', 'Anna', '', '', 'Maria', 'Beispiel', 'maria.beispiel@mail.com', '06641234567'],
            ['1A', '10001', 'Beispiel', 'Anna', '', '', 'Thomas', 'Beispiel', 'thomas.beispiel@mail.com', ''],
            ['1B', '10002', 'Grohl', 'Dave', '', '', 'Jane', 'Grohl', 'jane@mail.com', '']
        ];
    }

    function downloadCsv(filename, aoa) {
        const BOM = '\ufeff';
        const lines = (aoa || []).map(function (row) {
            return (row || [])
                .map(function (cell) {
                    const s = String(cell ?? '');
                    if (/[;"\r\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
                    return s;
                })
                .join(';');
        });
        const blob = new Blob([BOM + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename || 'vorlage.csv';
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(function () {
            URL.revokeObjectURL(url);
        }, 250);
        return true;
    }

    function downloadXlsxTemplates() {
        if (typeof XLSX === 'undefined' || !XLSX.utils || typeof XLSX.writeFile !== 'function') return false;
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ms365TemplateAoa()), 'Schueler_Eltern');
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(anleitungAoa()), 'Anleitung');
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sokratesBeispielAoa()), 'Beispiel_Sokrates');
        XLSX.writeFile(wb, 'MS365-Schueler-Eltern-Vorlage.xlsx');
        return true;
    }

    function sourceOptions() {
        return [
            { id: 'auto', label: 'Automatisch erkennen' },
            { id: 'ms365', label: 'MS365-Vorlage (Klasse;Name;E-Mail;Eltern…)' },
            { id: 'sokrates', label: 'Sokrates (Suche 111 / Elternabfrage)' },
            { id: 'webuntis', label: 'WebUntis Eltern-/Schüler-CSV' }
        ];
    }

    function parentEmailsOf(rec) {
        const out = [];
        const pairs = rec && Array.isArray(rec.parentPairs) ? rec.parentPairs : [];
        pairs.forEach(function (p) {
            const em = normEmail(p && p.email);
            if (em && em.indexOf('@') !== -1) out.push(em);
        });
        return out;
    }

    /**
     * Vergleicht lokale Schülerliste mit einem SIS-Import (ohne Graph).
     * @param {array} existingStudents
     * @param {array} incomingRecords
     */
    function diffSisImport(existingStudents, incomingRecords) {
        const existing = Array.isArray(existingStudents) ? existingStudents : [];
        const incoming = Array.isArray(incomingRecords) ? incomingRecords : [];
        const byEmail = new Map();
        const byExt = new Map();
        const byName = new Map();
        existing.forEach(function (s, idx) {
            const em = normEmail(s && s.email);
            const ext = normStr(s && s.externalId).toLowerCase();
            if (em) byEmail.set(em, idx);
            if (ext) byExt.set(ext, idx);
            byName.set(studentKey(s), idx);
        });

        const usedExisting = new Set();
        const added = [];
        const updated = [];
        const unchanged = [];
        const conflicts = [];

        function findExisting(rec) {
            const em = normEmail(rec.email);
            if (em && byEmail.has(em)) return byEmail.get(em);
            const ext = normStr(rec.externalId).toLowerCase();
            if (ext && byExt.has(ext)) return byExt.get(ext);
            const k = studentKey(rec);
            if (byName.has(k)) return byName.get(k);
            return -1;
        }

        incoming.forEach(function (rec) {
            const idx = findExisting(rec);
            if (idx < 0) {
                added.push(rec);
                return;
            }
            usedExisting.add(idx);
            const prev = existing[idx] || {};
            const prevEm = normEmail(prev.email);
            const recEm = normEmail(rec.email);
            const prevExt = normStr(prev.externalId).toLowerCase();
            const recExt = normStr(rec.externalId).toLowerCase();
            if (prevEm && recEm && prevEm === recEm && prevExt && recExt && prevExt !== recExt) {
                conflicts.push({
                    type: 'email-id',
                    email: recEm,
                    summary:
                        (rec.name || recEm) +
                        ': dieselbe E-Mail, aber andere Schülerkennzahl (' +
                        prev.externalId +
                        ' → ' +
                        rec.externalId +
                        ').'
                });
            }
            if (prevExt && recExt && prevExt === recExt && prevEm && recEm && prevEm !== recEm) {
                conflicts.push({
                    type: 'id-email',
                    email: recEm,
                    summary:
                        (rec.name || recExt) +
                        ': dieselbe Kennzahl, aber andere E-Mail (' +
                        prevEm +
                        ' → ' +
                        recEm +
                        ').'
                });
            }
            const klasseChanged = normStr(prev.klasse).toLowerCase() !== normStr(rec.klasse).toLowerCase();
            const nameChanged = normStr(prev.name).toLowerCase() !== normStr(rec.name).toLowerCase();
            const emailChanged = prevEm !== recEm && !!(prevEm || recEm);
            const parentsIn = parentEmailsOf(rec).slice().sort().join('|');
            const parentsPrev = parentEmailsOf(prev).slice().sort().join('|');
            const parentsChanged = parentsIn !== parentsPrev && !!(parentsIn || parentsPrev);
            if (klasseChanged || nameChanged || emailChanged || parentsChanged) {
                updated.push({
                    previous: prev,
                    incoming: rec,
                    klasseChanged: klasseChanged,
                    nameChanged: nameChanged,
                    emailChanged: emailChanged,
                    parentsChanged: parentsChanged
                });
            } else {
                unchanged.push(rec);
            }
        });

        const removed = [];
        existing.forEach(function (s, idx) {
            if (!usedExisting.has(idx)) removed.push(s);
        });

        return {
            added: added,
            updated: updated,
            removed: removed,
            unchanged: unchanged,
            conflicts: conflicts,
            counts: {
                existing: existing.length,
                incoming: incoming.length,
                added: added.length,
                updated: updated.length,
                removed: removed.length,
                unchanged: unchanged.length,
                conflicts: conflicts.length
            }
        };
    }

    /**
     * @param {array} existingStudents
     * @param {array} incomingRecords
     * @param {{ mode?: 'merge'|'replace' }} [opts]
     */
    function applySisImport(existingStudents, incomingRecords, opts) {
        const mode = opts && opts.mode === 'replace' ? 'replace' : 'merge';
        const incoming = Array.isArray(incomingRecords) ? incomingRecords : [];
        if (mode === 'replace') {
            return incoming.map(function (r) {
                return {
                    klasse: normStr(r.klasse),
                    name: normStr(r.name),
                    email: normEmail(r.email),
                    externalId: normStr(r.externalId),
                    parentPairs: Array.isArray(r.parentPairs) ? r.parentPairs : []
                };
            });
        }
        const diff = diffSisImport(existingStudents, incoming);
        const keep = (diff.removed || []).map(function (s) {
            return {
                klasse: normStr(s.klasse),
                name: normStr(s.name),
                email: normEmail(s.email),
                externalId: normStr(s.externalId),
                parentPairs: Array.isArray(s.parentPairs) ? s.parentPairs : []
            };
        });
        const incomingNorm = incoming.map(function (r) {
            return {
                klasse: normStr(r.klasse),
                name: normStr(r.name),
                email: normEmail(r.email),
                externalId: normStr(r.externalId),
                parentPairs: Array.isArray(r.parentPairs) ? r.parentPairs : []
            };
        });
        return keep.concat(incomingNorm);
    }

    function summarizeSisDiff(diff) {
        const c = (diff && diff.counts) || {};
        const parts = [];
        parts.push(String(c.added || 0) + ' neu');
        parts.push(String(c.updated || 0) + ' geändert');
        parts.push(String(c.unchanged || 0) + ' unverändert');
        parts.push(String(c.removed || 0) + ' nur lokal');
        if (c.conflicts) parts.push(String(c.conflicts) + ' Konflikt(e)');
        return parts.join(' · ');
    }

    window.ms365SchoolSisImport = {
        normHeaderKey,
        importStudentsAndGuardians,
        mapMs365ObjectRows,
        mapSokratesOrUntisAoa,
        detectSourceFromHeaders,
        detectSourceFromAoa,
        recordsToSemicolonLines,
        mergeStudentRecords,
        studentKey,
        diffSisImport,
        applySisImport,
        summarizeSisDiff,
        ms365TemplateAoa,
        anleitungAoa,
        sokratesBeispielAoa,
        downloadCsv,
        downloadXlsxTemplates,
        sourceOptions
    };
})();
