(function () {
    'use strict';

    const STORAGE_KEY_V2 = 'ms365-schooltool-data-v2';
    /** Schema 4: years.*.guardians, students.id/guardianIds, parentLists */
    const VERSION = 4;
    const SLG_LEGACY_KEY = 'ms365-schueler-lehrer-gruppen-v2';

    function safeJsonParse(s) {
        try {
            return JSON.parse(String(s));
        } catch {
            return null;
        }
    }

    function deepClone(obj) {
        try {
            return JSON.parse(JSON.stringify(obj));
        } catch {
            return obj;
        }
    }

    function currentSchoolYearLabel() {
        const y = new Date().getFullYear();
        return String(y) + '/' + String(y + 1).slice(2);
    }

    function defaultSetup() {
        return {
            wizardStep: 1,
            completedSteps: [],
            finishedAt: null,
            lastVisitedAt: null,
            matched: {
                schuelerGroupId: null,
                lehrerGroupId: null,
                verwaltungGroupId: null,
                sgaGroupId: null,
                studentCouncilGroupId: null
            },
            slgDraft: {
                activeKind: 'schueler',
                slgNewDisplayName: '',
                slgNewMailNick: '',
                slgNewDescription: '',
                slgNewCreateTeam: false,
                /** Besitzer Lehrer-Sammelgruppe: direktion | teachers | manual */
                slgOwnerSourceLehrer: 'direktion',
                slgOwnerManualEmailsLehrer: '',
                /** Besitzer Schüler-Sammelgruppe: direktion | admin | manual */
                slgOwnerSourceSchueler: 'direktion',
                slgOwnerManualEmailsSchueler: ''
            },
            verwaltungDraft: {
                vwNewDisplayName: 'Schulverwaltung',
                vwNewMailNick: 'verwaltung',
                vwNewDescription: '',
                vwNewCreateTeam: false,
                /** Besitzer Verwaltungs-Sammelgruppe: admin | direktion | manual */
                vwOwnerSource: 'admin',
                vwOwnerManualEmails: ''
            },
            /** Kleinbuchstaben/Ziffern; Vorschau/Anlage Fachgruppen (Einrichtungsassistent) */
            subjectGroupMailPrefix: 'fach',
            /** Kleinbuchstaben/Ziffern; Vorschau/Anlage ARGE-Gruppen */
            argeGroupMailPrefix: 'ag',
            /** Eltern-Verteiler: Baustein-Muster (wie Kursteam-Namen) */
            elternClassAliasPattern: [
                { type: 'text', value: 'eltern' },
                { type: 'klasse' }
            ],
            elternClassDisplayPattern: [
                { type: 'text', value: 'Eltern ' },
                { type: 'klasse' }
            ],
            elternYearAliasPattern: [
                { type: 'text', value: 'elternjg' },
                { type: 'year' }
            ],
            elternYearDisplayPattern: [
                { type: 'text', value: 'Eltern JG ' },
                { type: 'year' }
            ],
            /** E‑Mail (Kleinbuchstaben) → Entra-Benutzer (Einrichtungsassistent, optional) */
            directoryMatchByEmail: {},
            /** Klassen-Abkürzung (Großbuchstaben) → M365-Gruppeninfo (Einrichtungsassistent, optional) */
            classGroupMatchByKey: {},
            catalogLinks: [],
            actionLog: [],
            intranetSiteUrl: '',
            intranetHubAt: null,
            sisImportHistory: [],
            elternSetup: { completedSteps: [], lastDiagnoseAt: null }
        };
    }

    function normalizeElternNamePattern(raw, fallback) {
        if (window.ms365ElternGuardians && typeof window.ms365ElternGuardians.normalizeNamePattern === 'function') {
            return window.ms365ElternGuardians.normalizeNamePattern(raw, fallback);
        }
        const arr = Array.isArray(raw) ? raw : [];
        const out = [];
        arr.forEach(function (p) {
            if (!p || typeof p !== 'object') return;
            const type = String(p.type || '').trim();
            if (type === 'text') out.push({ type: 'text', value: String(p.value ?? '') });
            else if (type === 'klasse' || type === 'year') out.push({ type: type });
        });
        return out.length ? out : Array.isArray(fallback) ? fallback.slice() : [];
    }

    function normEmailKey(v) {
        return String(v ?? '')
            .trim()
            .toLowerCase();
    }

    /**
     * Präfix für group.mailNickname (Einrichtungsassistent): gleiche Zeichenregel wie
     * sanitizeUnifiedGroupMailNickname in graph-unified-groups.js (Microsoft Learn:
     * directoryObject validateProperties – ungültig u. a. @ ( ) \ [ ] " ; : < > , und Leerzeichen).
     * Zusätzlich nur ASCII, keine Steuerzeichen. Erlaubt u. a. . - _
     */
    function mailNicknamePrefixSanitize(raw, maxLen) {
        const lim = typeof maxLen === 'number' && maxLen > 0 ? maxLen : 24;
        const s = String(raw ?? '')
            .trim()
            .toLowerCase();
        let out = '';
        for (let i = 0; i < s.length; i++) {
            const c = s.charCodeAt(i);
            if (c < 32 || c === 127 || c > 127) continue;
            const ch = s.charAt(i);
            if (/[@()[\]\\";:<>,\s]/.test(ch)) continue;
            out += ch;
        }
        if (out.length > lim) out = out.slice(0, lim);
        return out;
    }

    function normalizeDirectoryMatchByEmail(raw) {
        const out = {};
        const src = raw && typeof raw === 'object' ? raw : {};
        Object.keys(src).forEach(function (k) {
            const em = normEmailKey(k);
            if (!em || em.indexOf('@') === -1) return;
            const v = src[k];
            if (!v || typeof v !== 'object') return;
            if (v.notFound === true) {
                out[em] = {
                    graphUserId: '',
                    displayName: '',
                    userPrincipalName: '',
                    notFound: true,
                    checkedAt: String(v.checkedAt || '')
                };
                return;
            }
            const id = String(v.graphUserId || v.id || '').trim();
            if (!id) return;
            out[em] = {
                graphUserId: id,
                displayName: String(v.displayName || '').trim(),
                userPrincipalName: String(v.userPrincipalName || '').trim(),
                notFound: false,
                checkedAt: String(v.checkedAt || '')
            };
        });
        return out;
    }

    function normCode(v) {
        return String(v ?? '')
            .trim()
            .toUpperCase();
    }

    function stableLocalId(prefix, key) {
        let h = 5381;
        const s = String(key || '');
        for (let i = 0; i < s.length; i++) h = (h << 5) + h + s.charCodeAt(i);
        return String(prefix || 'x') + Math.abs(h >>> 0).toString(36);
    }

    function emptyYearBucket() {
        return { students: [], studentCouncil: [], classes: [], guardians: [], parentLists: [] };
    }

    function normalizeGuardian(row) {
        const r = row && typeof row === 'object' ? row : {};
        const email = normEmailKey(r.email);
        const name = String(r.name || '').trim();
        const phone = String(r.phone || '').trim();
        const note = String(r.note || '').trim();
        if (!email && !name) return null;
        let id = String(r.id || '').trim();
        if (!id) id = stableLocalId('g_', (email || name).toLowerCase());
        return { id: id, name: name, email: email, phone: phone, note: note };
    }

    function normalizeParentList(row) {
        const r = row && typeof row === 'object' ? row : {};
        const scope = r.scope === 'year' ? 'year' : 'class';
        let code = '';
        if (scope === 'year') {
            code = String(r.code || '').trim();
            if (!/^\d{4}$/.test(code)) return null;
        } else {
            code = normCode(r.code);
            if (!code) return null;
        }
        return {
            scope: scope,
            code: code,
            displayName: String(r.displayName || '').trim(),
            mailNickname: String(r.mailNickname || '').trim(),
            graphGroupId: r.graphGroupId ? String(r.graphGroupId).trim() : '',
            lastExportAt: String(r.lastExportAt || '').trim()
        };
    }

    function normalizeStudentRow(row, usedIds) {
        const r = row && typeof row === 'object' ? row : {};
        const klasse = String(r.klasse || r.class || r.group || '').trim();
        const name = String(r.name || '').trim();
        const email = normEmailKey(r.email);
        if (!klasse && !name && !email) return null;
        let id = String(r.id || '').trim();
        if (!id) id = stableLocalId('s_', [klasse, email || name].join('|').toLowerCase());
        if (usedIds && usedIds.has(id)) {
            let n = 2;
            while (usedIds.has(id + '_' + n)) n++;
            id = id + '_' + n;
        }
        if (usedIds) usedIds.add(id);
        const guardianIds = [];
        const seenG = new Set();
        (Array.isArray(r.guardianIds) ? r.guardianIds : []).forEach(function (gid) {
            const g = String(gid || '').trim();
            if (!g || seenG.has(g)) return;
            seenG.add(g);
            guardianIds.push(g);
        });
        return { id: id, klasse: klasse, name: name, email: email, guardianIds: guardianIds };
    }

    function normalizeStudentCouncilRow(row) {
        const r = row && typeof row === 'object' ? row : {};
        const klasse = String(r.klasse || r.class || r.group || '').trim();
        const name = String(r.name || '').trim();
        const email = normEmailKey(r.email);
        if (!klasse && !name && !email) return null;
        return { klasse: klasse, name: name, email: email };
    }

    function normalizeYearBucket(raw) {
        const base = emptyYearBucket();
        const y = raw && typeof raw === 'object' ? raw : {};
        const usedStudentIds = new Set();
        const students = [];
        (Array.isArray(y.students) ? y.students : []).forEach(function (s) {
            const n = normalizeStudentRow(s, usedStudentIds);
            if (n) students.push(n);
        });

        const studentCouncil = [];
        (Array.isArray(y.studentCouncil) ? y.studentCouncil : []).forEach(function (s) {
            const n = normalizeStudentCouncilRow(s);
            if (n) studentCouncil.push(n);
        });

        const guardians = [];
        const guardianIdSet = new Set();
        const guardianEmailMap = new Map();
        (Array.isArray(y.guardians) ? y.guardians : []).forEach(function (g) {
            const n = normalizeGuardian(g);
            if (!n) return;
            if (guardianIdSet.has(n.id)) return;
            if (n.email && guardianEmailMap.has(n.email)) {
                const prev = guardianEmailMap.get(n.email);
                if (!prev.name && n.name) prev.name = n.name;
                if (!prev.phone && n.phone) prev.phone = n.phone;
                if (!prev.note && n.note) prev.note = n.note;
                return;
            }
            guardianIdSet.add(n.id);
            if (n.email) guardianEmailMap.set(n.email, n);
            guardians.push(n);
        });

        students.forEach(function (s) {
            s.guardianIds = (s.guardianIds || []).filter(function (gid) {
                return guardianIdSet.has(gid);
            });
        });

        const parentLists = [];
        const plKeys = new Set();
        (Array.isArray(y.parentLists) ? y.parentLists : []).forEach(function (p) {
            const n = normalizeParentList(p);
            if (!n) return;
            const key = n.scope + ':' + n.code;
            if (plKeys.has(key)) return;
            plKeys.add(key);
            parentLists.push(n);
        });

        base.students = students;
        base.studentCouncil = studentCouncil;
        base.classes = Array.isArray(y.classes) ? deepClone(y.classes) : [];
        base.guardians = guardians;
        base.parentLists = parentLists;
        return base;
    }

    /**
     * Schülerimport inkl. optionaler parentPairs; bestehende IDs/Zuordnungen erhalten.
     */
    function mergeStudentsImport(prevBucket, incomingStudents) {
        const prev = normalizeYearBucket(prevBucket);
        const oldByEmail = new Map();
        const oldByKey = new Map();
        prev.students.forEach(function (s) {
            if (s.email) oldByEmail.set(s.email, s);
            oldByKey.set([String(s.klasse || '').toLowerCase(), String(s.name || '').toLowerCase()].join('|'), s);
        });

        const guardians = prev.guardians.slice();
        const byEmail = new Map();
        guardians.forEach(function (g) {
            if (g.email) byEmail.set(g.email, g);
        });

        function upsertGuardian(pair) {
            const email = normEmailKey(pair && pair.email);
            if (!email || email.indexOf('@') === -1) return '';
            if (byEmail.has(email)) {
                const g = byEmail.get(email);
                const nm = String((pair && pair.name) || '').trim();
                if (nm && !g.name) g.name = nm;
                return g.id;
            }
            const g = normalizeGuardian({
                name: String((pair && pair.name) || '').trim(),
                email: email
            });
            if (!g) return '';
            guardians.push(g);
            byEmail.set(email, g);
            return g.id;
        }

        const usedIds = new Set();
        const students = [];
        (Array.isArray(incomingStudents) ? incomingStudents : []).forEach(function (raw) {
            const klasse = String(raw?.klasse || raw?.class || '').trim();
            const name = String(raw?.name || '').trim();
            const email = normEmailKey(raw?.email);
            if (!klasse && !name && !email) return;
            let prevS = email && oldByEmail.has(email) ? oldByEmail.get(email) : null;
            if (!prevS) {
                prevS = oldByKey.get([klasse.toLowerCase(), name.toLowerCase()].join('|')) || null;
            }
            const row = normalizeStudentRow(
                {
                    id: prevS ? prevS.id : raw?.id,
                    klasse: klasse,
                    name: name,
                    email: email,
                    guardianIds: prevS ? prevS.guardianIds : raw?.guardianIds
                },
                usedIds
            );
            if (!row) return;
            const pairs = Array.isArray(raw?.parentPairs)
                ? raw.parentPairs
                : Array.isArray(raw?.parents)
                  ? raw.parents
                  : [];
            if (pairs.length) {
                const ids = [];
                const seen = new Set();
                pairs.forEach(function (p) {
                    const gid = upsertGuardian(p);
                    if (gid && !seen.has(gid)) {
                        seen.add(gid);
                        ids.push(gid);
                    }
                });
                if (ids.length) row.guardianIds = ids;
            }
            students.push(row);
        });

        return normalizeYearBucket({
            students: students,
            classes: prev.classes,
            guardians: guardians,
            parentLists: prev.parentLists
        });
    }

    function normalizeSammelgruppeCode(v) {
        const c = String(v ?? '')
            .trim()
            .toLowerCase();
        if (c === 'schueler' || c === 'lehrer' || c === 'verwaltung') return c;
        return '';
    }

    function sammelgruppeFieldForCode(code) {
        if (code === 'schueler') return 'schuelerGroupId';
        if (code === 'lehrer') return 'lehrerGroupId';
        if (code === 'verwaltung') return 'verwaltungGroupId';
        return '';
    }

    function writeSammelgruppeCatalogLink(links, code, graphGroupId) {
        const arr = Array.isArray(links) ? links.slice() : [];
        const c = normalizeSammelgruppeCode(code);
        if (!c) return arr;
        const gid = graphGroupId ? String(graphGroupId).trim() : '';
        const idx = arr.findIndex(function (x) {
            return x && x.kind === 'sammelgruppe' && x.code === c;
        });
        if (idx >= 0) {
            arr[idx] = Object.assign({}, arr[idx], {
                graphGroupId: gid,
                mode: gid ? arr[idx].mode || 'matched' : ''
            });
            return arr;
        }
        if (!gid) return arr;
        arr.push({
            kind: 'sammelgruppe',
            code: c,
            graphGroupId: gid,
            displayName: '',
            mailNickname: '',
            mode: 'matched',
            syncStatus: ''
        });
        return arr;
    }

    /**
     * Eine Quelle: catalogLinks (inkl. Sammelgruppen). `matched.*GroupId` bleibt
     * Spiegel für Wizard/SLG/Verwaltung. Lücken: die nicht-leere Seite gewinnt.
     */
    function fillSammelgruppeGaps(matched, catalogLinks) {
        const m = matched && typeof matched === 'object' ? Object.assign({}, matched) : {};
        let links = Array.isArray(catalogLinks) ? catalogLinks.slice() : [];
        ['schueler', 'lehrer', 'verwaltung'].forEach(function (code) {
            const field = sammelgruppeFieldForCode(code);
            const link = links.find(function (x) {
                return x && x.kind === 'sammelgruppe' && x.code === code;
            });
            const catId = link && link.graphGroupId ? String(link.graphGroupId).trim() : '';
            const matId = m[field] ? String(m[field]).trim() : '';
            const id = matId || catId;
            m[field] = id || null;
            links = writeSammelgruppeCatalogLink(links, code, id);
        });
        return { matched: m, catalogLinks: links };
    }

    function normalizeCatalogLink(row) {
        const r = row && typeof row === 'object' ? row : {};
        if (r.kind === 'sammelgruppe') {
            const code = normalizeSammelgruppeCode(r.code);
            if (!code) return null;
            const mode = r.mode === 'created' || r.mode === 'matched' ? r.mode : '';
            return {
                kind: 'sammelgruppe',
                code: code,
                graphGroupId: r.graphGroupId ? String(r.graphGroupId).trim() : '',
                displayName: String(r.displayName || '').trim(),
                mailNickname: String(r.mailNickname || '').trim(),
                mode: mode,
                syncStatus: String(r.syncStatus || '').trim()
            };
        }
        if (r.kind === 'cohort' || r.kind === 'eltern') {
            const code = String(r.code || '').trim();
            if (!/^\d{4}$/.test(code)) return null;
            const mode = r.mode === 'created' || r.mode === 'matched' ? r.mode : '';
            return {
                kind: r.kind,
                code: code,
                graphGroupId: r.graphGroupId ? String(r.graphGroupId).trim() : '',
                displayName: String(r.displayName || '').trim(),
                mailNickname: String(r.mailNickname || '').trim(),
                mode: mode,
                syncStatus: String(r.syncStatus || '').trim()
            };
        }
        const kind = r.kind === 'arge' ? 'arge' : 'subject';
        const code = normCode(r.code);
        if (!code) return null;
        const mode = r.mode === 'created' || r.mode === 'matched' ? r.mode : '';
        return {
            kind: kind,
            code: code,
            graphGroupId: r.graphGroupId ? String(r.graphGroupId).trim() : '',
            displayName: String(r.displayName || '').trim(),
            mailNickname: String(r.mailNickname || '').trim(),
            mode: mode,
            syncStatus: String(r.syncStatus || '').trim()
        };
    }

    function normalizeSetup(s) {
        const d = defaultSetup();
        const x = s && typeof s === 'object' ? s : {};
        let ws = parseInt(x.wizardStep, 10);
        const layout11 = x._einrichtungWizardLayout === 11;
        const layout10 = x._einrichtungWizardLayout === 10;
        const layout9 = x._einrichtungWizardLayout === 9;
        const layout8 = x._einrichtungWizardLayout === 8;
        const layout7 = x._einrichtungWizardLayout === 7;
        const layout6 = x._einrichtungWizardLayout === 6;
        // Nur echte Layout-Upgrades; bei bereits aktuellem Layout keine erneute Verschiebung
        if (!layout8 && !layout9 && !layout10 && !layout11 && !isNaN(ws)) {
            // Layout 7 → 8: neuer Schritt „Verwaltung“ vor Lehrkräften (alte 3–7 werden 4–8)
            if (layout7 && ws >= 3 && ws <= 7) ws += 1;
            // Sehr alt: 5 Schritte (4=Katalog, 5=Klassen) → +1 für eingefügte Personen-Schritte
            if (!layout7 && !layout6 && ws >= 4 && ws <= 5) ws += 1;
            // Vorher 6 Schritte: Klassen war Schritt 6 → jetzt Schritt 7
            if (layout6 && ws === 6) ws = 7;
        }
        if (!isNaN(ws)) {
            if ((layout9 || layout10) && ws >= 6 && ws <= 9) ws += 2;
        }
        d.wizardStep = !isNaN(ws) && ws >= 1 && ws <= 11 ? ws : 1;
        d._einrichtungWizardLayout = 11;
        d.completedSteps = Array.isArray(x.completedSteps) ? x.completedSteps.map((t) => String(t)) : [];
        d.finishedAt = x.finishedAt != null && x.finishedAt !== '' ? String(x.finishedAt) : null;
        d.lastVisitedAt = x.lastVisitedAt != null && x.lastVisitedAt !== '' ? String(x.lastVisitedAt) : null;
        const m = x.matched && typeof x.matched === 'object' ? x.matched : {};
        d.matched = {
            schuelerGroupId: m.schuelerGroupId ? String(m.schuelerGroupId).trim() : null,
            lehrerGroupId: m.lehrerGroupId ? String(m.lehrerGroupId).trim() : null,
            verwaltungGroupId: m.verwaltungGroupId ? String(m.verwaltungGroupId).trim() : null,
            sgaGroupId: m.sgaGroupId ? String(m.sgaGroupId).trim() : null,
            studentCouncilGroupId: m.studentCouncilGroupId ? String(m.studentCouncilGroupId).trim() : null
        };
        const dr = x.slgDraft && typeof x.slgDraft === 'object' ? x.slgDraft : {};
        const srcLehrer = String(dr.slgOwnerSourceLehrer || '').trim();
        const srcSchueler = String(dr.slgOwnerSourceSchueler || '').trim();
        d.slgDraft = {
            activeKind: dr.activeKind === 'lehrer' ? 'lehrer' : 'schueler',
            slgNewDisplayName: String(dr.slgNewDisplayName != null ? dr.slgNewDisplayName : ''),
            slgNewMailNick: String(dr.slgNewMailNick != null ? dr.slgNewMailNick : ''),
            slgNewDescription: String(dr.slgNewDescription != null ? dr.slgNewDescription : ''),
            slgNewCreateTeam: !!dr.slgNewCreateTeam,
            slgOwnerSourceLehrer:
                srcLehrer === 'teachers' || srcLehrer === 'manual' ? srcLehrer : 'direktion',
            slgOwnerManualEmailsLehrer: String(dr.slgOwnerManualEmailsLehrer != null ? dr.slgOwnerManualEmailsLehrer : ''),
            slgOwnerSourceSchueler:
                srcSchueler === 'admin' || srcSchueler === 'manual' ? srcSchueler : 'direktion',
            slgOwnerManualEmailsSchueler: String(
                dr.slgOwnerManualEmailsSchueler != null ? dr.slgOwnerManualEmailsSchueler : ''
            )
        };
        const vd = x.verwaltungDraft && typeof x.verwaltungDraft === 'object' ? x.verwaltungDraft : {};
        const vwSrc = String(vd.vwOwnerSource || '').trim();
        d.verwaltungDraft = {
            vwNewDisplayName: String(vd.vwNewDisplayName != null ? vd.vwNewDisplayName : 'Schulverwaltung'),
            vwNewMailNick: mailNicknamePrefixSanitize(vd.vwNewMailNick || 'verwaltung', 60) || 'verwaltung',
            vwNewDescription: String(vd.vwNewDescription != null ? vd.vwNewDescription : ''),
            vwNewCreateTeam: !!vd.vwNewCreateTeam,
            vwOwnerSource: vwSrc === 'direktion' || vwSrc === 'manual' ? vwSrc : 'admin',
            vwOwnerManualEmails: String(vd.vwOwnerManualEmails != null ? vd.vwOwnerManualEmails : '')
        };
        d.subjectGroupMailPrefix = mailNicknamePrefixSanitize(x.subjectGroupMailPrefix, 24) || 'fach';
        d.argeGroupMailPrefix = mailNicknamePrefixSanitize(x.argeGroupMailPrefix, 24) || 'ag';
        const def = defaultSetup();
        d.elternClassAliasPattern = normalizeElternNamePattern(x.elternClassAliasPattern, def.elternClassAliasPattern);
        d.elternClassDisplayPattern = normalizeElternNamePattern(x.elternClassDisplayPattern, def.elternClassDisplayPattern);
        d.elternYearAliasPattern = normalizeElternNamePattern(x.elternYearAliasPattern, def.elternYearAliasPattern);
        d.elternYearDisplayPattern = normalizeElternNamePattern(x.elternYearDisplayPattern, def.elternYearDisplayPattern);
        const linksIn = Array.isArray(x.catalogLinks) ? x.catalogLinks : [];
        const seen = new Set();
        d.catalogLinks = [];
        linksIn.forEach(function (row) {
            const n = normalizeCatalogLink(row);
            if (!n) return;
            const k = n.kind + ':' + n.code.toLowerCase();
            if (seen.has(k)) return;
            seen.add(k);
            d.catalogLinks.push(n);
        });
        d.directoryMatchByEmail = normalizeDirectoryMatchByEmail(x.directoryMatchByEmail);
        const cgmRaw = x.classGroupMatchByKey && typeof x.classGroupMatchByKey === 'object' ? x.classGroupMatchByKey : {};
        d.classGroupMatchByKey = {};
        Object.keys(cgmRaw).forEach(function (k) {
            const key = String(k).trim().toUpperCase();
            if (!key) return;
            const v = cgmRaw[k];
            if (!v || typeof v !== 'object') return;
            d.classGroupMatchByKey[key] = v;
        });
        const filled = fillSammelgruppeGaps(d.matched, d.catalogLinks);
        d.matched = filled.matched;
        d.catalogLinks = filled.catalogLinks;
        d.intranetSiteUrl = x.intranetSiteUrl ? String(x.intranetSiteUrl).trim() : '';
        d.intranetHubAt = x.intranetHubAt != null && x.intranetHubAt !== '' ? String(x.intranetHubAt) : null;
        const es = x.elternSetup && typeof x.elternSetup === 'object' ? x.elternSetup : {};
        d.elternSetup = {
            completedSteps: Array.isArray(es.completedSteps) ? es.completedSteps.map(function (t) { return String(t); }) : [],
            lastDiagnoseAt: es.lastDiagnoseAt ? String(es.lastDiagnoseAt) : null
        };
        d.sisImportHistory = [];
        (Array.isArray(x.sisImportHistory) ? x.sisImportHistory : []).slice(-20).forEach(function (row) {
            if (!row || typeof row !== 'object') return;
            d.sisImportHistory.push({
                at: row.at ? String(row.at) : '',
                source: row.source ? String(row.source) : '',
                mode: row.mode === 'replace' ? 'replace' : 'merge',
                added: Number(row.added) || 0,
                updated: Number(row.updated) || 0,
                removed: Number(row.removed) || 0,
                conflicts: Number(row.conflicts) || 0
            });
        });
        d.actionLog = [];
        (Array.isArray(x.actionLog) ? x.actionLog : []).slice(-200).forEach(function (row) {
            if (!row || typeof row !== 'object') return;
            d.actionLog.push({
                at: row.at ? String(row.at) : '',
                tool: row.tool ? String(row.tool) : 'app',
                action: row.action ? String(row.action) : 'write',
                target: row.target ? String(row.target) : '',
                summary: row.summary ? String(row.summary) : '',
                result: row.result === 'error' || row.result === 'skip' ? row.result : 'ok'
            });
        });
        return d;
    }

    function deriveStableNickFromClassRow(cl) {
        if (typeof window.ms365DeriveClassStableMailNickname === 'function') {
            return String(window.ms365DeriveClassStableMailNickname(cl.year || '', cl.code || '') || '')
                .trim()
                .replace(/[^a-zA-Z0-9]/g, '')
                .toLowerCase()
                .slice(0, 60);
        }
        const y = String(cl.year || '').trim();
        const yy = /^\d{4}$/.test(y) ? y : '';
        const code = normCode(cl.code || '');
        const tail = String(code)
            .replace(/[^0-9A-Za-z]/g, '')
            .toLowerCase()
            .slice(0, 24);
        if (!yy || !tail) return '';
        return ('jg' + yy + tail).toLowerCase().slice(0, 60);
    }

    function classTeamIdentityNick(raw) {
        return String(raw || '')
            .trim()
            .replace(/[^a-zA-Z0-9]/g, '')
            .toLowerCase()
            .slice(0, 60);
    }

    function normalizeClassTeam(row) {
        const r = row && typeof row === 'object' ? row : {};
        let nick = classTeamIdentityNick(r.stableMailNickname);
        if (!nick) return null;
        const mode = r.mode === 'created' || r.mode === 'matched' ? r.mode : '';
        const y = String(r.abschlussJahr || r.year || '').trim();
        const abschlussJahr = /^\d{4}$/.test(y) ? y : '';
        const mailNickname = mailNicknamePrefixSanitize(r.mailNickname || r.graphMailNickname || '', 60);
        return {
            stableMailNickname: nick,
            mailNickname: mailNickname,
            graphGroupId: String(r.graphGroupId || '').trim(),
            classCode: normCode(r.classCode || r.code || ''),
            displayName: String(r.displayName || r.name || '').trim(),
            abschlussJahr: abschlussJahr,
            mode: mode,
            educationClassId: String(r.educationClassId || '').trim()
        };
    }

    function normalizeCoreClassTeams(arr) {
        const seen = new Set();
        const out = [];
        (Array.isArray(arr) ? arr : []).forEach(function (row) {
            const n = normalizeClassTeam(row);
            if (!n) return;
            if (seen.has(n.stableMailNickname)) return;
            seen.add(n.stableMailNickname);
            out.push(n);
        });
        return out;
    }

    function classTeamMatchesKlasse(ct, klasseRaw) {
        const k = String(klasseRaw ?? '').trim();
        if (!k) return false;
        const cc = normCode(ct.classCode || '');
        const dn = String(ct.displayName || '').trim();
        if (dn && k === dn) return true;
        const nk = normCode(k);
        if (cc && nk === cc) return true;
        if (cc && cc.length >= 2 && k.toUpperCase().indexOf(cc) !== -1) return true;
        return false;
    }

    function reconcileClassTeamsFromYearClasses(c, _yearKey, classesArr) {
        const classes = Array.isArray(classesArr) ? classesArr : [];
        let teams = normalizeCoreClassTeams(c.core.classTeams || []);
        const byNick = {};
        teams.forEach(function (t) {
            byNick[t.stableMailNickname] = t;
        });
        classes.forEach(function (cl) {
            let nick = String(cl.stableMailNickname || '')
                .trim()
                .replace(/[^a-zA-Z0-9]/g, '')
                .toLowerCase()
                .slice(0, 60);
            if (!nick && cl.year && cl.code) nick = deriveStableNickFromClassRow(cl);
            if (!nick) return;
            let ex = byNick[nick];
            if (!ex) {
                ex = normalizeClassTeam({
                    stableMailNickname: nick,
                    classCode: cl.code,
                    displayName: cl.name,
                    abschlussJahr: cl.year,
                    graphGroupId: '',
                    mode: ''
                });
                if (!ex) return;
                teams.push(ex);
                byNick[nick] = ex;
            } else {
                if (cl.name) ex.displayName = String(cl.name).trim();
                if (cl.code) ex.classCode = normCode(cl.code);
                const yr = String(cl.year || '').trim();
                if (/^\d{4}$/.test(yr)) ex.abschlussJahr = yr;
            }
        });
        c.core.classTeams = normalizeCoreClassTeams(teams);
    }

    function emptyContainer() {
        return {
            version: VERSION,
            core: {
                schoolName: '',
                domain: '',
                subjects: [],
                arges: [],
                teachers: [],
                administration: [],
                admin: [],
                adminRoles: [],
                sgaMode: 'group',
                sga: [],
                classTeams: []
            },
            years: {
                current: currentSchoolYearLabel(),
                byLabel: {}
            },
            structure: {
                rows: [],
                memberships: {},
                settings: {}
            },
            tenant: {
                cache: {
                    rows: [],
                    users: [],
                    loadedAt: ''
                }
            },
            match: {
                links: {}
            },
            setup: defaultSetup()
        };
    }

    function normalizeContainer(obj) {
        const base = emptyContainer();
        const o = obj && typeof obj === 'object' ? obj : {};
        const out = Object.assign({}, base, o);
        out.version = VERSION;

        out.core = Object.assign({}, base.core, (o.core && typeof o.core === 'object' ? o.core : {}));
        out.core.classTeams = normalizeCoreClassTeams(out.core.classTeams);
        out.years = Object.assign({}, base.years, (o.years && typeof o.years === 'object' ? o.years : {}));
        out.years.byLabel = Object.assign({}, base.years.byLabel, (out.years.byLabel && typeof out.years.byLabel === 'object' ? out.years.byLabel : {}));

        out.structure = Object.assign({}, base.structure, (o.structure && typeof o.structure === 'object' ? o.structure : {}));
        out.tenant = Object.assign({}, base.tenant, (o.tenant && typeof o.tenant === 'object' ? o.tenant : {}));
        out.tenant.cache = Object.assign({}, base.tenant.cache, (out.tenant.cache && typeof out.tenant.cache === 'object' ? out.tenant.cache : {}));
        out.match = Object.assign({}, base.match, (o.match && typeof o.match === 'object' ? o.match : {}));
        out.match.links = Object.assign({}, base.match.links, (out.match.links && typeof out.match.links === 'object' ? out.match.links : {}));

        out.setup = normalizeSetup(o.setup);

        if (!out.years.current) out.years.current = currentSchoolYearLabel();
        const labels = Object.keys(out.years.byLabel || {});
        labels.forEach(function (lab) {
            out.years.byLabel[lab] = normalizeYearBucket(out.years.byLabel[lab]);
        });
        if (!out.years.byLabel[out.years.current]) out.years.byLabel[out.years.current] = emptyYearBucket();

        return out;
    }

    function loadV2Raw() {
        try {
            const raw = localStorage.getItem(STORAGE_KEY_V2);
            if (!raw) return null;
            return safeJsonParse(raw);
        } catch {
            return null;
        }
    }

    function saveV2(container) {
        const normalized = normalizeContainer(container);
        try {
            localStorage.setItem(STORAGE_KEY_V2, JSON.stringify(normalized));
        } catch {
            // ignore
        }
        return normalized;
    }

    function migrateFromV1IfNeeded() {
        const existing = loadV2Raw();
        if (existing && typeof existing === 'object') {
            return saveV2(normalizeContainer(existing));
        }

        // Migrate from legacy keys (best-effort, non-destructive)
        const out = emptyContainer();

        // tenant-settings-core v1
        try {
            const rawCore = localStorage.getItem('ms365-tenant-settings-v1');
            const coreObj = rawCore ? safeJsonParse(rawCore) : null;
            if (coreObj && typeof coreObj === 'object') {
                out.core.schoolName = String(coreObj.schoolName || coreObj.name || '').trim();
                out.core.domain = String(coreObj.domain || '').trim();
                out.core.subjects = Array.isArray(coreObj.subjects) ? deepClone(coreObj.subjects) : [];
                out.core.arges = Array.isArray(coreObj.arges) ? deepClone(coreObj.arges) : [];
                out.core.teachers = Array.isArray(coreObj.teachers) ? deepClone(coreObj.teachers) : [];
                out.core.administration = Array.isArray(coreObj.administration) ? deepClone(coreObj.administration) : [];
                out.core.admin = Array.isArray(coreObj.admin) ? deepClone(coreObj.admin) : [];
                out.core.adminRoles = Array.isArray(coreObj.adminRoles) ? deepClone(coreObj.adminRoles) : [];

                const cur = out.years.current;
                out.years.byLabel[cur] = normalizeYearBucket({
                    students: Array.isArray(coreObj.students) ? deepClone(coreObj.students) : [],
                    classes: Array.isArray(coreObj.classes) ? deepClone(coreObj.classes) : []
                });
            }
        } catch {
            // ignore
        }

        // schulstruktur-sync v1
        try {
            const rawStruct = localStorage.getItem('ms365-schulstruktur-sync-v1');
            const st = rawStruct ? safeJsonParse(rawStruct) : null;
            if (st && typeof st === 'object') {
                out.structure.rows = Array.isArray(st.rows) ? deepClone(st.rows) : [];
                out.structure.memberships = st.memberships && typeof st.memberships === 'object' ? deepClone(st.memberships) : {};
                out.structure.settings = st.settings && typeof st.settings === 'object' ? deepClone(st.settings) : {};
            }
        } catch {
            // ignore
        }

        try {
            const rawMatch = localStorage.getItem('ms365-schulstruktur-match-v1');
            const m = rawMatch ? safeJsonParse(rawMatch) : null;
            if (m && typeof m === 'object' && m.links && typeof m.links === 'object') {
                out.match.links = deepClone(m.links);
            }
        } catch {
            // ignore
        }

        try {
            const rawCache = localStorage.getItem('ms365-schulstruktur-tenant-cache-v1');
            const c = rawCache ? safeJsonParse(rawCache) : null;
            if (c && typeof c === 'object') {
                out.tenant.cache.rows = Array.isArray(c.rows) ? deepClone(c.rows) : [];
                out.tenant.cache.users = Array.isArray(c.users) ? deepClone(c.users) : [];
                out.tenant.cache.loadedAt = String(c.loadedAt || '');
            }
        } catch {
            // ignore
        }

        return saveV2(out);
    }

    function maybeMergeSlgLocalIntoSetup(c) {
        try {
            const m = c.setup && c.setup.matched;
            if (m && (m.schuelerGroupId || m.lehrerGroupId)) return c;
            const raw = localStorage.getItem(SLG_LEGACY_KEY);
            if (!raw) return c;
            const o = safeJsonParse(raw);
            if (!o || typeof o !== 'object' || !o.matched || typeof o.matched !== 'object') return c;
            const sm = o.matched.schuelerGroupId ? String(o.matched.schuelerGroupId).trim() : '';
            const lm = o.matched.lehrerGroupId ? String(o.matched.lehrerGroupId).trim() : '';
            if (!sm && !lm) return c;
            c.setup = normalizeSetup(c.setup);
            if (sm) c.setup.matched.schuelerGroupId = sm;
            if (lm) c.setup.matched.lehrerGroupId = lm;
            if (o.activeKind === 'lehrer' || o.activeKind === 'schueler') {
                c.setup.slgDraft.activeKind = o.activeKind;
            }
            if (o.slgNewDisplayName !== undefined) c.setup.slgDraft.slgNewDisplayName = String(o.slgNewDisplayName);
            if (o.slgNewMailNick !== undefined) c.setup.slgDraft.slgNewMailNick = String(o.slgNewMailNick);
            if (o.slgNewDescription !== undefined) c.setup.slgDraft.slgNewDescription = String(o.slgNewDescription);
            if (o.slgNewCreateTeam !== undefined) c.setup.slgDraft.slgNewCreateTeam = !!o.slgNewCreateTeam;
            return saveV2(c);
        } catch {
            return c;
        }
    }

    function getContainer() {
        const c = migrateFromV1IfNeeded();
        return maybeMergeSlgLocalIntoSetup(c);
    }

    function setContainer(next) {
        return saveV2(next);
    }

    function setCoreFromTenantSettings(v1Settings) {
        const c = getContainer();
        const keepClassTeams = normalizeCoreClassTeams(c.core.classTeams || []);
        const s = v1Settings && typeof v1Settings === 'object' ? v1Settings : {};
        c.core.schoolName = String(s.schoolName || '').trim();
        c.core.domain = String(s.domain || '').trim();
        c.core.subjects = Array.isArray(s.subjects) ? deepClone(s.subjects) : [];
        c.core.arges = Array.isArray(s.arges) ? deepClone(s.arges) : [];
        c.core.teachers = Array.isArray(s.teachers) ? deepClone(s.teachers) : [];
        c.core.administration = Array.isArray(s.administration) ? deepClone(s.administration) : [];
        c.core.admin = Array.isArray(s.admin) ? deepClone(s.admin) : [];
        c.core.adminRoles = Array.isArray(s.adminRoles) ? deepClone(s.adminRoles) : [];
        c.core.sgaMode = String(s.sgaMode || '').trim() === 'distribution' ? 'distribution' : 'group';
        c.core.sga = Array.isArray(s.sga) ? deepClone(s.sga) : [];
        c.core.classTeams = keepClassTeams;
        const cur = String(c.years.current || currentSchoolYearLabel());
        const prev = c.years.byLabel[cur] || emptyYearBucket();
        const merged = mergeStudentsImport(prev, Array.isArray(s.students) ? s.students : []);
        merged.studentCouncil = Array.isArray(s.studentCouncil) ? deepClone(s.studentCouncil) : [];
        merged.classes = Array.isArray(s.classes) ? deepClone(s.classes) : [];
        c.years.byLabel[cur] = normalizeYearBucket(merged);
        reconcileClassTeamsFromYearClasses(c, cur, c.years.byLabel[cur].classes);
        return saveV2(c);
    }

    function listYears() {
        const c = getContainer();
        const by = c && c.years && c.years.byLabel && typeof c.years.byLabel === 'object' ? c.years.byLabel : {};
        return Object.keys(by).map((k) => String(k)).sort();
    }

    function setCurrentYear(label, opts) {
        const y = String(label || '').trim();
        if (!y) throw new Error('Schuljahr fehlt.');
        const o = opts && typeof opts === 'object' ? opts : {};
        const c = getContainer();
        const by = c.years.byLabel || {};
        if (!by[y]) {
            let seed = emptyYearBucket();
            const copyFrom = String(o.copyFrom || '').trim();
            if (copyFrom && by[copyFrom]) {
                // Kopie von Schüler/Klassen/Eltern; alles andere ist global.
                seed = normalizeYearBucket(deepClone(by[copyFrom]));
                seed.parentLists = (seed.parentLists || []).map(function (p) {
                    return Object.assign({}, p, { graphGroupId: '', lastExportAt: '' });
                });
            }
            by[y] = seed;
        }
        c.years.byLabel = by;
        c.years.current = y;
        return saveV2(c);
    }

    function getSetup() {
        const c = getContainer();
        return normalizeSetup(c.setup);
    }

    function patchSetup(partial) {
        const c = getContainer();
        const cur = normalizeSetup(c.setup);
        const p = partial && typeof partial === 'object' ? partial : {};
        const pCopy = Object.assign({}, p);
        delete pCopy.directoryMatchByEmail;
        delete pCopy.classGroupMatchByKey;
        const mergedDir = Object.assign(
            {},
            cur.directoryMatchByEmail || {},
            p.directoryMatchByEmail && typeof p.directoryMatchByEmail === 'object' ? p.directoryMatchByEmail : {}
        );
        if (p.directoryMatchByEmailRemove && typeof p.directoryMatchByEmailRemove === 'object') {
            Object.keys(p.directoryMatchByEmailRemove).forEach(function (k) {
                const em = normEmailKey(k);
                if (em) delete mergedDir[em];
            });
        }
        const mergedCgm = Object.assign(
            {},
            cur.classGroupMatchByKey || {},
            p.classGroupMatchByKey && typeof p.classGroupMatchByKey === 'object' ? p.classGroupMatchByKey : {}
        );
        let next = normalizeSetup(
            Object.assign({}, cur, pCopy, {
                matched: Object.assign({}, cur.matched, p.matched && typeof p.matched === 'object' ? p.matched : {}),
                slgDraft: Object.assign({}, cur.slgDraft, p.slgDraft && typeof p.slgDraft === 'object' ? p.slgDraft : {}),
                verwaltungDraft: Object.assign(
                    {},
                    cur.verwaltungDraft,
                    p.verwaltungDraft && typeof p.verwaltungDraft === 'object' ? p.verwaltungDraft : {}
                ),
                catalogLinks: Array.isArray(p.catalogLinks) ? p.catalogLinks : cur.catalogLinks,
                directoryMatchByEmail: mergedDir,
                classGroupMatchByKey: mergedCgm
            })
        );
        const matchedPatched = p.matched && typeof p.matched === 'object';
        const catalogPatched = Array.isArray(p.catalogLinks);
        if (matchedPatched) {
            ['schuelerGroupId', 'lehrerGroupId', 'verwaltungGroupId'].forEach(function (field) {
                if (!Object.prototype.hasOwnProperty.call(p.matched, field)) return;
                const id = p.matched[field] ? String(p.matched[field]).trim() : '';
                next.matched[field] = id || null;
                const code =
                    field === 'schuelerGroupId' ? 'schueler' : field === 'lehrerGroupId' ? 'lehrer' : 'verwaltung';
                next.catalogLinks = writeSammelgruppeCatalogLink(next.catalogLinks, code, id);
            });
        }
        if (catalogPatched) {
            ['schueler', 'lehrer', 'verwaltung'].forEach(function (code) {
                const link = next.catalogLinks.find(function (x) {
                    return x && x.kind === 'sammelgruppe' && x.code === code;
                });
                if (!link) return;
                const field = sammelgruppeFieldForCode(code);
                const id = link.graphGroupId ? String(link.graphGroupId).trim() : '';
                next.matched[field] = id || null;
            });
        }
        const filled = fillSammelgruppeGaps(next.matched, next.catalogLinks);
        next.matched = filled.matched;
        next.catalogLinks = filled.catalogLinks;
        c.setup = next;
        return saveV2(c);
    }

    function getClassTeamGruppenmailForKlasse(klasseRaw) {
        const c = getContainer();
        const teams = normalizeCoreClassTeams(c.core.classTeams || []);
        for (let i = 0; i < teams.length; i++) {
            if (classTeamMatchesKlasse(teams[i], klasseRaw)) {
                return teams[i].stableMailNickname || '';
            }
        }
        return '';
    }

    function upsertClassTeam(entry) {
        const c = getContainer();
        const n = normalizeClassTeam(entry);
        if (!n) throw new Error('Klassen-Team: stableMailNickname fehlt oder ungültig.');
        let teams = normalizeCoreClassTeams(c.core.classTeams || []);
        let idx = teams.findIndex(function (t) {
            return t.stableMailNickname === n.stableMailNickname;
        });
        if (idx < 0 && n.graphGroupId) {
            idx = teams.findIndex(function (t) {
                return t.graphGroupId && t.graphGroupId === n.graphGroupId;
            });
        }
        if (idx < 0 && n.classCode) {
            idx = teams.findIndex(function (t) {
                if (normCode(t.classCode) !== n.classCode) return false;
                if (n.abschlussJahr && t.abschlussJahr && t.abschlussJahr !== n.abschlussJahr) return false;
                return true;
            });
        }
        if (idx >= 0) {
            const prev = teams[idx];
            const merged = Object.assign({}, prev, n);
            if (!n.graphGroupId) {
                merged.mailNickname = n.mailNickname || '';
                merged.graphGroupId = '';
            } else if (!merged.mailNickname && prev.mailNickname) {
                merged.mailNickname = prev.mailNickname;
            }
            teams[idx] = merged;
        } else teams.push(n);
        c.core.classTeams = normalizeCoreClassTeams(teams);
        return saveV2(c);
    }

    function touchWizardVisit(step) {
        const c = getContainer();
        const s = normalizeSetup(c.setup);
        if (typeof step === 'number' && step >= 1 && step <= 9) s.wizardStep = step;
        try {
            s.lastVisitedAt = new Date().toISOString();
        } catch {
            s.lastVisitedAt = '';
        }
        c.setup = s;
        return saveV2(c);
    }

    function exportJson() {
        return getContainer();
    }

    function importJson(obj) {
        // Accept either v2/v3 container or legacy tenant-settings v1 JSON
        const o = obj && typeof obj === 'object' ? obj : null;
        if (!o) throw new Error('Keine gültige JSON.');
        if (o.version >= 2 && o.core && o.structure && o.match) {
            return saveV2(normalizeContainer(o));
        }
        // Legacy: treat as tenant-settings-core v1 payload, update only core+current year
        const cur = getContainer();
        cur.core.schoolName = String(o.schoolName || '').trim();
        cur.core.domain = String(o.domain || '').trim();
        cur.core.subjects = Array.isArray(o.subjects) ? deepClone(o.subjects) : [];
        cur.core.arges = Array.isArray(o.arges) ? deepClone(o.arges) : [];
        cur.core.teachers = Array.isArray(o.teachers) ? deepClone(o.teachers) : [];
        cur.core.administration = Array.isArray(o.administration) ? deepClone(o.administration) : [];
        cur.core.admin = Array.isArray(o.admin) ? deepClone(o.admin) : [];
        cur.core.adminRoles = Array.isArray(o.adminRoles) ? deepClone(o.adminRoles) : [];
        cur.core.sgaMode = String(o.sgaMode || '').trim() === 'distribution' ? 'distribution' : 'group';
        cur.core.sga = Array.isArray(o.sga) ? deepClone(o.sga) : [];
        const y = String(cur.years.current || currentSchoolYearLabel());
        const prev = cur.years.byLabel[y] || emptyYearBucket();
        const merged = mergeStudentsImport(prev, Array.isArray(o.students) ? o.students : []);
        merged.studentCouncil = Array.isArray(o.studentCouncil) ? deepClone(o.studentCouncil) : [];
        merged.classes = Array.isArray(o.classes) ? deepClone(o.classes) : [];
        cur.years.byLabel[y] = normalizeYearBucket(merged);
        if (!Array.isArray(cur.core.classTeams)) cur.core.classTeams = [];
        reconcileClassTeamsFromYearClasses(cur, y, cur.years.byLabel[y].classes);
        return saveV2(cur);
    }

    function catalogLinkSameKey(a, b) {
        if (!a || !b || a.kind !== b.kind) return false;
        if (a.kind === 'sammelgruppe') return a.code === b.code;
        if (a.kind === 'cohort' || a.kind === 'eltern') return String(a.code) === String(b.code);
        return normCode(a.code) === normCode(b.code);
    }

    function getCatalogLink(kind, code) {
        const setup = getSetup();
        const links = (setup && setup.catalogLinks) || [];
        if (kind === 'sammelgruppe') {
            const c = normalizeSammelgruppeCode(code);
            if (!c) return null;
            for (let i = 0; i < links.length; i++) {
                if (links[i].kind === 'sammelgruppe' && links[i].code === c) return links[i];
            }
            return null;
        }
        if (kind === 'cohort' || kind === 'eltern') {
            const c = String(code || '').trim();
            if (!/^\d{4}$/.test(c)) return null;
            for (let i = 0; i < links.length; i++) {
                if (links[i].kind === kind && String(links[i].code) === c) return links[i];
            }
            return null;
        }
        const k = kind === 'arge' ? 'arge' : 'subject';
        const c = normCode(code);
        if (!c) return null;
        for (let i = 0; i < links.length; i++) {
            if (links[i].kind === k && normCode(links[i].code) === c) return links[i];
        }
        return null;
    }

    function upsertCatalogLink(entry) {
        const n = normalizeCatalogLink(entry);
        if (!n) return null;
        const cur = getSetup();
        const links = Array.isArray(cur.catalogLinks) ? cur.catalogLinks.slice() : [];
        const idx = links.findIndex(function (x) {
            return catalogLinkSameKey(x, n);
        });
        if (idx >= 0) links[idx] = n;
        else links.push(n);
        patchSetup({ catalogLinks: links });
        return n;
    }

    function clearCatalogLinkGroup(kind, code) {
        const existing = getCatalogLink(kind, code);
        if (!existing) return null;
        return upsertCatalogLink({
            kind: kind,
            code: code,
            graphGroupId: '',
            displayName: existing.displayName,
            mailNickname: existing.mailNickname,
            mode: ''
        });
    }

    function removeCatalogLink(kind, code) {
        const existing = getCatalogLink(kind, code);
        if (!existing) return false;
        const cur = getSetup();
        const links = Array.isArray(cur.catalogLinks) ? cur.catalogLinks.slice() : [];
        const next = links.filter(function (x) {
            return !catalogLinkSameKey(x, existing);
        });
        if (next.length === links.length) return false;
        patchSetup({ catalogLinks: next });
        return true;
    }

    function renameCatalogLink(kind, oldCode, newCode) {
        const k = kind === 'arge' ? 'arge' : 'subject';
        const from = normCode(oldCode);
        const to = normCode(newCode);
        if (!from || !to) return null;
        if (from === to) return getCatalogLink(k, from);
        const existing = getCatalogLink(k, from);
        if (!existing) return null;
        if (getCatalogLink(k, to)) {
            throw new Error('Ziel-Kürzel hat bereits eine Verknüpfung.');
        }
        const cur = getSetup();
        const links = Array.isArray(cur.catalogLinks) ? cur.catalogLinks.slice() : [];
        const idx = links.findIndex(function (x) {
            return catalogLinkSameKey(x, existing);
        });
        if (idx < 0) return null;
        const moved = normalizeCatalogLink(
            Object.assign({}, links[idx], {
                kind: k,
                code: to
            })
        );
        if (!moved) return null;
        links[idx] = moved;
        patchSetup({ catalogLinks: links });
        return moved;
    }

    function findClassTeamIndex(teams, classCode, abschlussJahr) {
        const code = normCode(classCode);
        const year = String(abschlussJahr || '').trim();
        if (!code) return -1;
        let fallback = -1;
        for (let i = 0; i < teams.length; i++) {
            if (normCode(teams[i].classCode) !== code) continue;
            if (year && teams[i].abschlussJahr && teams[i].abschlussJahr === year) return i;
            if (!year || !teams[i].abschlussJahr) fallback = i;
        }
        return fallback;
    }

    function patchClassTeamMeta(oldClassCode, oldAbschlussJahr, patch) {
        const c = getContainer();
        const teams = normalizeCoreClassTeams(c.core.classTeams || []);
        const idx = findClassTeamIndex(teams, oldClassCode, oldAbschlussJahr);
        if (idx < 0) return null;
        const p = patch && typeof patch === 'object' ? patch : {};
        const next = Object.assign({}, teams[idx]);
        if (p.classCode != null) next.classCode = normCode(p.classCode);
        if (p.displayName != null) next.displayName = String(p.displayName || '').trim();
        if (p.abschlussJahr != null) {
            const y = String(p.abschlussJahr || '').trim();
            next.abschlussJahr = /^\d{4}$/.test(y) ? y : '';
        }
        teams[idx] = next;
        c.core.classTeams = normalizeCoreClassTeams(teams);
        saveV2(c);
        return teams[idx];
    }

    function removeClassTeamByClassCode(classCode, abschlussJahr) {
        const c = getContainer();
        const teams = normalizeCoreClassTeams(c.core.classTeams || []);
        const idx = findClassTeamIndex(teams, classCode, abschlussJahr);
        if (idx < 0) return false;
        teams.splice(idx, 1);
        c.core.classTeams = normalizeCoreClassTeams(teams);
        saveV2(c);
        return true;
    }

    function getYearBucket(label) {
        const c = getContainer();
        const y = String(label || c.years.current || '').trim() || currentSchoolYearLabel();
        if (!c.years.byLabel[y]) c.years.byLabel[y] = emptyYearBucket();
        return { year: y, bucket: normalizeYearBucket(c.years.byLabel[y]) };
    }

    function saveYearBucket(label, bucket) {
        const c = getContainer();
        const y = String(label || c.years.current || '').trim() || currentSchoolYearLabel();
        c.years.byLabel[y] = normalizeYearBucket(bucket);
        return saveV2(c);
    }

    function upsertGuardian(entry, yearLabel) {
        const { year, bucket } = getYearBucket(yearLabel);
        const n = normalizeGuardian(entry);
        if (!n) throw new Error('Erziehungsberechtigte: Name oder E-Mail fehlt.');
        let idx = bucket.guardians.findIndex(function (g) {
            return g.id === n.id;
        });
        if (idx < 0 && n.email) {
            idx = bucket.guardians.findIndex(function (g) {
                return g.email && g.email === n.email;
            });
        }
        if (idx >= 0) {
            const prev = bucket.guardians[idx];
            bucket.guardians[idx] = Object.assign({}, prev, n, { id: prev.id });
            saveYearBucket(year, bucket);
            return bucket.guardians[idx];
        }
        bucket.guardians.push(n);
        saveYearBucket(year, bucket);
        return n;
    }

    function removeGuardian(guardianId, yearLabel) {
        const gid = String(guardianId || '').trim();
        if (!gid) return false;
        const { year, bucket } = getYearBucket(yearLabel);
        const before = bucket.guardians.length;
        bucket.guardians = bucket.guardians.filter(function (g) {
            return g.id !== gid;
        });
        bucket.students.forEach(function (s) {
            s.guardianIds = (s.guardianIds || []).filter(function (id) {
                return id !== gid;
            });
        });
        if (bucket.guardians.length === before) return false;
        saveYearBucket(year, bucket);
        return true;
    }

    function pruneUnlinkedGuardians(yearLabel) {
        const { year, bucket } = getYearBucket(yearLabel);
        const linked = new Set();
        bucket.students.forEach(function (s) {
            (Array.isArray(s.guardianIds) ? s.guardianIds : []).forEach(function (id) {
                const gid = String(id || '').trim();
                if (gid) linked.add(gid);
            });
        });
        const before = bucket.guardians.length;
        bucket.guardians = bucket.guardians.filter(function (g) {
            return linked.has(String((g && g.id) || '').trim());
        });
        const removed = before - bucket.guardians.length;
        if (removed > 0) saveYearBucket(year, bucket);
        return removed;
    }

    function removeStudents(studentIds, yearLabel) {
        const ids = new Set(
            (Array.isArray(studentIds) ? studentIds : [])
                .map(function (id) {
                    return String(id || '').trim();
                })
                .filter(Boolean)
        );
        if (!ids.size) return { removedStudents: 0, removedGuardians: 0 };
        const { year, bucket } = getYearBucket(yearLabel);
        const before = bucket.students.length;
        bucket.students = bucket.students.filter(function (s) {
            return !ids.has(String((s && s.id) || '').trim());
        });
        const removedStudents = before - bucket.students.length;
        if (removedStudents <= 0) return { removedStudents: 0, removedGuardians: 0 };
        saveYearBucket(year, bucket);
        const removedGuardians = pruneUnlinkedGuardians(year);
        return { removedStudents: removedStudents, removedGuardians: removedGuardians };
    }

    function setStudentGuardianIds(studentId, guardianIds, yearLabel) {
        const sid = String(studentId || '').trim();
        if (!sid) throw new Error('Schüler-ID fehlt.');
        const { year, bucket } = getYearBucket(yearLabel);
        const s = bucket.students.find(function (x) {
            return x.id === sid;
        });
        if (!s) throw new Error('Schüler nicht gefunden.');
        const valid = new Set(bucket.guardians.map(function (g) {
            return g.id;
        }));
        const seen = new Set();
        const next = [];
        (Array.isArray(guardianIds) ? guardianIds : []).forEach(function (id) {
            const g = String(id || '').trim();
            if (!g || !valid.has(g) || seen.has(g)) return;
            seen.add(g);
            next.push(g);
        });
        s.guardianIds = next;
        saveYearBucket(year, bucket);
        return s;
    }

    function linkGuardianToStudent(studentId, guardianEntry, yearLabel) {
        const { year, bucket } = getYearBucket(yearLabel);
        const sid = String(studentId || '').trim();
        const s = bucket.students.find(function (x) {
            return x.id === sid;
        });
        if (!s) throw new Error('Schüler nicht gefunden.');
        const n = normalizeGuardian(guardianEntry);
        if (!n || !n.email) throw new Error('E-Mail der Erziehungsberechtigten fehlt.');
        let g = bucket.guardians.find(function (x) {
            return x.email === n.email;
        });
        if (!g) {
            bucket.guardians.push(n);
            g = n;
        } else if (n.name && !g.name) {
            g.name = n.name;
        }
        if (!s.guardianIds.includes(g.id)) s.guardianIds.push(g.id);
        saveYearBucket(year, bucket);
        return { student: s, guardian: g };
    }

    function unlinkGuardianFromStudent(studentId, guardianId, yearLabel) {
        const sid = String(studentId || '').trim();
        const gid = String(guardianId || '').trim();
        const { year, bucket } = getYearBucket(yearLabel);
        const s = bucket.students.find(function (x) {
            return x.id === sid;
        });
        if (!s) return false;
        const before = s.guardianIds.length;
        s.guardianIds = (s.guardianIds || []).filter(function (id) {
            return id !== gid;
        });
        if (s.guardianIds.length === before) return false;
        saveYearBucket(year, bucket);
        return true;
    }

    function upsertParentList(entry, yearLabel) {
        const { year, bucket } = getYearBucket(yearLabel);
        const n = normalizeParentList(entry);
        if (!n) throw new Error('Elternliste: scope/code ungültig.');
        const idx = bucket.parentLists.findIndex(function (p) {
            return p.scope === n.scope && String(p.code) === String(n.code);
        });
        if (idx >= 0) bucket.parentLists[idx] = Object.assign({}, bucket.parentLists[idx], n);
        else bucket.parentLists.push(n);
        saveYearBucket(year, bucket);
        return n;
    }

    window.ms365AppDataV2 = {
        STORAGE_KEY_V2,
        VERSION,
        getContainer,
        setContainer,
        exportJson,
        importJson,
        setCoreFromTenantSettings,
        listYears,
        setCurrentYear,
        getSetup,
        patchSetup,
        touchWizardVisit,
        defaultSetup,
        normalizeSetup,
        normalizeClassTeam,
        normalizeCoreClassTeams,
        upsertClassTeam,
        getClassTeamGruppenmailForKlasse,
        reconcileClassTeamsFromYearClasses,
        mailNicknamePrefixSanitize,
        getCatalogLink,
        upsertCatalogLink,
        clearCatalogLinkGroup,
        removeCatalogLink,
        renameCatalogLink,
        patchClassTeamMeta,
        removeClassTeamByClassCode,
        emptyYearBucket,
        normalizeYearBucket,
        normalizeGuardian,
        normalizeStudentRow,
        mergeStudentsImport,
        getYearBucket,
        saveYearBucket,
        upsertGuardian,
        removeGuardian,
        pruneUnlinkedGuardians,
        removeStudents,
        setStudentGuardianIds,
        linkGuardianToStudent,
        unlinkGuardianFromStudent,
        upsertParentList
    };
})();

