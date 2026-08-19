(function () {
    'use strict';

    var DEMO_MODE_KEY = 'ms365-demo-mode-v1';
    var TENANT_KEY = 'ms365-tenant-settings-v1';
    var APP_DATA_KEY = 'ms365-schooltool-data-v2';
    var KURSTEAM_KEY = 'webuntis-teams-creator-state-v1';
    var DOMAIN = 'ms365.schule';

    function em(local) {
        return String(local || '').trim().toLowerCase() + '@' + DOMAIN;
    }

    /** Deterministische Demo-GUIDs (nur lokal, kein echter Tenant). */
    function demoGuid(n) {
        var s = String(n >>> 0).padStart(12, '0');
        return 'demo0000-0000-4000-8000-' + s;
    }

    function buildDemoStudents() {
        var classCodes = ['1A', '1B', '1C', '2A', '2B', '3A', '4A', '5A'];
        var first = [
            'Anna', 'Ben', 'Carla', 'David', 'Eva', 'Felix', 'Greta', 'Hugo',
            'Ina', 'Jonas', 'Klara', 'Lukas', 'Mia', 'Noah', 'Olivia', 'Paul',
            'Quentin', 'Rosa', 'Simon', 'Tina', 'Uwe', 'Vera', 'Willi', 'Xenia'
        ];
        var last = ['Beispiel', 'Demo', 'Muster', 'Test', 'Probe'];
        var out = [];
        var gi = 0;
        classCodes.forEach(function (klasse) {
            var count = klasse.indexOf('5') === 0 ? 2 : 3;
            for (var i = 0; i < count; i++) {
                var fn = first[gi % first.length];
                var ln = last[gi % last.length];
                var slug = (fn + '.' + ln + '.' + klasse).toLowerCase().replace(/[^a-z0-9.]/g, '');
                out.push({
                    klasse: klasse,
                    name: fn + ' ' + ln,
                    email: em(slug)
                });
                gi += 1;
            }
        });
        return out;
    }

    function buildDemoGuardians(students) {
        var guardians = [];
        var seen = {};
        (students || []).forEach(function (s, idx) {
            var base = String(s.email || '').split('@')[0] || 'eltern' + idx;
            var g1 = {
                id: 'g_demo_' + idx + '_1',
                name: 'Eltern von ' + String(s.name || '').split(' ')[0],
                email: em('eltern.' + base),
                phone: '+43 660 000' + String(1000 + idx).slice(-4),
                note: 'Demo-Erziehungsberechtigte'
            };
            var g2 = {
                id: 'g_demo_' + idx + '_2',
                name: 'Partner/in ' + String(s.name || '').split(' ')[0],
                email: em('partner.' + base),
                phone: '',
                note: ''
            };
            if (!seen[g1.email]) {
                guardians.push(g1);
                seen[g1.email] = true;
            }
            if (!seen[g2.email]) {
                guardians.push(g2);
                seen[g2.email] = true;
            }
        });
        return guardians;
    }

    function linkStudentsToGuardians(students, guardians) {
        var byEmail = {};
        guardians.forEach(function (g) {
            if (g && g.email) byEmail[g.email] = g.id;
        });
        return (students || []).map(function (s, idx) {
            var row = Object.assign({}, s);
            row.id = 'stu_demo_' + idx;
            var base = String(s.email || '').split('@')[0];
            var g1 = byEmail[em('eltern.' + base)];
            var g2 = byEmail[em('partner.' + base)];
            row.guardianIds = [g1, g2].filter(Boolean);
            return row;
        });
    }

    function getDemoTenantData() {
        var students = buildDemoStudents();
        var studentCouncil = students.filter(function (_, i) {
            return i % 4 === 0;
        }).slice(0, 8);

        return {
            schoolName: 'MS365 Musterschule',
            domain: DOMAIN,
            subjects: [
                { code: 'M', name: 'Mathematik' },
                { code: 'D', name: 'Deutsch' },
                { code: 'E', name: 'Englisch' },
                { code: 'F', name: 'Französisch' },
                { code: 'BIO', name: 'Biologie' },
                { code: 'CH', name: 'Chemie' },
                { code: 'PH', name: 'Physik' },
                { code: 'GES', name: 'Geschichte' },
                { code: 'GEO', name: 'Geographie' },
                { code: 'MUS', name: 'Musik' },
                { code: 'KU', name: 'Kunst' },
                { code: 'INF', name: 'Informatik' },
                { code: 'REL', name: 'Religion' },
                { code: 'SPO', name: 'Sport' },
                { code: 'LAT', name: 'Latein' }
            ],
            arges: [
                { code: 'SPRACHEN', name: 'Sprachen', subjects: ['D', 'E', 'F', 'LAT'] },
                { code: 'NAWI', name: 'Naturwissenschaften', subjects: ['BIO', 'CH', 'PH'] },
                { code: 'GESELL', name: 'Geistes- & Gesellschaftswissenschaften', subjects: ['GES', 'GEO', 'REL'] },
                { code: 'KUNST', name: 'Kunst & Musik', subjects: ['MUS', 'KU'] },
                { code: 'DIGITAL', name: 'Digitales', subjects: ['INF'] }
            ],
            teachers: [
                { code: 'LEH', name: 'Vorname Lehrer', email: em('vorname.lehrer') },
                { code: 'MUS', name: 'Max Muster', email: em('max.muster') },
                { code: 'HUB', name: 'Hannah Huber', email: em('hannah.huber') },
                { code: 'WEG', name: 'Wolfgang Wagner', email: em('wolfgang.wagner') },
                { code: 'SCH', name: 'Sandra Schmidt', email: em('sandra.schmidt') },
                { code: 'BRA', name: 'Bruno Braun', email: em('bruno.braun') },
                { code: 'FIS', name: 'Fiona Fischer', email: em('fiona.fischer') },
                { code: 'GRU', name: 'Gregor Grün', email: em('gregor.gruen') },
                { code: 'WEI', name: 'Wendy Weiss', email: em('wendy.weiss') },
                { code: 'KOE', name: 'Karl Koch', email: em('karl.koch') },
                { code: 'MEI', name: 'Maria Meier', email: em('maria.meier') },
                { code: 'LANG', name: 'Lisa Lang', email: em('lisa.lang') }
            ],
            administration: [
                {
                    code: 'DIREKTION',
                    name: 'Direktion',
                    people: [{ name: 'Wolfgang Wagner', email: em('wolfgang.wagner') }]
                },
                {
                    code: 'SEKRETARIAT',
                    name: 'Sekretariat',
                    people: [
                        { name: 'Sandra Schmidt', email: em('sandra.schmidt') },
                        { name: 'Petra Post', email: em('petra.post') }
                    ]
                },
                {
                    code: 'IT-SUPPORT',
                    name: 'IT-Support',
                    people: [{ name: 'Karl Koch', email: em('karl.koch') }]
                },
                {
                    code: 'BIBLIOTHEK',
                    name: 'Bibliothek',
                    people: [{ name: 'Berta Buch', email: em('berta.buch') }]
                }
            ],
            sgaMode: 'group',
            sga: [
                { scope: 'teacher', name: 'Vorname Lehrer', email: em('vorname.lehrer') },
                { scope: 'teacher', name: 'Max Muster', email: em('max.muster') },
                { scope: 'student', name: 'Anna Beispiel', email: students[0] ? students[0].email : em('anna.beispiel.1a') },
                { scope: 'student', name: 'Ben Demo', email: students[1] ? students[1].email : em('ben.demo.1a') },
                { scope: 'external', name: 'Eva Extern', email: 'eva.extern@example.org' }
            ],
            students: students,
            studentCouncil: studentCouncil,
            classes: [
                { code: '1A', year: '2030', name: '1A', headName: 'Vorname Lehrer', headEmail: em('vorname.lehrer') },
                { code: '1B', year: '2030', name: '1B', headName: 'Max Muster', headEmail: em('max.muster') },
                { code: '1C', year: '2030', name: '1C', headName: 'Hannah Huber', headEmail: em('hannah.huber') },
                { code: '2A', year: '2029', name: '2A', headName: 'Bruno Braun', headEmail: em('bruno.braun') },
                { code: '2B', year: '2029', name: '2B', headName: 'Fiona Fischer', headEmail: em('fiona.fischer') },
                { code: '3A', year: '2028', name: '3A', headName: 'Gregor Grün', headEmail: em('gregor.gruen') },
                { code: '4A', year: '2027', name: '4A', headName: 'Wendy Weiss', headEmail: em('wendy.weiss') },
                { code: '5A', year: '2026', name: '5A', headName: 'Maria Meier', headEmail: em('maria.meier') }
            ]
        };
    }

    function buildDemoCatalogLinks(tenant) {
        var links = [];
        var seq = 1;

        function push(link) {
            links.push(link);
            seq += 1;
        }

        push({
            kind: 'sammelgruppe',
            code: 'schueler',
            graphGroupId: demoGuid(seq),
            displayName: 'Alle Schülerinnen und Schüler',
            mailNickname: 'schueler',
            mode: 'matched',
            syncStatus: 'demo'
        });
        push({
            kind: 'sammelgruppe',
            code: 'lehrer',
            graphGroupId: demoGuid(seq),
            displayName: 'Alle Lehrkräfte',
            mailNickname: 'lehrer',
            mode: 'matched',
            syncStatus: 'demo'
        });
        push({
            kind: 'sammelgruppe',
            code: 'verwaltung',
            graphGroupId: demoGuid(seq),
            displayName: 'Schulverwaltung',
            mailNickname: 'verwaltung',
            mode: 'matched',
            syncStatus: 'demo'
        });

        (tenant.subjects || []).forEach(function (s) {
            if (!s || !s.code) return;
            push({
                kind: 'subject',
                code: s.code,
                graphGroupId: demoGuid(seq),
                displayName: 'Fach ' + (s.name || s.code),
                mailNickname: 'fach-' + String(s.code).toLowerCase(),
                mode: 'matched',
                syncStatus: 'demo'
            });
        });

        (tenant.arges || []).forEach(function (a) {
            if (!a || !a.code) return;
            push({
                kind: 'arge',
                code: a.code,
                graphGroupId: demoGuid(seq),
                displayName: 'ARGE ' + (a.name || a.code),
                mailNickname: 'ag-' + String(a.code).toLowerCase(),
                mode: 'matched',
                syncStatus: 'demo'
            });
        });

        ['2030', '2029', '2028'].forEach(function (yr) {
            push({
                kind: 'cohort',
                code: yr,
                graphGroupId: demoGuid(seq),
                displayName: 'Jahrgang ' + yr,
                mailNickname: 'jg' + yr,
                mode: 'matched',
                syncStatus: 'demo'
            });
            push({
                kind: 'eltern',
                code: yr,
                graphGroupId: demoGuid(seq),
                displayName: 'Eltern JG ' + yr,
                mailNickname: 'elternjg' + yr,
                mode: 'matched',
                syncStatus: 'demo'
            });
        });

        return links;
    }

    function buildDemoDirectoryMatches(tenant) {
        var map = {};
        var seq = 500;
        function addPerson(email, displayName) {
            var eml = String(email || '').trim().toLowerCase();
            if (!eml) return;
            map[eml] = {
                objectId: demoGuid(seq),
                displayName: displayName || eml,
                matchedAt: new Date().toISOString(),
                demo: true
            };
            seq += 1;
        }
        (tenant.teachers || []).forEach(function (t) {
            addPerson(t.email, t.name);
        });
        (tenant.students || []).forEach(function (s) {
            addPerson(s.email, s.name);
        });
        (tenant.administration || []).forEach(function (g) {
            (g.people || []).forEach(function (p) {
                addPerson(p.email, p.name);
            });
        });
        return map;
    }

    function buildDemoClassGroupMatches(tenant) {
        var map = {};
        var seq = 800;
        (tenant.classes || []).forEach(function (c) {
            if (!c || !c.code) return;
            var nick =
                String(c.stableMailNickname || '').trim() ||
                ('jg' + String(c.year || '2030') + String(c.code).toLowerCase());
            map[String(c.code).toUpperCase()] = {
                graphGroupId: demoGuid(seq),
                displayName: 'Klasse ' + (c.name || c.code),
                mailNickname: nick.replace(/[^a-z0-9]/gi, '').toLowerCase(),
                mode: 'matched',
                matchedAt: new Date().toISOString(),
                demo: true
            };
            seq += 1;
        });
        return map;
    }

    function currentSchoolYearLabel() {
        var y = new Date().getFullYear();
        return String(y) + '/' + String(y + 1).slice(2);
    }

    function previousSchoolYearLabel() {
        var y = new Date().getFullYear() - 1;
        return String(y) + '/' + String(y + 1).slice(2);
    }

    function seedDemoAppData(tenant) {
        if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.getContainer !== 'function') return;

        var api = window.ms365AppDataV2;
        var curYear = currentSchoolYearLabel();
        var prevYear = previousSchoolYearLabel();
        var guardians = buildDemoGuardians(tenant.students);
        var studentsLinked = linkStudentsToGuardians(tenant.students, guardians);

        var parentLists = [
            { scope: 'class', code: '1A', displayName: 'Eltern 1A', mailNickname: 'eltern1a', graphGroupId: demoGuid(901) },
            { scope: 'class', code: '1B', displayName: 'Eltern 1B', mailNickname: 'eltern1b', graphGroupId: demoGuid(902) },
            { scope: 'class', code: '2A', displayName: 'Eltern 2A', mailNickname: 'eltern2a', graphGroupId: demoGuid(903) },
            { scope: 'year', code: '2030', displayName: 'Eltern JG 2030', mailNickname: 'elternjg2030', graphGroupId: demoGuid(904) }
        ];

        try {
            if (typeof api.setCurrentYear === 'function') {
                api.setCurrentYear(curYear);
            }
        } catch {
            /* ignore */
        }

        var bucket = {
            students: studentsLinked,
            studentCouncil: Array.isArray(tenant.studentCouncil) ? tenant.studentCouncil.slice() : [],
            classes: Array.isArray(tenant.classes) ? tenant.classes.slice() : [],
            guardians: guardians,
            parentLists: parentLists
        };

        if (typeof api.saveYearBucket === 'function') {
            api.saveYearBucket(curYear, bucket);
        }

        if (typeof api.setCurrentYear === 'function') {
            try {
                api.setCurrentYear(prevYear, { copyFrom: curYear });
                var prevBucket = api.getYearBucket(prevYear);
                if (prevBucket && prevBucket.bucket) {
                    prevBucket.bucket.classes = (tenant.classes || []).slice(0, 4).map(function (c) {
                        var row = Object.assign({}, c);
                        if (row.code === '1A') row.name = '0A';
                        if (row.code === '1B') row.name = '0B';
                        return row;
                    });
                    api.saveYearBucket(prevYear, prevBucket.bucket);
                }
                api.setCurrentYear(curYear);
            } catch {
                /* ignore */
            }
        }

        var catalogLinks = buildDemoCatalogLinks(tenant);
        var matched = {
            schuelerGroupId: catalogLinks[0] ? catalogLinks[0].graphGroupId : null,
            lehrerGroupId: catalogLinks[1] ? catalogLinks[1].graphGroupId : null,
            verwaltungGroupId: catalogLinks[2] ? catalogLinks[2].graphGroupId : null,
            sgaGroupId: demoGuid(950),
            studentCouncilGroupId: demoGuid(951)
        };

        api.patchSetup({
            wizardStep: 11,
            _einrichtungWizardLayout: 11,
            finishedAt: new Date().toISOString(),
            completedSteps: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11],
            matched: matched,
            catalogLinks: catalogLinks,
            directoryMatchByEmail: buildDemoDirectoryMatches(tenant),
            classGroupMatchByKey: buildDemoClassGroupMatches(tenant),
            intranetSiteUrl: 'https://' + DOMAIN.replace(/\./g, '') + '.sharepoint.com/sites/intranet',
            actionLog: [
                {
                    at: new Date().toISOString(),
                    tool: 'demo',
                    action: 'seed',
                    target: 'MS365 Musterschule',
                    summary: 'Demo-Datenbank geladen',
                    result: 'ok'
                }
            ]
        });

        var c = api.getContainer();
        if (c && c.core && Array.isArray(c.core.classTeams)) {
            c.core.classTeams = c.core.classTeams.map(function (ct, idx) {
                return Object.assign({}, ct, {
                    graphGroupId: ct.graphGroupId || demoGuid(1000 + idx),
                    mode: ct.mode || 'matched',
                    syncStatus: 'demo'
                });
            });
        }
        if (c && c.structure) {
            c.structure.settings = Object.assign({}, c.structure.settings || {}, {
                organisationAssist: {
                    cohortPlans: [],
                    playbook: {
                        targetYear: curYear,
                        done: {
                            year: true,
                            names: false,
                            graduates: false,
                            students: true,
                            kursteams: false,
                            subjects: true,
                            expert: false
                        }
                    },
                    runLog: []
                }
            });
        }
        if (typeof api.setContainer === 'function') {
            api.setContainer(c);
        }
    }

    function seedDemoKursteamState(tenant) {
        var teachers = tenant.teachers || [];
        var subjects = tenant.subjects || [];
        var classes = tenant.classes || [];
        function code(list, i, fb) {
            var row = list[i];
            return row && row.code ? String(row.code).toUpperCase() : fb;
        }
        var lines = [
            code(teachers, 0, 'LEH') + '\t' + code(subjects, 0, 'M') + '\t' + (classes[0] && classes[0].code ? classes[0].code : '1A'),
            code(teachers, 0, 'LEH') + '\t' + code(subjects, 1, 'D') + '\t' + (classes[1] && classes[1].code ? classes[1].code : '1B'),
            code(teachers, 0, 'LEH') + '\t' + code(subjects, 2, 'E') + '\t' + (classes[2] && classes[2].code ? classes[2].code : '1C'),
            code(teachers, 1, 'MUS') + '\t' + code(subjects, 0, 'M') + '\t' + (classes[3] && classes[3].code ? classes[3].code : '2A'),
            code(teachers, 2, 'HUB') + '\t' + code(subjects, 3, 'BIO') + '\t' + (classes[0] && classes[0].code ? classes[0].code : '1A'),
            code(teachers, 3, 'BRA') + '\t' + code(subjects, 4, 'CH') + '\t' + (classes[4] && classes[4].code ? classes[4].code : '2B'),
            code(teachers, 4, 'FIS') + '\t' + code(subjects, 1, 'D') + '\t' + (classes[5] && classes[5].code ? classes[5].code : '3A'),
            code(teachers, 5, 'GRU') + '\t' + code(subjects, 2, 'E') + '\t' + (classes[6] && classes[6].code ? classes[6].code : '4A')
        ];
        var mapping = {};
        teachers.forEach(function (t) {
            if (t && t.code && t.email) mapping[String(t.code).toUpperCase()] = t.email;
        });
        var y = new Date().getFullYear();
        var state = {
            stepSchema: 2,
            rawData: [],
            filteredData: [],
            teamsData: [],
            teacherEmailMapping: mapping,
            teamsGenerated: false,
            currentStep: 0,
            yearPrefix: 'SJ' + String(y).slice(2),
            schoolDomain: tenant.domain || DOMAIN,
            teamSeparator: '-',
            teamNamePattern: null,
            excludeSubjects: '',
            removeDuplicates: true,
            kursteamEntryMode: 'webuntis',
            studentRosterRaw: (tenant.students || [])
                .map(function (s) {
                    return s.klasse + ';' + s.email;
                })
                .join('\n'),
            studentRosterPreferGroup: true,
            studentRosterSkipCombinedClasses: true,
            studentRosterHideNoMatch: true,
            studentRosterTeamSelection: {},
            webuntisPaste: lines.join('\n')
        };
        try {
            localStorage.setItem(KURSTEAM_KEY, JSON.stringify(state));
        } catch {
            /* ignore */
        }
    }

    function isActive() {
        try {
            return localStorage.getItem(DEMO_MODE_KEY) === '1';
        } catch {
            return false;
        }
    }

    function hasMeaningfulTenantData(settings) {
        if (!settings || typeof settings !== 'object') return false;
        var domain = String(settings.domain || '').trim();
        var schoolName = String(settings.schoolName || '').trim();
        var subjects = Array.isArray(settings.subjects) ? settings.subjects.length : 0;
        var teachers = Array.isArray(settings.teachers) ? settings.teachers.length : 0;
        var students = Array.isArray(settings.students) ? settings.students.length : 0;
        var classes = Array.isArray(settings.classes) ? settings.classes.length : 0;
        return !!(domain || schoolName || subjects || teachers || students || classes);
    }

    function activate() {
        if (typeof window.ms365TenantSettingsSave !== 'function') return false;
        var tenant = getDemoTenantData();
        var saved = window.ms365TenantSettingsSave(tenant);
        seedDemoAppData(tenant);
        seedDemoKursteamState(tenant);
        try {
            localStorage.setItem(DEMO_MODE_KEY, '1');
        } catch {
            return false;
        }
        try {
            if (typeof window.ms365SetSchoolDomainNoAt === 'function' && saved && saved.domain) {
                window.ms365SetSchoolDomainNoAt(saved.domain);
            }
        } catch {
            /* ignore */
        }
        try {
            window.dispatchEvent(
                new CustomEvent('ms365-demo-mode-changed', { detail: { active: true } })
            );
        } catch {
            /* ignore */
        }
        return true;
    }

    function confirmExit() {
        var msg =
            'Demo beenden und Beispieldaten aus diesem Browser entfernen?\n\nIhre echten Schuldaten (falls vorhanden) bleiben unberührt – es werden nur die Demo-Daten gelöscht.';
        if (typeof window.ms365AppDialogConfirm === 'function') {
            return window.ms365AppDialogConfirm(msg, { title: 'Demo beenden' });
        }
        return Promise.resolve(window.confirm(msg));
    }

    function deactivate() {
        return confirmExit().then(function (ok) {
            if (!ok) return false;
            try {
                localStorage.removeItem(DEMO_MODE_KEY);
                localStorage.removeItem(TENANT_KEY);
                localStorage.removeItem(APP_DATA_KEY);
                localStorage.removeItem(KURSTEAM_KEY);
            } catch {
                return false;
            }
            try {
                window.dispatchEvent(
                    new CustomEvent('ms365-demo-mode-changed', { detail: { active: false } })
                );
            } catch {
                /* ignore */
            }
            return true;
        });
    }

    window.ms365DemoMode = {
        isActive: isActive,
        hasMeaningfulTenantData: hasMeaningfulTenantData,
        getDemoTenantData: getDemoTenantData,
        activate: activate,
        deactivate: deactivate
    };
})();
