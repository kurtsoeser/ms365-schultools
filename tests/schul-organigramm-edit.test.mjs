import { describe, expect, it } from 'vitest';
import {
    findAdminRole,
    peopleForAdminRole,
    findArge,
    findTeacherByEmail,
    findClass,
    findStudentCouncil,
    findSga,
    listPeopleForSchulOrgNode,
    guardiansForClass,
    teachersForClassChat,
    teachersForArge,
    listM365PeopleForArea,
    listMailboxDelegatePeopleForArea,
    previewPeopleForSchulOrgNode,
    moveSchulOrgPerson,
    moveSchulOrgRole,
    SCHUL_ORG_PEOPLE_PREVIEW_LIMIT
} from '../src/tools/schul-organigramm/schul-organigramm-edit.js';
import {
    seedSchulOrganigrammFromStammdaten,
    SCHUL_ORG_BOARD_REF,
    audienceSectionId,
    elternClassNodeId,
    klassenchatClassNodeId
} from '../src/tools/schul-organigramm/schul-organigramm-logic.js';

describe('schul-organigramm-edit helpers', () => {
    const settings = {
        adminRoles: [
            { code: 'DIREKTION', name: 'Direktion', seatMode: 'single' },
            { code: 'SEK', name: 'Sekretariat', seatMode: 'multi' }
        ],
        admin: [
            { role: 'Direktion', name: 'Holtmann', email: 'h@schule.at' },
            { role: 'Sekretariat', name: 'Anna', email: 'a@schule.at' },
            { role: 'Sekretariat', name: 'Ben', email: 'b@schule.at' },
            { role: 'Sekretariat', name: '', email: '' }
        ],
        teachers: [
            { code: 'MUEL', name: 'Müller', email: 'mueller@schule.at' },
            { code: 'SCHM', name: 'Schmidt', email: 'schmidt@schule.at' }
        ],
        arges: [
            {
                code: 'MATHE',
                name: 'Mathematik',
                subjects: ['MATHE'],
                headName: 'Müller',
                headEmail: 'mueller@schule.at'
            },
            {
                code: 'DEUT',
                name: 'Deutsch',
                subjects: ['DEUT'],
                headName: '',
                headEmail: ''
            },
            {
                code: 'INFO',
                name: 'Informatik',
                subjects: ['INFO'],
                headName: 'Extern',
                headEmail: 'extern@andere.at'
            }
        ],
        classes: [
            {
                code: '1A',
                name: '1A',
                year: '2034',
                headName: 'Schmidt',
                headEmail: 'schmidt@schule.at'
            },
            { code: '2B', name: '2B', year: '', headName: '', headEmail: '' }
        ],
        students: [
            {
                klasse: '1A',
                name: 'Lea',
                email: 'lea@schule.at',
                parentPairs: [
                    { name: 'Eltern Lea', email: 'eltern.lea@familie.at' },
                    { name: 'Vater Lea', email: 'vater.lea@familie.at' }
                ]
            },
            { klasse: '1A', name: 'Max', email: 'max@schule.at', parentPairs: [{ name: 'Eltern Max', email: 'eltern.max@familie.at' }] },
            { klasse: '1A', name: 'Nina', email: 'nina@schule.at' },
            { klasse: '1A', name: 'Otto', email: 'otto@schule.at' },
            {
                klasse: '2B',
                name: 'Tom',
                email: 'tom@schule.at',
                parentPairs: [{ name: 'Eltern Tom', email: 'eltern.tom@familie.at' }]
            }
        ],
        studentCouncil: [
            { klasse: '1A', name: 'Lea', email: 'lea@schule.at' },
            { klasse: '2B', name: 'Tom', email: '' }
        ],
        sga: [
            { scope: 'teacher', name: 'Müller', email: 'mueller@schule.at' },
            { scope: 'student', name: 'Sam', email: 'sam@schule.at' }
        ],
        unterrichtsbelegung: {
            rows: [
                {
                    klasse: '1A',
                    fach: 'MAT',
                    lehrerEmail: 'mueller@schule.at',
                    lehrerCode: 'MUEL',
                    lehrerName: 'Müller'
                },
                {
                    klasse: '1A',
                    fach: 'DEU',
                    lehrerEmail: 'schmidt@schule.at',
                    lehrerCode: 'SCHM',
                    lehrerName: 'Schmidt'
                }
            ]
        },
        classChats: {
            items: [
                {
                    klasse: '1A',
                    chatId: 'chat-1a',
                    topic: 'SJ26 | 1A | Lehrer',
                    memberEmails: ['mueller@schule.at', 'schmidt@schule.at']
                }
            ]
        }
    };

    it('findAdminRole findet per Code und Name', () => {
        expect(findAdminRole(settings, 'SEK')?.name).toBe('Sekretariat');
        expect(findAdminRole(settings, 'direktion')?.code).toBe('DIREKTION');
        expect(findAdminRole(settings, 'XYZ')).toBeNull();
    });

    it('peopleForAdminRole liefert alle Zeilen der Rolle inkl. Entwurf', () => {
        const role = findAdminRole(settings, 'SEK');
        const people = peopleForAdminRole(settings, role);
        expect(people).toHaveLength(3);
        expect(people.filter((p) => p.email).map((p) => p.name).sort()).toEqual(['Anna', 'Ben']);
    });

    it('findArge findet per Code und Name', () => {
        expect(findArge(settings, 'MATHE')?.name).toBe('Mathematik');
        expect(findArge(settings, 'deutsch')?.code).toBe('DEUT');
        expect(findArge(settings, 'XYZ')).toBeNull();
    });

    it('findTeacherByEmail matched Lehrerliste', () => {
        expect(findTeacherByEmail(settings, 'mueller@schule.at')?.code).toBe('MUEL');
        expect(findTeacherByEmail(settings, 'MUELLER@SCHULE.AT')?.name).toBe('Müller');
        expect(findTeacherByEmail(settings, 'extern@andere.at')).toBeNull();
        expect(findTeacherByEmail(settings, '')).toBeNull();
    });

    it('findClass findet per Code und Name', () => {
        expect(findClass(settings, '1A')?.year).toBe('2034');
        expect(findClass(settings, '2b')?.code).toBe('2B');
        expect(findClass(settings, '9Z')).toBeNull();
    });

    it('findStudentCouncil / findSga finden per E-Mail und Name', () => {
        expect(findStudentCouncil(settings, 'lea@schule.at')?.klasse).toBe('1A');
        expect(findStudentCouncil(settings, 'Tom')?.klasse).toBe('2B');
        expect(findSga(settings, 'sam@schule.at')?.scope).toBe('student');
        expect(findSga(settings, 'Müller')?.scope).toBe('teacher');
        expect(findSga(settings, 'xyz')).toBeNull();
    });

    it('Seed-Meta: Lehrer-Bezug ARGE und Klasse', () => {
        const layout = seedSchulOrganigrammFromStammdaten(settings);
        const mathe = layout.nodes.find((n) => n.refKind === 'arge' && n.refKey === 'MATHE');
        const deut = layout.nodes.find((n) => n.refKind === 'arge' && n.refKey === 'DEUT');
        const info = layout.nodes.find((n) => n.refKind === 'arge' && n.refKey === 'INFO');
        expect(mathe?.type).toBe('area');
        const c1 = layout.nodes.find((n) => n.refKind === 'class' && n.refKey === '1A');
        const c2 = layout.nodes.find((n) => n.refKind === 'class' && n.refKey === '2B');
        expect(c1?.type).toBe('area');
        expect(mathe?.meta?.teacherLinked).toBe(true);
        expect(mathe?.meta?.teacherCode).toBe('MUEL');
        expect(deut?.meta?.teacherLinked).toBe(false);
        expect(info?.meta?.teacherLinked).toBe(false);
        expect(c1?.meta?.teacherLinked).toBe(true);
        expect(c1?.meta?.teacherCode).toBe('SCHM');
        expect(c2?.meta?.headEmail).toBe('');
        expect(c2?.meta?.teacherLinked).toBe(false);
    });

    it('listPeopleForSchulOrgNode: Rolle, ARGE-Leitung, Klasse inkl. Schüler', () => {
        const layout = seedSchulOrganigrammFromStammdaten(settings);
        const sek = layout.nodes.find((n) => n.type === 'role' && n.refKey === 'SEK');
        const mathe = layout.nodes.find((n) => n.refKind === 'arge' && n.refKey === 'MATHE');
        const deut = layout.nodes.find((n) => n.refKind === 'arge' && n.refKey === 'DEUT');
        const c1 = layout.nodes.find((n) => n.refKind === 'class' && n.refKey === '1A');
        const rolePeople = listPeopleForSchulOrgNode(settings, sek);
        expect(rolePeople.map((p) => p.name).sort()).toEqual(['Anna', 'Ben']);
        expect(listPeopleForSchulOrgNode(settings, mathe)[0]?.kind).toBe('head');
        expect(listPeopleForSchulOrgNode(settings, mathe)[0]?.seatLabel).toBe('ARGE-Leitung');
        const vacantArge = listPeopleForSchulOrgNode(settings, deut);
        expect(vacantArge).toHaveLength(1);
        expect(vacantArge[0].vacant).toBe(true);
        expect(vacantArge[0].seatLabel).toBe('ARGE-Leitung');
        const withBelegung = Object.assign({}, settings, {
            unterrichtsbelegung: {
                rows: [
                    {
                        klasse: '1A',
                        fach: 'MATHE',
                        lehrerEmail: 'schmidt@schule.at',
                        lehrerName: 'Schmidt',
                        lehrerCode: 'SCHM'
                    }
                ]
            }
        });
        const mathePeople = listPeopleForSchulOrgNode(withBelegung, mathe);
        expect(mathePeople[0]?.kind).toBe('head');
        expect(mathePeople.map((p) => p.email)).toContain('schmidt@schule.at');
        expect(teachersForArge(withBelegung, 'MATHE').map((t) => t.role)).toEqual(['head', 'lehrer']);
        const classPeople = listPeopleForSchulOrgNode(settings, c1);
        expect(classPeople[0]?.kind).toBe('head');
        expect(classPeople[0]?.seatLabel).toBe('Klassenvorstand');
        expect(classPeople.some((p) => p.functionId === 'klassensprecher' && p.vacant)).toBe(true);
        expect(classPeople.filter((p) => p.kind === 'student')).toHaveLength(4);
        // KV + offener Klassensprecher-Sitz + 4 Schüler
        expect(classPeople).toHaveLength(6);
        const preview = previewPeopleForSchulOrgNode(settings, c1);
        expect(preview.shown).toHaveLength(SCHUL_ORG_PEOPLE_PREVIEW_LIMIT);
        expect(preview.more).toBe(6 - SCHUL_ORG_PEOPLE_PREVIEW_LIMIT);
    });

    it('listM365PeopleForArea: Rollenpersonen unter dem Bereich', () => {
        const layout = seedSchulOrganigrammFromStammdaten(settings);
        const verwaltung = layout.nodes.find((n) => n.id === 'sec_verwaltung');
        const seeded = { ...settings, schulOrganigramm: layout };
        const dels = listM365PeopleForArea(seeded, verwaltung);
        expect(dels.map((p) => p.email).sort()).toEqual(['a@schule.at', 'b@schule.at']);
        expect(dels.find((p) => p.email === 'a@schule.at')?.source).toBe('role');
        expect(listMailboxDelegatePeopleForArea(seeded, verwaltung)).toEqual(dels);
    });

    it('listM365PeopleForArea: Personen aus Unterbereichen flach mitnehmen', () => {
        const layout = seedSchulOrganigrammFromStammdaten(settings);
        layout.nodes.push(
            {
                id: 'sec_sek_nested',
                type: 'area',
                label: 'Sekretariat innen',
                parentId: 'sec_verwaltung',
                order: 90,
                refKind: 'none',
                refKey: '',
                m365: { groupMode: 'none', mailboxMode: 'none' }
            },
            {
                id: 'person_eva',
                type: 'person',
                label: 'Eva Nested',
                parentId: 'sec_sek_nested',
                order: 1,
                refKind: 'adminPerson',
                refKey: 'eva@schule.at'
            },
            {
                id: 'role_nested_it',
                type: 'role',
                label: 'IT-Hilfe',
                parentId: 'sec_sek_nested',
                order: 2,
                refKind: 'adminRole',
                refKey: 'ITNEST'
            }
        );
        const nestedSettings = {
            ...settings,
            adminRoles: [
                ...settings.adminRoles,
                { code: 'ITNEST', name: 'IT-Hilfe', seatMode: 'multi' }
            ],
            admin: [
                ...settings.admin,
                { role: 'IT-Hilfe', name: 'Chris Nested', email: 'chris@schule.at' }
            ],
            schulOrganigramm: layout
        };
        const verwaltung = layout.nodes.find((n) => n.id === 'sec_verwaltung');
        const dels = listM365PeopleForArea(nestedSettings, verwaltung);
        expect(dels.map((p) => p.email).sort()).toEqual([
            'a@schule.at',
            'b@schule.at',
            'chris@schule.at',
            'eva@schule.at'
        ]);
        expect(dels.find((p) => p.email === 'eva@schule.at')?.source).toBe('nested');
        expect(dels.find((p) => p.email === 'chris@schule.at')?.source).toBe('nested');
    });

    it('moveSchulOrgRole nimmt Rollen-Personen in den neuen Bereich mit', () => {
        const audA = audienceSectionId('va-team-a');
        const audB = audienceSectionId('va-team-b');
        const layout = {
            version: 2,
            nodes: [
                {
                    id: audA,
                    type: 'area',
                    label: 'Team A',
                    parentId: null,
                    refKind: 'audience',
                    refKey: 'va-team-a'
                },
                {
                    id: audB,
                    type: 'area',
                    label: 'Team B',
                    parentId: null,
                    refKind: 'audience',
                    refKey: 'va-team-b'
                },
                {
                    id: 'role_QMS',
                    type: 'role',
                    label: 'QMS',
                    parentId: audA,
                    refKind: 'adminRole',
                    refKey: 'QMS'
                }
            ]
        };
        const base = {
            adminRoles: [{ code: 'QMS', name: 'QMS', seatMode: 'single', tier: 'verwaltung' }],
            admin: [{ role: 'QMS', name: 'Clara', email: 'clara@schule.at', defaultKey: 'QMS' }],
            adminAudienceMemberships: [
                { groupId: 'va-team-a', email: 'clara@schule.at', name: 'Clara' }
            ],
            verwaltungAudienceGroups: [
                { id: 'va-team-a', label: 'Team A', builtin: false, order: 1 },
                { id: 'va-team-b', label: 'Team B', builtin: false, order: 2 }
            ],
            schulOrganigramm: layout
        };
        const result = moveSchulOrgRole(base, 'role_QMS', audB);
        expect(result.ok).toBe(true);
        expect(result.layout.nodes.find((n) => n.id === 'role_QMS')?.parentId).toBe(audB);
        const clara = result.settings.admin.find((p) => p.email === 'clara@schule.at');
        expect(clara?.role).toBe('QMS');
        expect(
            (result.settings.adminAudienceMemberships || []).some(
                (m) => m.email === 'clara@schule.at' && m.groupId === 'va-team-a'
            )
        ).toBe(false);
        expect(
            (result.settings.adminAudienceMemberships || []).some(
                (m) => m.email === 'clara@schule.at' && m.groupId === 'va-team-b'
            )
        ).toBe(true);
        expect(
            listPeopleForSchulOrgNode(result.settings, { type: 'role', refKind: 'adminRole', refKey: 'QMS' }).map(
                (p) => p.email
            )
        ).toEqual(['clara@schule.at']);
    });

    it('moveSchulOrgPerson verschiebt Admin zwischen Rollen', () => {
        const layout = seedSchulOrganigrammFromStammdaten(settings);
        const sek = layout.nodes.find((n) => n.type === 'role' && n.refKey === 'SEK');
        const dir = layout.nodes.find((n) => n.type === 'role' && n.refKey === 'DIREKTION');
        const result = moveSchulOrgPerson(
            settings,
            {
                kind: 'admin',
                name: 'Anna',
                email: 'a@schule.at',
                sourceRefKind: 'adminRole',
                sourceRefKey: 'SEK'
            },
            dir
        );
        expect(result.ok).toBe(true);
        const anna = result.settings.admin.find((p) => p.email === 'a@schule.at');
        expect(anna.role).toBe('Direktion');
        expect(peopleForAdminRole(result.settings, findAdminRole(result.settings, 'SEK')).filter((p) => p.email)).toHaveLength(1);
        expect(listPeopleForSchulOrgNode(result.settings, sek).map((p) => p.name)).toEqual(['Ben']);
    });

    it('moveSchulOrgPerson ordnet Bereichsperson einer Rolle zu', () => {
        const layout = seedSchulOrganigrammFromStammdaten({
            ...settings,
            adminAudienceMemberships: [
                { groupId: 'verwaltung', email: 'eva@schule.at', name: 'Eva Serafini' }
            ],
            schulOrganigramm: undefined
        });
        const seeded = seedSchulOrganigrammFromStammdaten({
            ...settings,
            adminAudienceMemberships: [
                { groupId: 'verwaltung', email: 'eva@schule.at', name: 'Eva Serafini' }
            ]
        });
        const area = seeded.nodes.find((n) => n.id === 'sec_verwaltung');
        const role = seeded.nodes.find((n) => n.type === 'role' && n.refKey === 'SEK');
        expect(area && role).toBeTruthy();
        const before = listPeopleForSchulOrgNode(
            {
                ...settings,
                adminAudienceMemberships: [
                    { groupId: 'verwaltung', email: 'eva@schule.at', name: 'Eva Serafini' }
                ],
                schulOrganigramm: seeded
            },
            area
        );
        expect(before.some((p) => p.email === 'eva@schule.at')).toBe(true);
        const result = moveSchulOrgPerson(
            {
                ...settings,
                admin: [],
                adminAudienceMemberships: [
                    { groupId: 'verwaltung', email: 'eva@schule.at', name: 'Eva Serafini' }
                ],
                schulOrganigramm: seeded
            },
            {
                kind: 'areaPerson',
                name: 'Eva Serafini',
                email: 'eva@schule.at',
                sourceNodeId: 'sec_verwaltung',
                sourceRefKind: 'area',
                sourceRefKey: 'sec_verwaltung'
            },
            role
        );
        expect(result.ok).toBe(true);
        expect(result.message).toMatch(/Rolle/);
        const eva = result.settings.admin.find((p) => p.email === 'eva@schule.at');
        expect(eva?.role).toBe('Sekretariat');
        expect(
            (result.settings.adminAudienceMemberships || []).some(
                (m) => m.email === 'eva@schule.at' && m.groupId === 'verwaltung'
            )
        ).toBe(false);
        expect(listPeopleForSchulOrgNode(result.settings, role).map((p) => p.name)).toContain(
            'Eva Serafini'
        );
        // administration muss mitlaufen – sonst verwirft Tenant-Save die Person wieder
        const sekGroup = (result.settings.administration || []).find(
            (g) => String(g?.name || '').toLowerCase() === 'sekretariat'
        );
        expect(sekGroup?.people?.some((p) => p.email === 'eva@schule.at')).toBe(true);
        void layout;
    });

    it('moveSchulOrgPerson verschiebt Schüler zwischen Klassen', () => {
        const layout = seedSchulOrganigrammFromStammdaten(settings);
        const c2 = layout.nodes.find((n) => n.refKind === 'class' && n.refKey === '2B');
        const result = moveSchulOrgPerson(
            settings,
            {
                kind: 'student',
                name: 'Lea',
                email: 'lea@schule.at',
                sourceRefKind: 'class',
                sourceRefKey: '1A'
            },
            c2
        );
        expect(result.ok).toBe(true);
        expect(result.settings.students.find((s) => s.email === 'lea@schule.at')?.klasse).toBe('2B');
    });

    it('listPeopleForSchulOrgNode: SV-Sitze und SGA-Scope-Bereiche', () => {
        const layout = seedSchulOrganigrammFromStammdaten(settings);
        const schulsprecher = layout.nodes.find((n) => n.id === 'sv_schulsprecher');
        const sgaTeacher = layout.nodes.find((n) => n.id === 'sga_gruppe_lehrer');
        const sgaStudent = layout.nodes.find((n) => n.id === 'sga_gruppe_schueler');
        const sgaBoard = layout.nodes.find((n) => n.id === 'sga_board');
        expect(sgaStudent?.label).toBe('Schülervertretung');
        expect(listPeopleForSchulOrgNode(settings, schulsprecher).map((p) => p.name)).toEqual(['Lea']);
        expect(listPeopleForSchulOrgNode(settings, sgaStudent).map((p) => p.name)).toEqual(['Lea', 'Tom']);
        expect(listPeopleForSchulOrgNode(settings, sgaBoard)).toHaveLength(0);
        expect(listPeopleForSchulOrgNode(settings, sgaTeacher).filter((p) => p.scope === 'teacher')).toHaveLength(
            1
        );
    });

    it('Eltern-Klassen-Kachel zeigt Erziehungsberechtigte aus Schüler-parentPairs', () => {
        expect(guardiansForClass(settings, '1A').map((g) => g.email).sort()).toEqual([
            'eltern.lea@familie.at',
            'eltern.max@familie.at',
            'vater.lea@familie.at'
        ]);
        const layout = seedSchulOrganigrammFromStammdaten(settings);
        const eltern1a = layout.nodes.find((n) => n.id === elternClassNodeId('1A'));
        expect(eltern1a?.refKind).toBe('elternClass');
        expect(eltern1a?.parentId).toMatch(/sec_eltern_jahrgang_/);
        const people = listPeopleForSchulOrgNode(settings, eltern1a);
        expect(people.every((p) => p.kind === 'guardian')).toBe(true);
        expect(people.map((p) => p.email).sort()).toEqual([
            'eltern.lea@familie.at',
            'eltern.max@familie.at',
            'vater.lea@familie.at'
        ]);
    });

    it('guardiansForClass löst kanonische guardianIds→guardians[] auf', () => {
        const canon = {
            classes: [{ code: '1A', name: '1A' }],
            students: [
                {
                    id: 's1',
                    klasse: '1A',
                    name: 'Lea',
                    email: 'lea@schule.at',
                    guardianIds: ['g1', 'g2']
                },
                {
                    id: 's2',
                    klasse: '1A',
                    name: 'Max',
                    email: 'max@schule.at',
                    guardianIds: ['g3']
                }
            ],
            guardians: [
                { id: 'g1', name: 'Eltern Lea', email: 'eltern.lea@familie.at' },
                { id: 'g2', name: 'Vater Lea', email: 'vater.lea@familie.at' },
                { id: 'g3', name: 'Eltern Max', email: 'eltern.max@familie.at' }
            ]
        };
        expect(guardiansForClass(canon, '1A').map((g) => g.email).sort()).toEqual([
            'eltern.lea@familie.at',
            'eltern.max@familie.at',
            'vater.lea@familie.at'
        ]);
        const people = listPeopleForSchulOrgNode(
            canon,
            { id: elternClassNodeId('1A'), type: 'area', refKind: 'elternClass', refKey: '1A' }
        );
        expect(people.map((p) => p.email).sort()).toEqual([
            'eltern.lea@familie.at',
            'eltern.max@familie.at',
            'vater.lea@familie.at'
        ]);
    });

    it('KlassenChat-Kachel zeigt Lehrkräfte aus Belegung inkl. KV', () => {
        const members = teachersForClassChat(settings, '1A');
        expect(members.map((m) => m.email).sort()).toEqual([
            'mueller@schule.at',
            'schmidt@schule.at'
        ]);
        expect(members.find((m) => m.email === 'schmidt@schule.at')?.role).toBe('kv');
        const layout = seedSchulOrganigrammFromStammdaten(settings);
        const chat1a = layout.nodes.find((n) => n.id === klassenchatClassNodeId('1A'));
        expect(chat1a?.refKind).toBe('klassenchatClass');
        expect(chat1a?.parentId).toMatch(/sec_klassenchat_jahrgang_/);
        expect(chat1a?.meta?.chatId).toBe('chat-1a');
        expect(chat1a?.meta?.chatKind).toBe('groupChat');
        const people = listPeopleForSchulOrgNode(settings, chat1a);
        expect(people.map((p) => p.email).sort()).toEqual([
            'mueller@schule.at',
            'schmidt@schule.at'
        ]);
        expect(people.find((p) => p.email === 'schmidt@schule.at')?.kind).toBe('head');
    });

    it('KlassenChat-Kachel spiegelt privateChannel (teamId+channelId) in meta', () => {
        const withChannel = Object.assign({}, settings, {
            classChats: {
                version: 1,
                items: [
                    {
                        klasse: '1A',
                        kind: 'privateChannel',
                        channelId: 'chan-1a',
                        teamId: 'team-lehrer',
                        topic: '1A Kanal',
                        webUrl: 'https://teams.microsoft.com/l/channel/chan-1a'
                    }
                ]
            }
        });
        const layout = seedSchulOrganigrammFromStammdaten(withChannel);
        const chat1a = layout.nodes.find((n) => n.id === klassenchatClassNodeId('1A'));
        expect(chat1a?.meta?.chatKind).toBe('privateChannel');
        expect(chat1a?.meta?.channelId).toBe('chan-1a');
        expect(chat1a?.meta?.teamId).toBe('team-lehrer');
        expect(chat1a?.meta?.chatId).toBe('chan-1a');
        expect(chat1a?.meta?.chatTopic).toBe('1A Kanal');
    });

    it('moveSchulOrgPerson verschiebt ARGE-Leitung', () => {
        const layout = seedSchulOrganigrammFromStammdaten(settings);
        const deut = layout.nodes.find((n) => n.refKind === 'arge' && n.refKey === 'DEUT');
        const result = moveSchulOrgPerson(
            settings,
            {
                kind: 'head',
                name: 'Müller',
                email: 'mueller@schule.at',
                sourceRefKind: 'arge',
                sourceRefKey: 'MATHE'
            },
            deut
        );
        expect(result.ok).toBe(true);
        expect(findArge(result.settings, 'MATHE')?.headEmail).toBe('');
        expect(findArge(result.settings, 'DEUT')?.headEmail).toBe('mueller@schule.at');
    });
});
