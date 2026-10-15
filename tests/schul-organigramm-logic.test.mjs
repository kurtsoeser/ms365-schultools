import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
    seedSchulOrganigrammFromStammdaten,
    syncSchulOrganigrammWithStammdaten,
    rebuildSchulOrganigrammFromStammdaten,
    resolveSchulOrganigrammLayout,
    convertVerwaltungOrganigrammToSchul,
    migrateVerwaltungOrganigrammToSchul,
    moveSchulOrganigrammNode,
    moveSchulOrganigrammArea,
    applySchulOrganigrammToStammdaten,
    reorderSchulOrganigrammSections,
    nodeMatchesSchulOrgFilter,
    schulOrganigrammStats,
    dedupeSchulOrganigrammLayout,
    childrenOf,
    audienceSectionId,
    canPlaceSchulOrgNodeInSection,
    listSchulOrgRoleParentAreas,
    normalizeSchulOrganigramm,
    setSchulOrganigrammView,
    setSchulOrganigrammNodePosition,
    setSchulOrganigrammNodeSize,
    clampSchulOrganigrammAreaSize,
    setSchulOrganigrammAreaCollapsed,
    parseSchulOrgCanvasCoord,
    mergeSchulOrganigrammLivePositions,
    applyLandkarteRootPositionsFromBoxes,
    applyLandkartePositionsFromBoxes,
    SCHUL_ORG_AREA_MIN_W,
    SCHUL_ORG_AREA_MIN_H,
    defaultCanvasSlot,
    setSchulOrganigrammAreaM365,
    upsertSchulOrganigrammAreaPerson,
    removeSchulOrganigrammAreaPerson,
    isAreaNode,
    isSammelgruppeArea,
    canDeleteSchulOrganigrammArea,
    isBuiltinSchulOrganigrammArea,
    SCHUL_ORG_BOARD_REF,
    SCHUL_ORG_SEC_LEHRER,
    SCHUL_ORG_SEC_SCHUELER,
    SCHUL_ORG_SEC_ELTERN,
    SCHUL_ORG_SEC_KLASSENCHATS,
    SCHUL_ORG_SGA_BOARD_ID,
    elternJahrgangAreaId,
    elternClassNodeId,
    isElternClassNode,
    isElternJahrgangArea,
    ensureSchulOrganigrammElternStructure,
    klassenchatJahrgangAreaId,
    klassenchatClassNodeId,
    isKlassenchatClassNode,
    isKlassenchatJahrgangArea,
    ensureSchulOrganigrammKlassenchatsStructure,
    SCHUL_ORG_SGA_SV_AREA_ID,
    SCHUL_ORG_SGA_SCOPE_AREAS,
    SCHUL_ORG_SV_SEATS,
    ensureSchulOrganigrammSgaStructure,
    ensureSchulOrganigrammJahrgangStructure,
    ensureSchulOrganigrammArgeLeiterStructure,
    collectArgeLeiterPeopleFromSettings,
    listArgeCatalogGroupRefs,
    isArgeLeiterArea,
    isArgeAreaNode,
    isClassNode,
    SCHUL_ORG_SEC_ARGE_LEITER,
    schulstufeFromClassCode,
    canonicalSchulOrgClassCode,
    classCodeFromSettingsRow,
    jahrgangAreaId,
    isJahrgangArea,
    SCHUL_ORG_VERSION,
    NODE,
    kursteamPlacementsFromBelegung,
    kursteamStableKey,
    mergeDuplicateSchulOrganigrammClasses,
    stripPhantomSchulOrgClassAreas,
    isPhantomSchulOrgClassNode,
    normalizeSchulOrgGroupKind,
    schulOrgGroupKindLabel,
    inferSchulOrgGroupKind,
    suggestSchulOrgMailNickname
} from '../src/tools/schul-organigramm/schul-organigramm-logic.js';
import {
    computeLandkarteLayout,
    computeFocusLayout,
    isDrillableNode,
    pickLandkarteDropAreaId,
    pickDeepestAreaAtPoint,
    resolveLandkarteNestParent,
    findInsertBeforeSiblingId,
    snapLandkarteLeavesInAreas
} from '../src/tools/schul-organigramm/schul-organigramm-canvas.js';
import {
    listPeopleForSchulOrgNode,
    addSchulOrgPersonToArea,
    removeSchulOrgPersonFromArea,
    moveSchulOrgPerson
} from '../src/tools/schul-organigramm/schul-organigramm-edit.js';

const root = dirname(fileURLToPath(import.meta.url));
const projectRoot = join(root, '..');

describe('schul-organigramm-logic', () => {
    it('seedet Bereiche und Rollen/ARGEs/Klassen aus Stammdaten', () => {
        const layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [
                { code: 'DIREKTION', name: 'Direktion', tier: 'schulleitung' },
                { code: 'SEK', name: 'Sekretariat', tier: 'verwaltung' }
            ],
            admin: [{ role: 'Direktion', name: 'Holtmann', email: 'h@schule.at' }],
            teachers: [
                { name: 'Anna Lehrer', email: 'anna@schule.at' },
                { name: 'Bert Lehrer', email: 'bert@schule.at' }
            ],
            students: [{ name: 'Clara Schüler', email: 'clara@schule.at', klasse: '1A' }],
            arges: [{ code: 'MATHE', name: 'ARGE Mathe', headName: 'Anna' }],
            classes: [{ code: '1A', name: '1A', year: '2034', headName: 'KV' }],
            studentCouncil: [],
            sga: []
        });
        const roots = childrenOf(layout, null);
        expect(roots.map((r) => r.id)).toContain('sec_schulleitung');
        expect(roots.map((r) => r.id)).toContain('sec_arge');
        expect(roots.map((r) => r.id)).toContain('sec_lehrer');
        expect(roots.map((r) => r.id)).toContain('sec_schueler');
        expect(roots.map((r) => r.id)).toContain('sec_eltern');
        expect(roots.map((r) => r.id)).not.toContain('sec_klassen');
        const argeLeiter = layout.nodes.find((n) => n.id === SCHUL_ORG_SEC_ARGE_LEITER);
        expect(argeLeiter?.parentId).toBe('sec_arge');
        expect(isArgeLeiterArea(argeLeiter)).toBe(true);
        expect(argeLeiter?.refKind).toBe('argeLeiter');
        const lehrer = layout.nodes.find((n) => n.id === 'sec_lehrer');
        expect(lehrer?.refKind).toBe('sammelgruppe');
        expect(lehrer?.refKey).toBe('lehrer');
        expect(lehrer?.meta?.peopleCount).toBe(2);
        const schueler = layout.nodes.find((n) => n.id === 'sec_schueler');
        expect(schueler?.refKind).toBe('sammelgruppe');
        expect(schueler?.meta?.classCount).toBe(1);
        expect(layout.nodes.find((n) => n.id === 'sec_klassen')).toBeFalsy();
        const jg1 = layout.nodes.find((n) => n.id === jahrgangAreaId('1'));
        expect(jg1?.parentId).toBe('sec_schueler');
        expect(jg1?.refKind).toBe('jahrgang');
        expect(isJahrgangArea(jg1)).toBe(true);
        expect(layout.nodes.find((n) => isClassNode(n) && n.refKey === '1A')?.parentId).toBe(
            jahrgangAreaId('1')
        );
        expect(schulstufeFromClassCode('2BK')).toBe('2');
        expect(canonicalSchulOrgClassCode('Klasse 5HM - SJ2026/27')).toBe('5HM');
        expect(schulstufeFromClassCode('Klasse 4HMA - SJ2026/27')).toBe('4');
        expect(classCodeFromSettingsRow({ code: '', name: 'Klasse 5HM - SJ2026/27' })).toBe('5HM');
        const roles = layout.nodes.filter((n) => n.type === NODE.ROLE);
        expect(roles).toHaveLength(2);
        expect(roles.find((r) => r.refKey === 'DIREKTION')?.parentId).toBe('sec_schulleitung');
        expect(
            layout.nodes.some(
                (n) => isArgeAreaNode(n) && n.refKey === 'MATHE' && isAreaNode(n)
            )
        ).toBe(true);
    });

    it('SGA enthält Schülervertretung (3 Funktionen) / Lehrer / Eltern', () => {
        const layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            classes: [],
            studentCouncil: [
                { name: 'Anna', email: 'anna@schule.at', klasse: '5A' },
                { name: 'Ben', email: 'ben@schule.at', klasse: '4B' }
            ],
            sga: [
                { scope: 'student', name: 'S1', email: 's1@schule.at' },
                { scope: 'teacher', name: 'L1', email: 'l1@schule.at' },
                { scope: 'external', name: 'E1', email: 'e1@schule.at' }
            ]
        });
        const board = layout.nodes.find((n) => n.id === SCHUL_ORG_SGA_BOARD_ID);
        expect(board?.type).toBe('area');
        expect(board?.label).toBe('SGA');
        expect(board?.parentId == null || board?.parentId === '').toBe(true);
        expect(layout.nodes.find((n) => n.id === 'sec_vertretung')).toBeFalsy();
        expect(layout.nodes.find((n) => n.id === 'council_board')).toBeFalsy();
        const svArea = layout.nodes.find((n) => n.id === SCHUL_ORG_SGA_SV_AREA_ID);
        expect(svArea?.label).toBe('Schülervertretung');
        const scopes = SCHUL_ORG_SGA_SCOPE_AREAS.map((s) => s.id);
        scopes.forEach((id) => {
            const n = layout.nodes.find((x) => x.id === id);
            expect(n?.parentId).toBe(SCHUL_ORG_SGA_BOARD_ID);
            expect(n?.refKind).toBe('sgaScope');
        });
        SCHUL_ORG_SV_SEATS.forEach((seat) => {
            const n = layout.nodes.find((x) => x.id === seat.id);
            expect(n?.parentId).toBe(SCHUL_ORG_SGA_SV_AREA_ID);
            expect(n?.label).toBe(seat.label);
        });
        expect(layout.nodes.find((n) => n.id === 'sv_schulsprecher')?.meta?.personName).toBe('Anna');
        const map = computeLandkarteLayout(layout);
        const outer = map.boxes.find((b) => b.kind === 'area' && b.node.id === SCHUL_ORG_SGA_BOARD_ID);
        const innerIds = map.boxes
            .filter((b) => b.kind === 'area' && b.parentAreaId === SCHUL_ORG_SGA_BOARD_ID)
            .map((b) => b.node.id);
        expect(outer).toBeTruthy();
        expect(innerIds.sort()).toEqual(scopes.slice().sort());
        expect(outer.h).toBeGreaterThan(200);

        const legacy = {
            version: 2,
            nodes: [
                {
                    id: 'sec_vertretung',
                    type: 'area',
                    label: 'Schülervertretung & SGA',
                    parentId: null,
                    tone: 'council'
                },
                {
                    id: 'sga_board',
                    type: 'sga',
                    label: 'SGA',
                    parentId: 'sec_vertretung',
                    refKind: 'sga',
                    refKey: '_board'
                },
                {
                    id: 'sga_gruppe_lehrer',
                    type: 'area',
                    label: 'Lehrer',
                    parentId: 'sec_vertretung',
                    refKind: 'sgaScope',
                    refKey: 'teacher'
                }
            ]
        };
        const ensured = ensureSchulOrganigrammSgaStructure(legacy, { sga: [] });
        expect(ensured.nodes.find((n) => n.id === 'sec_vertretung')).toBeFalsy();
        const ensuredBoard = ensured.nodes.find((n) => n.id === SCHUL_ORG_SGA_BOARD_ID);
        expect(ensuredBoard?.parentId == null || ensuredBoard?.parentId === '').toBe(true);
        expect(ensuredBoard?.label).toBe('SGA');
        expect(ensured.nodes.filter((n) => String(n.id).startsWith('sga_gruppe_'))).toHaveLength(3);
        expect(ensured.nodes.find((n) => n.id === 'sga_gruppe_lehrer')?.parentId).toBe(
            SCHUL_ORG_SGA_BOARD_ID
        );
    });

    it('SGA-Board behält zugeordnetes M365-Objekt beim Sync', () => {
        const layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            classes: [],
            sga: []
        });
        const withM365 = setSchulOrganigrammAreaM365(layout, SCHUL_ORG_SGA_BOARD_ID, {
            groupMode: 'match',
            groupKind: 'distribution',
            groupId: 'gid-sga-dl',
            groupLabel: 'SGA-Verteiler',
            groupMail: 'sga@schule.at'
        });
        const synced = syncSchulOrganigrammWithStammdaten(withM365, {
            adminRoles: [],
            admin: [],
            arges: [],
            classes: [],
            sga: [{ scope: 'teacher', name: 'L1', email: 'l1@schule.at' }]
        });
        const board = synced.layout.nodes.find((n) => n.id === SCHUL_ORG_SGA_BOARD_ID);
        expect(board?.m365).toMatchObject({
            groupMode: 'match',
            groupKind: 'distribution',
            groupId: 'gid-sga-dl',
            groupLabel: 'SGA-Verteiler'
        });
    });

    it('Standard-Bereiche: Lehrer geschützt, Schulleitung/Verwaltung und Custom löschbar', () => {
        const layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            classes: [],
            verwaltungAudienceGroups: [{ id: 'va-ext', label: 'Erweiterte Schulleitung', builtin: false }]
        });
        expect(isBuiltinSchulOrganigrammArea(layout.nodes.find((n) => n.id === 'sec_schulleitung'))).toBe(
            true
        );
        expect(canDeleteSchulOrganigrammArea(layout.nodes.find((n) => n.id === SCHUL_ORG_SEC_LEHRER))).toBe(
            false
        );
        expect(canDeleteSchulOrganigrammArea(layout.nodes.find((n) => n.id === 'sec_verwaltung'))).toBe(
            true
        );
        expect(canDeleteSchulOrganigrammArea(layout.nodes.find((n) => n.id === 'sec_schulleitung'))).toBe(
            true
        );
        const custom = layout.nodes.find((n) => n.id === 'sec_aud_va-ext');
        expect(custom).toBeTruthy();
        expect(canDeleteSchulOrganigrammArea(custom)).toBe(true);
    });

    it('entfernte Verwaltung bleibt nach Sync weg', () => {
        let layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [{ code: 'SEK', name: 'Sekretariat', tier: 'verwaltung' }],
            admin: [],
            arges: [],
            classes: []
        });
        layout = {
            ...layout,
            nodes: layout.nodes
                .filter((n) => n.id !== 'sec_verwaltung')
                .map((n) =>
                    n.parentId === 'sec_verwaltung' ? { ...n, parentId: null } : n
                ),
            removedAreaIds: ['sec_verwaltung']
        };
        const synced = syncSchulOrganigrammWithStammdaten(layout, {
            adminRoles: [{ code: 'SEK', name: 'Sekretariat', tier: 'verwaltung' }],
            admin: [],
            arges: [],
            classes: []
        });
        expect(synced.layout.nodes.some((n) => n.id === 'sec_verwaltung')).toBe(false);
        expect(synced.layout.removedAreaIds).toContain('sec_verwaltung');
    });

    it('sync ergänzt fehlende Lehrer-/Schüler-Standardbereiche', () => {
        const seeded = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            teachers: [{ name: 'T', email: 't@schule.at' }],
            students: [{ name: 'S', email: 's@schule.at' }],
            arges: [],
            classes: []
        });
        const without = {
            version: 2,
            nodes: seeded.nodes.filter((n) => n.id !== 'sec_lehrer' && n.id !== 'sec_schueler')
        };
        const synced = syncSchulOrganigrammWithStammdaten(without, {
            adminRoles: [],
            admin: [],
            teachers: [{ name: 'T', email: 't@schule.at' }],
            students: [{ name: 'S', email: 's@schule.at' }],
            arges: [],
            classes: []
        });
        expect(synced.added).toBeGreaterThanOrEqual(2);
        expect(synced.layout.nodes.some((n) => n.id === 'sec_lehrer')).toBe(true);
        expect(synced.layout.nodes.some((n) => n.id === 'sec_schueler')).toBe(true);
    });

    it('sync ergänzt fehlende Rollen und aktualisiert Labels', () => {
        const seeded = seedSchulOrganigrammFromStammdaten({
            adminRoles: [{ code: 'SEK', name: 'Sekretariat', tier: 'verwaltung' }],
            admin: [],
            arges: [],
            classes: []
        });
        const synced = syncSchulOrganigrammWithStammdaten(seeded, {
            adminRoles: [
                { code: 'SEK', name: 'Sekretariat GS', tier: 'verwaltung' },
                { code: 'IT', name: 'IT-Support', tier: 'verwaltung' }
            ],
            admin: [],
            arges: [],
            classes: []
        });
        expect(synced.added).toBe(1);
        expect(synced.labelsUpdated).toBe(1);
        const codes = synced.layout.nodes
            .filter((n) => n.type === NODE.ROLE)
            .map((n) => n.refKey)
            .sort();
        expect(codes).toEqual(['IT', 'SEK']);
        expect(synced.layout.nodes.find((n) => n.refKey === 'SEK')?.label).toBe('Sekretariat GS');
    });

    it('resolve seedet leer und markiert dirty', () => {
        const out = resolveSchulOrganigrammLayout(null, {
            adminRoles: [{ code: 'SEK', name: 'Sekretariat', tier: 'verwaltung' }],
            admin: [],
            arges: [],
            classes: []
        });
        expect(out.seeded).toBe(true);
        expect(out.dirty).toBe(true);
        expect(out.layout.nodes.some((n) => n.type === NODE.ROLE)).toBe(true);
    });

    it('resolve mit bestehendem Layout ist nicht dirty ohne Änderungen', () => {
        const seeded = seedSchulOrganigrammFromStammdaten({
            adminRoles: [{ code: 'SEK', name: 'Sekretariat', tier: 'verwaltung' }],
            admin: [],
            arges: [],
            classes: []
        });
        const out = resolveSchulOrganigrammLayout(seeded, {
            adminRoles: [{ code: 'SEK', name: 'Sekretariat', tier: 'verwaltung' }],
            admin: [],
            arges: [],
            classes: []
        });
        expect(out.seeded).toBe(false);
        expect(out.dirty).toBe(false);
        expect(out.migratedFromVerwaltung).toBe(false);
    });

    it('convert mappt Verwaltung group/role/sammel auf Schul-Knoten', () => {
        const { layout, stats } = convertVerwaltungOrganigrammToSchul({
            version: 1,
            nodes: [
                { id: 'g1', type: 'group', label: 'Schulleitung', parentId: null, order: 0, tone: 'lead' },
                { id: 'g2', type: 'group', label: 'Verwaltung (Personal)', parentId: null, order: 1, tone: 'staff' },
                {
                    id: 'r1',
                    type: 'role',
                    label: 'Direktion',
                    parentId: 'g1',
                    order: 2,
                    roleCode: 'DIREKTION',
                    tone: 'lead'
                },
                {
                    id: 's1',
                    type: 'sammel',
                    label: 'Alle Verwaltung',
                    parentId: 'g2',
                    order: 3,
                    audienceGroupId: 'aud1',
                    m365GroupId: 'gid-1',
                    m365GroupMail: 'vw@schule.at'
                }
            ]
        });
        expect(stats.roles).toBe(1);
        expect(stats.sammel).toBe(1);
        expect(layout.nodes.find((n) => n.id === 'sec_schulleitung')?.type).toBe(NODE.SECTION);
        expect(layout.nodes.find((n) => n.refKey === 'DIREKTION')?.parentId).toBe('sec_schulleitung');
        const sammel = layout.nodes.find((n) => n.tone === 'sammel');
        expect(sammel?.meta?.m365GroupId).toBe('gid-1');
        expect(sammel?.refKind).toBe('audience');
    });

    it('resolve migriert aus verwaltungOrganigramm wenn Schul-Layout leer', () => {
        const out = resolveSchulOrganigrammLayout(null, {
            adminRoles: [
                { code: 'DIREKTION', name: 'Direktion', tier: 'schulleitung' },
                { code: 'SEK', name: 'Sekretariat', tier: 'verwaltung' }
            ],
            admin: [],
            arges: [{ code: 'MATHE', name: 'Mathe' }],
            classes: [],
            verwaltungOrganigramm: {
                version: 1,
                nodes: [
                    { id: 'g1', type: 'group', label: 'Schulleitung', parentId: null, order: 0, tone: 'lead' },
                    { id: 'g2', type: 'group', label: 'Verwaltung', parentId: null, order: 1, tone: 'staff' },
                    {
                        id: 'r1',
                        type: 'role',
                        label: 'Direktion',
                        parentId: 'g1',
                        order: 2,
                        roleCode: 'DIREKTION'
                    }
                ]
            }
        });
        expect(out.migratedFromVerwaltung).toBe(true);
        expect(out.dirty).toBe(true);
        expect(out.layout.nodes.some((n) => n.type === NODE.ROLE && n.refKey === 'DIREKTION')).toBe(true);
        expect(out.layout.nodes.some((n) => n.type === NODE.ROLE && n.refKey === 'SEK')).toBe(true);
        expect(out.layout.nodes.some((n) => isArgeAreaNode(n) && n.refKey === 'MATHE')).toBe(true);
    });

    it('migrateVerwaltungOrganigrammToSchul gibt null ohne Legacy', () => {
        expect(migrateVerwaltungOrganigrammToSchul({ adminRoles: [] })).toBeNull();
    });

    it('moveSchulOrganigrammNode ordnet und verschiebt Rollen zwischen Bereichen', () => {
        const layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [
                { code: 'DIREKTION', name: 'Direktion', tier: 'schulleitung' },
                { code: 'SEK', name: 'Sekretariat', tier: 'verwaltung' },
                { code: 'IT', name: 'IT', tier: 'verwaltung' }
            ],
            admin: [],
            arges: [],
            classes: []
        });
        const sektion = childrenOf(layout, 'sec_verwaltung').map((n) => n.refKey);
        expect(sektion).toContain('SEK');
        expect(sektion).toContain('IT');
        const moved = moveSchulOrganigrammNode(layout, 'role_IT', 'sec_schulleitung', 'role_DIREKTION');
        expect(moved.nodes.find((n) => n.refKey === 'IT')?.parentId).toBe('sec_schulleitung');
        const lead = childrenOf(moved, 'sec_schulleitung').map((n) => n.refKey);
        expect(lead.indexOf('IT')).toBeLessThan(lead.indexOf('DIREKTION'));
        const blocked = moveSchulOrganigrammNode(layout, 'role_SEK', 'sec_arge', '');
        expect(blocked.nodes.find((n) => n.refKey === 'SEK')?.parentId).toBe('sec_verwaltung');
    });

    it('Rollen im gleichen Bereich nach Drop-Y neu einreihen und dicht packen', () => {
        let layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [
                { code: 'A', name: 'Rolle A', tier: 'verwaltung' },
                { code: 'B', name: 'Rolle B', tier: 'verwaltung' },
                { code: 'C', name: 'Rolle C', tier: 'verwaltung' }
            ],
            admin: [],
            arges: [],
            classes: []
        });
        // Freie Koordinaten (wie nach fehlerhaftem Drag) → müssen verschwinden
        layout = setSchulOrganigrammNodePosition(layout, 'role_C', 12, 999);
        const map0 = computeLandkarteLayout(layout);
        const leaves0 = map0.boxes
            .filter((b) => b.kind === 'leaf' && b.parentAreaId === 'sec_verwaltung')
            .sort((a, b) => a.y - b.y);
        expect(leaves0.length).toBeGreaterThanOrEqual(3);
        // C unter die Mitte der ersten Rolle ziehen → vor A einfügen
        const before = findInsertBeforeSiblingId(
            map0.boxes,
            'sec_verwaltung',
            'role_C',
            leaves0[0].y + 2,
            'leaf'
        );
        expect(before).toBe(leaves0[0].node.id);
        layout = moveSchulOrganigrammNode(layout, 'role_C', 'sec_verwaltung', before);
        const nodeC = layout.nodes.find((n) => n.id === 'role_C');
        expect(nodeC?.x).toBeUndefined();
        expect(nodeC?.y).toBeUndefined();
        const order = childrenOf(layout, 'sec_verwaltung').map((n) => n.id);
        expect(order[0]).toBe('role_C');
        const map1 = computeLandkarteLayout(layout);
        const leaves1 = map1.boxes
            .filter((b) => b.kind === 'leaf' && b.parentAreaId === 'sec_verwaltung')
            .sort((a, b) => a.y - b.y);
        for (let i = 1; i < leaves1.length; i++) {
            expect(leaves1[i].y).toBeGreaterThanOrEqual(leaves1[i - 1].y + leaves1[i - 1].h - 0.5);
        }
        // Gleichmäßiger Abstand (kein großer Leerraum)
        if (leaves1.length >= 2) {
            const gap = leaves1[1].y - (leaves1[0].y + leaves1[0].h);
            expect(gap).toBeLessThanOrEqual(20);
        }
        // Spalten-Snap: alle Rollen gleiche X-Position im Bereich
        const parent = map1.boxes.find((b) => b.kind === 'area' && b.node.id === 'sec_verwaltung');
        expect(parent).toBeTruthy();
        const xs = new Set(leaves1.map((l) => Math.round(l.x)));
        expect(xs.size).toBe(1);
        expect(leaves1[0].x).toBeGreaterThanOrEqual(parent.x + 10);
        // absichtlich schräge Boxen → Snap korrigiert
        const skewed = map1.boxes.map((b) => Object.assign({}, b));
        skewed.forEach((b) => {
            if (b.kind === 'leaf' && b.parentAreaId === 'sec_verwaltung') {
                b.x = parent.x - 40;
            }
        });
        snapLandkarteLeavesInAreas(skewed);
        skewed
            .filter((b) => b.kind === 'leaf' && b.parentAreaId === 'sec_verwaltung')
            .forEach((l) => {
                expect(Math.round(l.x)).toBe(Math.round(parent.x + 12));
            });
    });

    it('reorderSchulOrganigrammSections ändert die Bereichsreihenfolge', () => {
        const layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            classes: []
        });
        const next = reorderSchulOrganigrammSections(layout, 'sec_arge', 'sec_schulleitung');
        const roots = childrenOf(next, null).map((n) => n.id);
        expect(roots.indexOf('sec_arge')).toBeLessThan(roots.indexOf('sec_schulleitung'));
    });

    it('nodeMatchesSchulOrgFilter und Stats content', () => {
        expect(nodeMatchesSchulOrgFilter({ label: 'Sekretariat', refKey: 'SEK' }, 'sek')).toBe(true);
        expect(nodeMatchesSchulOrgFilter({ label: 'IT', refKey: 'IT' }, 'mathe')).toBe(false);
        const layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [{ code: 'SEK', name: 'Sekretariat', tier: 'verwaltung' }],
            admin: [],
            arges: [{ code: 'MATHE', name: 'Mathe' }],
            classes: [{ code: '1A', name: '1A' }]
        });
        const st = schulOrganigrammStats(layout);
        expect(st.content).toBe(st.roles + st.arges + st.classes + st.other);
        expect(st.total).toBe(st.sections + st.content);
    });

    it('dedupe entfernt vervielfachte Rollen nach verlorenem refKey (Sync-Schleife)', () => {
        const bloated = {
            version: 1,
            nodes: [
                { id: 'sec_schulleitung', type: 'group', label: 'Schulleitung', parentId: null, order: 0, tone: 'lead' },
                { id: 'sec_verwaltung', type: 'group', label: 'Verwaltung', parentId: null, order: 1, tone: 'staff' },
                {
                    id: 'role_SL',
                    type: 'role',
                    label: 'Direktion',
                    parentId: 'sec_schulleitung',
                    order: 2,
                    roleCode: 'SL',
                    refKind: 'adminRole',
                    refKey: 'SL'
                },
                { id: 'og_a', type: 'group', label: 'Direktion', parentId: 'sec_schulleitung', order: 3, roleCode: 'SL' },
                { id: 'og_b', type: 'group', label: 'Direktion', parentId: 'sec_schulleitung', order: 4, roleCode: 'SL' },
                { id: 'OG_KLJ43FJC1A', type: 'role', label: 'Direktion', parentId: 'sec_schulleitung', order: 5, refKind: 'adminRole', refKey: 'OG_KLJ43FJC1A' },
                { id: 'OG_K82R0VJK8I', type: 'role', label: 'Direktion', parentId: 'sec_schulleitung', order: 6, refKind: 'adminRole', refKey: 'OG_K82R0VJK8I' },
                {
                    id: 'role_SEK',
                    type: 'role',
                    label: 'Sekretariat',
                    parentId: 'sec_verwaltung',
                    order: 7,
                    refKind: 'adminRole',
                    refKey: 'SEK'
                },
                { id: 'OG_0ACELTJC1A', type: 'role', label: 'Sekretariat', parentId: 'sec_verwaltung', order: 8, refKey: 'OG_0ACELTJC1A' }
            ]
        };
        const settings = {
            adminRoles: [
                { code: 'SL', name: 'Direktion', tier: 'schulleitung' },
                { code: 'SEK', name: 'Sekretariat', tier: 'verwaltung' }
            ],
            admin: [],
            arges: [],
            classes: []
        };
        const { layout, removed } = dedupeSchulOrganigrammLayout(bloated);
        expect(removed).toBeGreaterThan(0);
        expect(layout.nodes.filter((n) => n.refKey && String(n.refKey).startsWith('OG_'))).toHaveLength(0);
        const synced = syncSchulOrganigrammWithStammdaten(bloated, settings);
        expect(synced.healed).toBe(true);
        expect(synced.layout.nodes.filter((n) => n.type === NODE.ROLE).map((r) => r.refKey).sort()).toEqual([
            'SEK',
            'SL'
        ]);
        expect(synced.layout.nodes.filter((n) => n.label === 'Direktion')).toHaveLength(1);
    });

    it('seedet Custom-Verwaltungsgruppen als Organigramm-Bereiche', () => {
        const layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            classes: [],
            verwaltungAudienceGroups: [
                { id: 'schulleitung', label: 'Schulleitung', builtin: true, order: 0 },
                { id: 'verwaltung', label: 'Verwaltung (Personal)', builtin: true, order: 1 },
                { id: 'va-erw', label: 'Erweiterte Schulleitung', builtin: false, order: 10 },
                { id: 'va-warte', label: 'Schulwarte', builtin: false, order: 20 }
            ]
        });
        const erwId = audienceSectionId('va-erw');
        const warteId = audienceSectionId('va-warte');
        const erw = layout.nodes.find((n) => n.id === erwId);
        const warte = layout.nodes.find((n) => n.id === warteId);
        expect(erw?.type).toBe(NODE.SECTION);
        expect(erw?.refKind).toBe('audience');
        expect(erw?.refKey).toBe('va-erw');
        expect(erw?.label).toBe('Erweiterte Schulleitung');
        expect(warte?.label).toBe('Schulwarte');
        const lead = layout.nodes.find((n) => n.id === 'sec_schulleitung');
        expect(lead?.refKind).toBe('audience');
        expect(lead?.refKey).toBe('schulleitung');
    });

    it('sync behält umbenannte Builtin-Bereiche (z. B. ARGEs)', () => {
        let layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [{ code: 'MAT', name: 'Mathematik' }],
            classes: []
        });
        const argeSec = layout.nodes.find((n) => n.id === 'sec_arge');
        expect(argeSec).toBeTruthy();
        argeSec.label = 'Fachschaften & Teams';
        const synced = syncSchulOrganigrammWithStammdaten(layout, {
            adminRoles: [],
            admin: [],
            arges: [{ code: 'MAT', name: 'Mathematik' }],
            classes: []
        });
        expect(synced.layout.nodes.find((n) => n.id === 'sec_arge')?.label).toBe(
            'Fachschaften & Teams'
        );
        const resolved = resolveSchulOrganigrammLayout(synced.layout, {
            adminRoles: [],
            admin: [],
            arges: [{ code: 'MAT', name: 'Mathematik' }],
            classes: []
        });
        expect(resolved.layout.nodes.find((n) => n.id === 'sec_arge')?.label).toBe(
            'Fachschaften & Teams'
        );
    });

    it('sync behält umbenannten KlassenChats-Bereich', () => {
        const settings = {
            adminRoles: [],
            admin: [],
            arges: [],
            classes: [{ code: '1A', name: 'DEMO Klasse 1A' }]
        };
        let layout = seedSchulOrganigrammFromStammdaten(settings);
        const root = layout.nodes.find((n) => n.id === SCHUL_ORG_SEC_KLASSENCHATS);
        expect(root).toBeTruthy();
        root.label = 'Lehrer-Chats';
        const synced = syncSchulOrganigrammWithStammdaten(layout, settings);
        expect(synced.layout.nodes.find((n) => n.id === SCHUL_ORG_SEC_KLASSENCHATS)?.label).toBe(
            'Lehrer-Chats'
        );
        const resolved = resolveSchulOrganigrammLayout(synced.layout, settings);
        expect(resolved.layout.nodes.find((n) => n.id === SCHUL_ORG_SEC_KLASSENCHATS)?.label).toBe(
            'Lehrer-Chats'
        );
        const rebuilt = rebuildSchulOrganigrammFromStammdaten(synced.layout, settings);
        expect(rebuilt.nodes.find((n) => n.id === SCHUL_ORG_SEC_KLASSENCHATS)?.label).toBe(
            'Lehrer-Chats'
        );
    });

    it('sync ergänzt fehlende Audience-Sections und aktualisiert Labels', () => {
        const base = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            classes: [],
            verwaltungAudienceGroups: []
        });
        const synced = syncSchulOrganigrammWithStammdaten(base, {
            adminRoles: [],
            admin: [],
            arges: [],
            classes: [],
            verwaltungAudienceGroups: [
                { id: 'va-erw', label: 'Erweiterte Schulleitung', builtin: false, order: 10 }
            ]
        });
        expect(synced.added).toBeGreaterThanOrEqual(1);
        const sec = synced.layout.nodes.find((n) => n.id === audienceSectionId('va-erw'));
        expect(sec?.label).toBe('Erweiterte Schulleitung');
        const renamed = syncSchulOrganigrammWithStammdaten(synced.layout, {
            adminRoles: [],
            admin: [],
            arges: [],
            classes: [],
            verwaltungAudienceGroups: [
                { id: 'va-erw', label: 'Erw. Schulleitung', builtin: false, order: 10 }
            ]
        });
        expect(renamed.labelsUpdated).toBeGreaterThanOrEqual(1);
        expect(
            renamed.layout.nodes.find((n) => n.id === audienceSectionId('va-erw'))?.label
        ).toBe('Erw. Schulleitung');
    });

    it('erlaubt Rollen in Custom-Audience-Sections', () => {
        const role = { type: NODE.ROLE, id: 'role_X', refKind: 'adminRole', refKey: 'X' };
        expect(canPlaceSchulOrgNodeInSection(role, audienceSectionId('va-erw'))).toBe(true);
        expect(canPlaceSchulOrgNodeInSection(role, 'sec_verwaltung')).toBe(true);
        expect(canPlaceSchulOrgNodeInSection(role, 'sec_arge')).toBe(false);
        // Custom-Area ohne Audience-Ref: trotzdem Rollen ablegen
        const customLayout = {
            version: 2,
            nodes: [
                { id: 'sec_verwaltung', type: 'area', parentId: null, label: 'Verwaltung' },
                { id: 'area_custom', type: 'area', label: 'Extra', parentId: null, refKind: 'none' },
                { id: 'role_X', type: 'role', refKind: 'adminRole', refKey: 'X', parentId: 'sec_verwaltung' }
            ]
        };
        expect(
            canPlaceSchulOrgNodeInSection(
                { type: 'role', refKind: 'adminRole' },
                'area_custom',
                customLayout
            )
        ).toBe(true);
        const movedRole = moveSchulOrganigrammNode(customLayout, 'role_X', 'area_custom');
        expect(movedRole.nodes.find((n) => n.id === 'role_X')?.parentId).toBe('area_custom');
        const parents = listSchulOrgRoleParentAreas(customLayout);
        expect(parents.map((a) => a.id).sort()).toEqual(['area_custom', 'sec_verwaltung'].sort());
        expect(parents.some((a) => a.id === 'sec_arge')).toBe(false);
    });

    it('zeigt Personen-Chips in Bereichen mit direkten Personen (nicht nur Lehrkräfte)', () => {
        const areaId = audienceSectionId('va-it');
        const layout = normalizeSchulOrganigramm({
            version: 2,
            nodes: [
                {
                    id: areaId,
                    type: 'area',
                    label: 'IT-Team',
                    parentId: null,
                    order: 0,
                    tone: 'staff',
                    refKind: 'audience',
                    refKey: 'va-it',
                    x: 40,
                    y: 40
                },
                {
                    id: 'role_PL',
                    type: 'role',
                    label: 'Projektleitung MS365',
                    parentId: areaId,
                    order: 0,
                    refKind: 'adminRole',
                    refKey: 'PL'
                }
            ]
        });
        const counts = { [areaId]: 2 };
        const map = computeLandkarteLayout(layout, {
            peopleCountForArea: function (area) {
                return counts[area.id] || 0;
            }
        });
        const box = map.boxes.find((b) => b.kind === 'area' && b.node.id === areaId);
        expect(box?.peoplePreview).toBe(true);
        expect(box?.previewCount).toBe(2);
        expect(box?.h).toBeGreaterThan(120);
        const noPeople = computeLandkarteLayout(layout, {
            peopleCountForArea: function () {
                return 0;
            }
        });
        expect(
            noPeople.boxes.find((b) => b.kind === 'area' && b.node.id === areaId)?.peoplePreview
        ).toBe(false);
    });

    it('seedet KlassenChats → Jahrgang → Klasse parallel zu Schüler', () => {
        const layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            teachers: [{ code: 'MUEL', name: 'Müller', email: 'mueller@schule.at' }],
            classes: [
                { code: '1HMA', name: 'Klasse 1HMA', year: '2031', headEmail: 'schmidt@schule.at' },
                { code: '2HMB', name: 'Klasse 2HMB', year: '2030' },
                { code: 'Meisterklasse', name: 'Meisterklasse', year: '' }
            ],
            students: [],
            unterrichtsbelegung: {
                rows: [
                    {
                        klasse: '1HMA',
                        fach: 'MAT',
                        lehrerEmail: 'mueller@schule.at',
                        lehrerCode: 'MUEL'
                    }
                ]
            },
            classChats: {
                items: [
                    {
                        klasse: '1HMA',
                        chatId: 'chat-1hma',
                        topic: 'SJ26 | 1HMA | Lehrer',
                        memberEmails: ['mueller@schule.at', 'schmidt@schule.at']
                    }
                ]
            }
        });
        const root = layout.nodes.find((n) => n.id === SCHUL_ORG_SEC_KLASSENCHATS);
        expect(root?.refKind).toBe('sammelgruppe');
        expect(root?.refKey).toBe('klassenchats');
        expect(isKlassenchatJahrgangArea(layout.nodes.find((n) => n.id === klassenchatJahrgangAreaId('1')))).toBe(
            true
        );
        const kc = layout.nodes.find((n) => n.id === klassenchatClassNodeId('1HMA'));
        expect(isKlassenchatClassNode(kc)).toBe(true);
        expect(kc?.parentId).toBe(klassenchatJahrgangAreaId('1'));
        expect(kc?.meta?.chatId).toBe('chat-1hma');
        expect(layout.nodes.find((n) => n.id === klassenchatClassNodeId('MEISTERKLASSE'))?.parentId).toBe(
            klassenchatJahrgangAreaId('sonst')
        );

        const synced = syncSchulOrganigrammWithStammdaten(
            { version: 2, nodes: layout.nodes.filter((n) => n.id !== SCHUL_ORG_SEC_KLASSENCHATS) },
            {
                adminRoles: [],
                admin: [],
                arges: [],
                classes: [{ code: '1HMA', name: 'Klasse 1HMA' }],
                students: []
            }
        );
        expect(synced.layout.nodes.some((n) => n.id === SCHUL_ORG_SEC_KLASSENCHATS)).toBe(true);
        expect(
            ensureSchulOrganigrammKlassenchatsStructure(synced.layout, {
                classes: [{ code: '1HMA' }]
            }).nodes.some((n) => isKlassenchatClassNode(n))
        ).toBe(true);
    });

    it('seedet Eltern → Jahrgang → Klasse parallel zu Schüler', () => {
        const layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            teachers: [],
            classes: [
                { code: '1HMA', name: 'Klasse 1HMA - SJ2026/27', year: '2031' },
                { code: '2HMB', name: 'Klasse 2HMB - SJ2026/27', year: '2030' },
                { code: 'Meisterklasse', name: 'Meisterklasse', year: '' }
            ],
            students: [
                {
                    klasse: '1HMA',
                    name: 'Kind',
                    email: 'kind@schule.at',
                    parentPairs: [{ name: 'Eltern', email: 'e@familie.at' }]
                }
            ]
        });
        const eltern = layout.nodes.find((n) => n.id === SCHUL_ORG_SEC_ELTERN);
        expect(eltern?.refKind).toBe('sammelgruppe');
        expect(eltern?.refKey).toBe('eltern');
        expect(isElternJahrgangArea(layout.nodes.find((n) => n.id === elternJahrgangAreaId('1')))).toBe(
            true
        );
        expect(layout.nodes.find((n) => n.id === elternJahrgangAreaId('1'))?.parentId).toBe(
            SCHUL_ORG_SEC_ELTERN
        );
        const ec = layout.nodes.find((n) => n.id === elternClassNodeId('1HMA'));
        expect(isElternClassNode(ec)).toBe(true);
        expect(ec?.parentId).toBe(elternJahrgangAreaId('1'));
        expect(ec?.refKey).toBe('1HMA');
        expect(layout.nodes.find((n) => n.id === elternClassNodeId('MEISTERKLASSE'))?.parentId).toBe(
            elternJahrgangAreaId('sonst')
        );

        const synced = syncSchulOrganigrammWithStammdaten(
            { version: 2, nodes: layout.nodes.filter((n) => n.id !== SCHUL_ORG_SEC_ELTERN) },
            {
                adminRoles: [],
                admin: [],
                arges: [],
                classes: [{ code: '1HMA', name: 'Klasse 1HMA' }],
                students: []
            }
        );
        expect(synced.layout.nodes.some((n) => n.id === SCHUL_ORG_SEC_ELTERN)).toBe(true);
        expect(ensureSchulOrganigrammElternStructure(synced.layout, { classes: [{ code: '1HMA' }] }).nodes.some((n) =>
            isElternClassNode(n)
        )).toBe(true);
    });

    it('SGA Lehrer/Eltern zeigen Personen-Vorschau, Schülervertretung nicht (Sitze als Karten)', () => {
        const layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            classes: [],
            teachers: [],
            students: [],
            studentCouncil: [
                { name: 'Anna', email: 'anna@schule.at', klasse: '5A' },
                { name: 'Ben', email: 'ben@schule.at', klasse: '4B' }
            ],
            sga: [
                { scope: 'teacher', name: 'Lehrer Eins', email: 'l1@schule.at' },
                { scope: 'teacher', name: 'Lehrer Zwei', email: 'l2@schule.at' },
                { scope: 'teacher', name: 'Lehrer Drei', email: 'l3@schule.at' },
                { scope: 'external', name: 'Eltern Eins', email: 'e1@schule.at' },
                { scope: 'external', name: 'Eltern Zwei', email: 'e2@schule.at' }
            ]
        });
        const counts = {
            sga_gruppe_lehrer: 3,
            sga_gruppe_eltern: 2,
            sga_gruppe_schueler: 2
        };
        const map = computeLandkarteLayout(layout, {
            peopleCountForArea: function (area) {
                return counts[area && area.id] || 0;
            }
        });
        const lehrer = map.boxes.find((b) => b.kind === 'area' && b.node.id === 'sga_gruppe_lehrer');
        const eltern = map.boxes.find((b) => b.kind === 'area' && b.node.id === 'sga_gruppe_eltern');
        const sv = map.boxes.find((b) => b.kind === 'area' && b.node.id === 'sga_gruppe_schueler');
        expect(lehrer?.peoplePreview).toBe(true);
        expect(lehrer?.previewCount).toBe(3);
        expect(eltern?.peoplePreview).toBe(true);
        expect(eltern?.previewCount).toBe(2);
        // SV: drei Funktions-Karten, keine Personen-Chip-Leiste
        expect(sv?.peoplePreview).toBe(false);
        const svLeaves = map.boxes.filter(
            (b) => b.kind === 'leaf' && b.parentAreaId === 'sga_gruppe_schueler'
        );
        expect(svLeaves.length).toBeGreaterThanOrEqual(2);
    });

    it('Eltern-Klassen-Kacheln erlauben Personen-Vorschau', () => {
        const layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            classes: [{ code: '1A', name: '1A' }],
            students: []
        });
        const eid = elternClassNodeId('1A');
        const map = computeLandkarteLayout(layout, {
            peopleCountForArea: function (area) {
                return area && area.id === eid ? 4 : 0;
            }
        });
        const box = map.boxes.find((b) => b.kind === 'area' && b.node.id === eid);
        expect(box?.peoplePreview).toBe(true);
        expect(box?.previewCount).toBe(4);
    });

    it('KlassenChat-Klassen-Kacheln erlauben Personen-Vorschau', () => {
        const layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            classes: [{ code: '1A', name: '1A' }],
            students: []
        });
        const kid = klassenchatClassNodeId('1A');
        const map = computeLandkarteLayout(layout, {
            peopleCountForArea: function (area) {
                return area && area.id === kid ? 3 : 0;
            }
        });
        const box = map.boxes.find((b) => b.kind === 'area' && b.node.id === kid);
        expect(box?.peoplePreview).toBe(true);
        expect(box?.previewCount).toBe(3);
    });

    it('Klassen-Bereich: Personen-Vorschau ohne has-leaves trotz ausgeblendeter Kursteams', () => {
        const layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            teachers: [{ code: 'PAGE', name: 'Jimmy Page', email: 'page@schule.at' }],
            classes: [{ code: '4A', name: 'DEMO Klasse 4A', year: '2026', headEmail: 'page@schule.at' }],
            students: Array.from({ length: 8 }, (_, i) => ({
                code: 'S' + (i + 1),
                name: 'Schüler' + (i + 1),
                email: 's' + (i + 1) + '@schule.at',
                classCode: '4A'
            })),
            unterrichtsbelegung: {
                rows: [
                    {
                        klasse: '4A',
                        fach: 'MAT',
                        lehrerEmail: 'page@schule.at',
                        lehrerCode: 'PAGE'
                    }
                ]
            }
        });
        const classNode = layout.nodes.find((n) => isClassNode(n) && n.refKey === '4A');
        expect(classNode).toBeTruthy();
        const kurs = layout.nodes.filter(
            (n) =>
                String(n.type || '').toLowerCase() === 'kursteam' &&
                n.parentId === classNode.id
        );
        expect(kurs.length).toBeGreaterThanOrEqual(1);
        const map = computeLandkarteLayout(layout, {
            peopleCountForArea: function (area) {
                return area && area.id === classNode.id ? 9 : 0;
            }
        });
        const box = map.boxes.find((b) => b.kind === 'area' && b.node.id === classNode.id);
        expect(box?.peoplePreview).toBe(true);
        expect(box?.previewCount).toBe(9);
        // Kursteams zählen in contentCount, sind in der Übersicht aber nicht sichtbar
        expect(box?.contentCount).toBeGreaterThanOrEqual(1);
        expect(box?.visibleInnerCount).toBe(0);
        // 8 Chips + „+1“-Zeile brauchen ausreichend Höhe unter dem 74px-Kopf
        expect(box?.previewH).toBeGreaterThanOrEqual(9 * 36 + 16);
        expect(box?.h).toBeGreaterThanOrEqual(74 + (box?.previewH || 0));
    });

    it('klappt Bereiche auf der Landkarte ein und speichert den Zustand', () => {
        let layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [
                { code: 'A1', name: 'ARGE Eins' },
                { code: 'A2', name: 'ARGE Zwei' },
                { code: 'A3', name: 'ARGE Drei' }
            ],
            classes: []
        });
        const open = computeLandkarteLayout(layout);
        const openBox = open.boxes.find((b) => b.kind === 'area' && b.node.id === 'sec_arge');
        const openArgeAreas = open.boxes.filter(
            (b) =>
                b.kind === 'area' &&
                b.parentAreaId === 'sec_arge' &&
                isArgeAreaNode(b.node)
        );
        expect(openArgeAreas.length).toBeGreaterThanOrEqual(3);
        expect(openBox.h).toBeGreaterThan(100);

        layout = setSchulOrganigrammAreaCollapsed(layout, 'sec_arge', true);
        expect(layout.nodes.find((n) => n.id === 'sec_arge')?.collapsed).toBe(true);
        const closed = computeLandkarteLayout(layout);
        const closedBox = closed.boxes.find((b) => b.kind === 'area' && b.node.id === 'sec_arge');
        expect(closedBox?.collapsed).toBe(true);
        expect(closedBox?.h).toBeLessThanOrEqual(80);
        expect(
            closed.boxes.filter((b) => b.parentAreaId === 'sec_arge')
        ).toHaveLength(0);

        layout = setSchulOrganigrammAreaCollapsed(layout, 'sec_arge'); // toggle
        expect(layout.nodes.find((n) => n.id === 'sec_arge')?.collapsed).toBeFalsy();
        const again = computeLandkarteLayout(layout);
        expect(
            again.boxes.filter(
                (b) =>
                    b.kind === 'area' &&
                    b.parentAreaId === 'sec_arge' &&
                    isArgeAreaNode(b.node)
            ).length
        ).toBeGreaterThanOrEqual(3);

        // Persistenz über Sync
        layout = setSchulOrganigrammAreaCollapsed(layout, 'sec_arge', true);
        const synced = syncSchulOrganigrammWithStammdaten(layout, {
            adminRoles: [],
            admin: [],
            arges: [
                { code: 'A1', name: 'ARGE Eins' },
                { code: 'A2', name: 'ARGE Zwei' },
                { code: 'A3', name: 'ARGE Drei' }
            ],
            classes: []
        });
        expect(synced.layout.nodes.find((n) => n.id === 'sec_arge')?.collapsed).toBe(true);
    });

    it('erlaubt ARGEs in Nested-Bereichen unter ARGEs / Fachschaften', () => {
        let layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [
                { code: 'GSDE', name: 'ARGE FS GS Deutsch/DaF' },
                { code: 'GYDE', name: 'ARGE FS Gym Deutsch' }
            ],
            classes: []
        });
        // Nested Bereiche wie Fachleitungen GS / Gym
        layout = normalizeSchulOrganigramm({
            version: layout.version,
            view: layout.view,
            nodes: layout.nodes.concat([
                {
                    id: 'area_fach_gs',
                    type: 'area',
                    label: 'Fachleitungen Grundschule',
                    parentId: 'sec_arge',
                    order: 0,
                    tone: 'arge',
                    refKind: 'none',
                    refKey: ''
                },
                {
                    id: 'area_fach_gym',
                    type: 'area',
                    label: 'Fachleitungen Gymnasium',
                    parentId: 'sec_arge',
                    order: 1,
                    tone: 'arge',
                    refKind: 'none',
                    refKey: ''
                }
            ])
        });
        const argeGs = layout.nodes.find((n) => n.refKey === 'GSDE');
        const argeGym = layout.nodes.find((n) => n.refKey === 'GYDE');
        expect(argeGs?.parentId).toBe('sec_arge');
        expect(
            canPlaceSchulOrgNodeInSection(argeGs, 'area_fach_gs', layout)
        ).toBe(true);
        expect(
            canPlaceSchulOrgNodeInSection(argeGym, 'area_fach_gym', layout)
        ).toBe(true);
        expect(canPlaceSchulOrgNodeInSection(argeGs, 'sec_verwaltung', layout)).toBe(false);
        expect(canPlaceSchulOrgNodeInSection(argeGs, SCHUL_ORG_SEC_LEHRER, layout)).toBe(false);

        const moved = moveSchulOrganigrammArea(layout, argeGs.id, 'area_fach_gs');
        expect(moved.nodes.find((n) => n.id === argeGs.id)?.parentId).toBe('area_fach_gs');
        const moved2 = moveSchulOrganigrammArea(moved, argeGym.id, 'area_fach_gym');
        expect(moved2.nodes.find((n) => n.id === argeGym.id)?.parentId).toBe('area_fach_gym');
        // Zurück nach sec_arge
        const back = moveSchulOrganigrammArea(moved2, argeGs.id, 'sec_arge');
        expect(back.nodes.find((n) => n.id === argeGs.id)?.parentId).toBe('sec_arge');
        // Platzierung überlebt Rebuild/Sync
        const settings = {
            adminRoles: [],
            admin: [],
            arges: [
                { code: 'GSDE', name: 'ARGE FS GS Deutsch/DaF' },
                { code: 'GYDE', name: 'ARGE FS Gym Deutsch' }
            ],
            classes: []
        };
        const synced = syncSchulOrganigrammWithStammdaten(moved2, settings);
        expect(synced.layout.nodes.find((n) => n.refKey === 'GSDE')?.parentId).toBe('area_fach_gs');
        expect(synced.layout.nodes.find((n) => n.refKey === 'GYDE')?.parentId).toBe('area_fach_gym');
        const kidsGs = childrenOf(synced.layout, 'area_fach_gs').map((n) => n.refKey);
        expect(kidsGs).toContain('GSDE');
    });

    it('seedet SV-Board und SGA als Bereich mit drei Scope-Gruppen', () => {
        const layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            classes: [],
            studentCouncil: [
                { name: 'Lea', email: 'lea@schule.at', klasse: '1A' },
                { name: 'Tom', email: 'tom@schule.at', klasse: '2B' }
            ],
            sga: [
                { scope: 'student', name: 'Sam', email: 'sam@schule.at' },
                { scope: 'teacher', name: 'Müller', email: 'm@schule.at' },
                { scope: 'external', name: 'Eltern', email: 'e@schule.at' }
            ]
        });
        const councils = layout.nodes.filter((n) => n.type === NODE.COUNCIL);
        const sgaBoard = layout.nodes.find((n) => n.id === SCHUL_ORG_SGA_BOARD_ID);
        expect(councils).toHaveLength(3);
        expect(councils.map((c) => c.refKey).sort()).toEqual(
            SCHUL_ORG_SV_SEATS.map((s) => s.refKey).sort()
        );
        expect(councils.find((c) => c.refKey === 'schulsprecher')?.meta?.peopleCount).toBe(1);
        expect(sgaBoard?.type).toBe('area');
        expect(sgaBoard?.meta.peopleCount).toBe(3);
        expect(sgaBoard?.meta.studentCount).toBe(1);
        expect(sgaBoard?.meta.teacherCount).toBe(1);
        expect(sgaBoard?.meta.externalCount).toBe(1);
        expect(layout.nodes.filter((n) => n.refKind === 'sgaScope')).toHaveLength(3);
        expect(layout.nodes.find((n) => n.id === SCHUL_ORG_SGA_SV_AREA_ID)?.label).toBe(
            'Schülervertretung'
        );
    });

    it('sync entfernt alte Personen-Kacheln zugunsten der Boards', () => {
        const bloated = {
            version: 1,
            nodes: [
                {
                    id: 'sec_vertretung',
                    type: 'section',
                    label: 'Schülervertretung & SGA',
                    parentId: null,
                    order: 0,
                    tone: 'council'
                },
                {
                    id: 'council_old',
                    type: 'council',
                    label: 'Lea',
                    parentId: 'sec_vertretung',
                    order: 1,
                    refKind: 'studentCouncil',
                    refKey: 'lea@schule.at'
                },
                {
                    id: 'sga_old',
                    type: 'sga',
                    label: 'Sam',
                    parentId: 'sec_vertretung',
                    order: 2,
                    refKind: 'sga',
                    refKey: 'sam@schule.at'
                }
            ]
        };
        const synced = syncSchulOrganigrammWithStammdaten(bloated, {
            adminRoles: [],
            admin: [],
            arges: [],
            classes: [],
            studentCouncil: [{ name: 'Lea', email: 'lea@schule.at', klasse: '1A' }],
            sga: [{ scope: 'student', name: 'Sam', email: 'sam@schule.at' }]
        });
        expect(synced.layout.nodes.filter((n) => n.type === NODE.COUNCIL)).toHaveLength(3);
        expect(synced.layout.nodes.find((n) => n.id === SCHUL_ORG_SGA_BOARD_ID)?.type).toBe('area');
        expect(synced.layout.nodes.filter((n) => n.refKind === 'sgaScope')).toHaveLength(3);
        expect(synced.layout.nodes.find((n) => n.id === 'sv_schulsprecher')?.parentId).toBe(
            SCHUL_ORG_SGA_SV_AREA_ID
        );
        expect(synced.layout.nodes.some((n) => n.id === 'sga_old')).toBe(false);
        expect(synced.layout.nodes.some((n) => n.id === 'council_board')).toBe(false);
    });

    it('migriert v1 section → area inkl. view und Canvas-Koordinaten', () => {
        const v1 = {
            version: 1,
            nodes: [
                {
                    id: 'sec_schulleitung',
                    type: 'section',
                    label: 'Schulleitung',
                    parentId: null,
                    order: 0,
                    refKind: 'audience',
                    refKey: 'schulleitung'
                },
                {
                    id: 'role_x',
                    type: 'role',
                    label: 'Direktion',
                    parentId: 'sec_schulleitung',
                    order: 1,
                    refKind: 'adminRole',
                    refKey: 'DIREKTION'
                }
            ]
        };
        const n = normalizeSchulOrganigramm(v1);
        expect(n.version).toBe(SCHUL_ORG_VERSION);
        expect(n.view).toEqual({ x: 0, y: 0, k: 1 });
        const area = n.nodes.find((x) => x.id === 'sec_schulleitung');
        expect(area?.type).toBe('area');
        expect(isAreaNode(area)).toBe(true);
        expect(Number.isFinite(area?.x)).toBe(true);
        expect(Number.isFinite(area?.y)).toBe(true);
        const resolved = resolveSchulOrganigrammLayout(v1, {
            adminRoles: [{ code: 'DIREKTION', name: 'Direktion', tier: 'schulleitung' }],
            admin: [],
            arges: [],
            classes: []
        });
        expect(resolved.dirty).toBe(true);
        expect(resolved.layout.version).toBe(2);
    });

    it('verschachtelt Areas und speichert Viewport/Position', () => {
        let layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [
                { code: 'DIREKTION', name: 'Schulleitung', tier: 'schulleitung' },
                { code: 'STV', name: 'Stellv. Schulleitung', tier: 'schulleitung' }
            ],
            admin: [],
            arges: [],
            classes: []
        });
        layout = moveSchulOrganigrammArea(layout, 'sec_schulleitung', 'sec_verwaltung');
        const lead = layout.nodes.find((n) => n.id === 'sec_schulleitung');
        expect(lead?.parentId).toBe('sec_verwaltung');
        layout = setSchulOrganigrammNodePosition(layout, 'sec_schulleitung', 12, 8);
        layout = setSchulOrganigrammNodePosition(layout, 'sec_verwaltung', 120, 80);
        layout = setSchulOrganigrammView(layout, { x: 10, y: 20, k: 1.2 });
        expect(layout.view.k).toBeCloseTo(1.2);
        expect(layout.nodes.find((n) => n.id === 'sec_verwaltung')?.x).toBe(120);
        const map = computeLandkarteLayout(layout);
        const outer = map.boxes.find((b) => b.kind === 'area' && b.node.id === 'sec_verwaltung');
        const inner = map.boxes.find((b) => b.kind === 'area' && b.node.id === 'sec_schulleitung');
        expect(outer && inner).toBeTruthy();
        expect(inner.x + inner.w).toBeLessThanOrEqual(outer.x + outer.w + 0.5);
        expect(inner.y + inner.h).toBeLessThanOrEqual(outer.y + outer.h + 0.5);
        expect(outer.w).toBeGreaterThanOrEqual(inner.w);
        expect(outer.h).toBeGreaterThan(inner.h);
        // Absolute Alt-Koordinaten eines Nest-Kinds werden ignoriert (Packen oben → Eltern wächst)
        layout = setSchulOrganigrammNodePosition(layout, 'sec_schulleitung', 720, 40);
        const mapPacked = computeLandkarteLayout(layout);
        const outer2 = mapPacked.boxes.find((b) => b.kind === 'area' && b.node.id === 'sec_verwaltung');
        const inner2 = mapPacked.boxes.find((b) => b.kind === 'area' && b.node.id === 'sec_schulleitung');
        expect(inner2.x + inner2.w).toBeLessThanOrEqual(outer2.x + outer2.w + 0.5);
        expect(outer2.w).toBeLessThan(900);
        // Unterbereich über Rollen, Eltern enthält beides
        const roleLeaf = mapPacked.boxes.find(
            (b) => b.kind === 'leaf' && b.parentAreaId === 'sec_verwaltung'
        );
        if (roleLeaf) {
            expect(inner2.y).toBeLessThan(roleLeaf.y);
            expect(roleLeaf.y + roleLeaf.h).toBeLessThanOrEqual(outer2.y + outer2.h + 0.5);
        }
        // Zyklus verhindern
        const again = moveSchulOrganigrammArea(layout, 'sec_verwaltung', 'sec_schulleitung');
        expect(again.nodes.find((n) => n.id === 'sec_verwaltung')?.parentId).toBeNull();
        expect(map.width).toBeGreaterThan(100);
    });

    it('Root-Orte überleben normalize/sync/resolve (Speichern-Remount)', () => {
        expect(parseSchulOrgCanvasCoord(null)).toBeNull();
        expect(parseSchulOrgCanvasCoord('')).toBeNull();
        expect(parseSchulOrgCanvasCoord(0)).toBe(0);
        expect(parseSchulOrgCanvasCoord('120.4')).toBe(120);
        expect(defaultCanvasSlot(1).x).toBeGreaterThan(800);

        const settings = {
            adminRoles: [{ code: 'SL', name: 'Direktion', tier: 'schulleitung' }],
            admin: [],
            arges: Array.from({ length: 12 }, (_, i) => ({ code: 'A' + i, name: 'ARGE ' + i })),
            classes: [],
            teachers: [],
            students: [],
            sga: [],
            studentCouncil: []
        };
        let layout = seedSchulOrganigrammFromStammdaten(settings);
        layout = setSchulOrganigrammNodePosition(layout, 'sec_arge', 1600, 120);
        layout = setSchulOrganigrammNodePosition(layout, 'sec_lehrer', 2400, 900);
        layout = setSchulOrganigrammNodePosition(layout, SCHUL_ORG_SGA_BOARD_ID, 200, 1100);
        layout = setSchulOrganigrammView(layout, { x: 40, y: 20, k: 0.7 });

        // null/'' dürfen nicht zu (0,0) werden
        const withNull = normalizeSchulOrganigramm({
            version: 2,
            nodes: [
                { id: 'sec_verwaltung', type: 'area', label: 'V', parentId: null, order: 0, x: null, y: null },
                { id: 'sec_arge', type: 'area', label: 'A', parentId: null, order: 1, x: '', y: '' }
            ]
        });
        expect(withNull.nodes.find((n) => n.id === 'sec_verwaltung')?.x).not.toBe(0);
        expect(withNull.nodes.find((n) => n.id === 'sec_arge')?.x).not.toBe(
            withNull.nodes.find((n) => n.id === 'sec_verwaltung')?.x
        );

        const synced = syncSchulOrganigrammWithStammdaten(layout, settings);
        expect(synced.layout.nodes.find((n) => n.id === 'sec_arge')?.x).toBe(1600);
        expect(synced.layout.nodes.find((n) => n.id === 'sec_lehrer')?.y).toBe(900);

        const resolved = resolveSchulOrganigrammLayout(synced.layout, settings);
        expect(resolved.layout.nodes.find((n) => n.id === 'sec_arge')?.x).toBe(1600);
        expect(resolved.layout.view.k).toBeCloseTo(0.7);

        // Live-Merge wie nach Speichern/Notify
        const live = setSchulOrganigrammNodePosition(resolved.layout, 'sec_arge', 1800, 200);
        const merged = mergeSchulOrganigrammLivePositions(resolved.layout, live);
        expect(merged.nodes.find((n) => n.id === 'sec_arge')?.x).toBe(1800);

        const map = computeLandkarteLayout(merged);
        const persisted = applyLandkarteRootPositionsFromBoxes(merged, map.boxes);
        expect(persisted.nodes.find((n) => n.id === 'sec_arge')?.x).toBe(
            map.boxes.find((b) => b.kind === 'area' && b.node.id === 'sec_arge')?.x
        );

        // Enge Nutzer-Anordnung: keine automatische Entflechtung der Root-Positionen
        let packed = setSchulOrganigrammNodePosition(resolved.layout, 'sec_arge', 100, 80);
        packed = setSchulOrganigrammNodePosition(packed, 'sec_lehrer', 140, 120);
        packed = setSchulOrganigrammNodePosition(packed, SCHUL_ORG_SGA_BOARD_ID, 180, 160);
        const packedMap = computeLandkarteLayout(packed);
        expect(packedMap.boxes.find((b) => b.kind === 'area' && b.node.id === 'sec_arge')?.x).toBe(
            100
        );
        expect(packedMap.boxes.find((b) => b.kind === 'area' && b.node.id === 'sec_lehrer')?.y).toBe(
            120
        );
        expect(
            packedMap.boxes.find((b) => b.kind === 'area' && b.node.id === SCHUL_ORG_SGA_BOARD_ID)?.x
        ).toBe(180);
        const packedBack = applyLandkarteRootPositionsFromBoxes(packed, packedMap.boxes);
        expect(packedBack.nodes.find((n) => n.id === 'sec_lehrer')?.x).toBe(140);
        expect(packedBack.nodes.find((n) => n.id === SCHUL_ORG_SGA_BOARD_ID)?.y).toBe(160);

        // Speichern hält Root-Orte; Jahrgänge werden nicht als freie Orte persistiert
        const appliedSave = applySchulOrganigrammToStammdaten(
            { ...settings, schulOrganigramm: resolved.layout },
            resolved.layout
        );
        expect(appliedSave.layout.nodes.find((n) => n.id === 'sec_arge')?.x).toBe(1600);
        expect(appliedSave.layout.nodes.find((n) => n.id === 'sec_verwaltung')?.x).toBe(
            resolved.layout.nodes.find((n) => n.id === 'sec_verwaltung')?.x
        );

        const mapFlush = computeLandkarteLayout(appliedSave.layout);
        const flushed = applyLandkartePositionsFromBoxes(appliedSave.layout, mapFlush.boxes, {
            rootsOnly: true
        });
        expect(flushed.nodes.find((n) => n.id === 'sec_arge')?.x).toBe(
            mapFlush.boxes.find((b) => b.kind === 'area' && b.node.id === 'sec_arge')?.x
        );
    });

    it('Jahrgänge unter Schüler liegen nebeneinander ohne Überlappung', () => {
        const settings = {
            adminRoles: [],
            admin: [],
            arges: [],
            classes: [
                { code: '1A', name: '1A' },
                { code: '2A', name: '2A' },
                { code: '3A', name: '3A' },
                { code: '4A', name: '4A' }
            ],
            teachers: [],
            students: [],
            sga: [],
            studentCouncil: []
        };
        let layout = seedSchulOrganigrammFromStammdaten(settings);
        // Simuliere kaputte gleiche Relativ-Orte (früherer Bug)
        layout.nodes.forEach(function (n) {
            if (String(n.refKind) === 'jahrgang') {
                n.x = 12;
                n.y = 8;
            }
        });
        layout = ensureSchulOrganigrammJahrgangStructure(layout, settings);
        const map = computeLandkarteLayout(layout);
        const jahrgangBoxes = map.boxes.filter(
            (b) => b.kind === 'area' && b.node && b.node.refKind === 'jahrgang'
        );
        expect(jahrgangBoxes.length).toBeGreaterThanOrEqual(4);
        // Pairwise: keine starke Überlappung in X
        for (let i = 0; i < jahrgangBoxes.length; i++) {
            for (let j = i + 1; j < jahrgangBoxes.length; j++) {
                const a = jahrgangBoxes[i];
                const b = jahrgangBoxes[j];
                const overlap =
                    a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y;
                expect(overlap).toBe(false);
            }
        }
    });

    it('applySchulOrganigrammToStammdaten übernimmt Rollen-Tier und Labels in die Listen', () => {
        let layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [
                { code: 'SEK', name: 'Sekretariat', tier: 'verwaltung' },
                { code: 'DIR', name: 'Direktion', tier: 'schulleitung' }
            ],
            admin: [],
            arges: [{ code: 'MATHE', name: 'Mathematik' }],
            classes: [{ code: '1A', name: '1A' }],
            teachers: [],
            students: [],
            studentCouncil: [],
            sga: []
        });
        layout = moveSchulOrganigrammNode(layout, 'role_SEK', 'sec_schulleitung');
        expect(layout.nodes.find((n) => n.id === 'role_SEK')?.parentId).toBe('sec_schulleitung');
        const sek = layout.nodes.find((n) => n.id === 'role_SEK');
        if (sek) sek.label = 'Sekretariat / Empfang';
        const arge = layout.nodes.find((n) => n.refKey === 'MATHE');
        if (arge) arge.label = 'ARGE Mathematik';
        const applied = applySchulOrganigrammToStammdaten(
            {
                adminRoles: [
                    { code: 'SEK', name: 'Sekretariat', tier: 'verwaltung' },
                    { code: 'DIR', name: 'Direktion', tier: 'schulleitung' }
                ],
                arges: [{ code: 'MATHE', name: 'Mathematik' }],
                classes: [{ code: '1A', name: '1A' }],
                schulOrganigramm: layout
            },
            layout
        );
        expect(applied.settings.adminRoles.find((r) => r.code === 'SEK')?.tier).toBe('schulleitung');
        expect(applied.settings.adminRoles.find((r) => r.code === 'SEK')?.name).toBe(
            'Sekretariat / Empfang'
        );
        expect(applied.settings.arges.find((a) => a.code === 'MATHE')?.name).toBe('ARGE Mathematik');
        expect(applied.settings.schulOrganigramm.nodes.some((n) => n.id === 'role_SEK')).toBe(true);
        expect(applied.changed.roles).toBeGreaterThanOrEqual(1);
        expect(applied.changed.arges).toBeGreaterThanOrEqual(1);
    });

    it('Nest-Drop: tiefstes Kind / Herausziehen zum Großeltern', () => {
        const nodes = [
            { id: 'P', type: 'area', parentId: null },
            { id: 'A', type: 'area', parentId: 'P' },
            { id: 'B', type: 'area', parentId: 'P' },
            { id: 'C', type: 'area', parentId: 'A' }
        ];
        const boxes = [
            { kind: 'area', node: nodes[0], x: 0, y: 0, w: 400, h: 400 },
            { kind: 'area', node: nodes[1], x: 20, y: 40, w: 160, h: 160 },
            { kind: 'area', node: nodes[2], x: 220, y: 40, w: 160, h: 160 },
            { kind: 'area', node: nodes[3], x: 40, y: 60, w: 100, h: 80 }
        ];
        // Über parallelem Bereich B → B statt großer Parent P
        expect(pickDeepestAreaAtPoint(boxes, 250, 80, new Set(['C']), nodes)).toBe('B');
        expect(
            resolveLandkarteNestParent({
                boxes,
                nodes,
                worldX: 250,
                worldY: 80,
                dragId: 'C',
                blockedIds: new Set(['C']),
                currentParentId: 'A'
            })
        ).toEqual({ parentId: 'B', mode: 'nest' });
        // Außerhalb von A, noch in P → eine Ebene hoch (Geschwister von A)
        expect(
            resolveLandkarteNestParent({
                boxes,
                nodes,
                worldX: 200,
                worldY: 300,
                dragId: 'C',
                blockedIds: new Set(['C']),
                currentParentId: 'A'
            })
        ).toEqual({ parentId: 'P', mode: 'promote' });
        // Komplett außerhalb → Root
        expect(
            resolveLandkarteNestParent({
                boxes,
                nodes,
                worldX: 800,
                worldY: 800,
                dragId: 'C',
                blockedIds: new Set(['C']),
                currentParentId: 'A'
            })
        ).toEqual({ parentId: null, mode: 'root' });
    });

    it('Bereich herauslösen (Root) und in anderen Bereich verschieben', () => {
        let layout = normalizeSchulOrganigramm({
            version: 2,
            nodes: [
                { id: 'root_a', type: 'area', label: 'A', parentId: null, order: 1, x: 40, y: 40 },
                { id: 'root_b', type: 'area', label: 'B', parentId: null, order: 2, x: 500, y: 40 },
                { id: 'nest_c', type: 'area', label: 'C', parentId: 'root_a', order: 3 },
                { id: 'nest_d', type: 'area', label: 'D', parentId: 'nest_c', order: 4 }
            ]
        });
        // D aus C heraus → Root
        layout = moveSchulOrganigrammArea(layout, 'nest_d', null);
        expect(layout.nodes.find((n) => n.id === 'nest_d')?.parentId == null).toBe(true);
        // C von A nach B
        layout = moveSchulOrganigrammArea(layout, 'nest_c', 'root_b');
        expect(layout.nodes.find((n) => n.id === 'nest_c')?.parentId).toBe('root_b');
        // D wieder unter C (mitwandern der Hierarchie)
        layout = moveSchulOrganigrammArea(layout, 'nest_d', 'nest_c');
        expect(layout.nodes.find((n) => n.id === 'nest_d')?.parentId).toBe('nest_c');
        // C inkl. D bleiben unter B
        expect(layout.nodes.find((n) => n.id === 'nest_c')?.parentId).toBe('root_b');

        const nodes = layout.nodes;
        const boxes = [
            { kind: 'area', node: nodes.find((n) => n.id === 'root_a'), x: 0, y: 0, w: 200, h: 200 },
            { kind: 'area', node: nodes.find((n) => n.id === 'root_b'), x: 300, y: 0, w: 280, h: 260 },
            {
                kind: 'area',
                node: nodes.find((n) => n.id === 'nest_c'),
                x: 320,
                y: 80,
                w: 200,
                h: 140,
                parentAreaId: 'root_b'
            },
            {
                kind: 'area',
                node: nodes.find((n) => n.id === 'nest_d'),
                x: 340,
                y: 120,
                w: 160,
                h: 80,
                parentAreaId: 'nest_c'
            }
        ];
        // D außerhalb von C+B → Root
        expect(
            resolveLandkarteNestParent({
                boxes,
                nodes,
                worldX: 700,
                worldY: 400,
                dragId: 'nest_d',
                blockedIds: new Set(['nest_d']),
                currentParentId: 'nest_c'
            })
        ).toEqual({ parentId: null, mode: 'root' });
        // D außerhalb C, noch in B → Promote auf B
        expect(
            resolveLandkarteNestParent({
                boxes,
                nodes,
                worldX: 420,
                worldY: 240,
                dragId: 'nest_d',
                blockedIds: new Set(['nest_d']),
                currentParentId: 'nest_c'
            })
        ).toEqual({ parentId: 'root_b', mode: 'promote' });
        // C zurück nach A nesten
        expect(
            resolveLandkarteNestParent({
                boxes,
                nodes,
                worldX: 80,
                worldY: 40,
                dragId: 'nest_c',
                blockedIds: new Set(['nest_c', 'nest_d']),
                currentParentId: 'root_b'
            })
        ).toEqual({ parentId: 'root_a', mode: 'nest' });
    });

    it('verschachtelt Schulleitung/Verwaltung unter Custom-Audience und behält parentId', () => {
        let layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [
                { code: 'DIREKTION', name: 'Schulleitung', tier: 'schulleitung' },
                { code: 'STV', name: 'Stellv. Schulleitung', tier: 'schulleitung' }
            ],
            admin: [],
            arges: [],
            classes: [],
            verwaltungAudienceGroups: [{ id: 'va-erw', label: 'Erweiterte Schulleitung' }]
        });
        const audId = audienceSectionId('va-erw');
        expect(layout.nodes.some((n) => n.id === audId)).toBe(true);
        layout = moveSchulOrganigrammArea(layout, 'sec_schulleitung', audId);
        layout = moveSchulOrganigrammArea(layout, 'sec_verwaltung', audId);
        expect(layout.nodes.find((n) => n.id === 'sec_schulleitung')?.parentId).toBe(audId);
        expect(layout.nodes.find((n) => n.id === 'sec_verwaltung')?.parentId).toBe(audId);
        // normalize/repair darf Nesting nicht zurücksetzen
        const again = normalizeSchulOrganigramm(layout);
        expect(again.nodes.find((n) => n.id === 'sec_schulleitung')?.parentId).toBe(audId);
        expect(again.nodes.find((n) => n.id === audId)?.parentId == null).toBe(true);
        // Rollen bleiben unter Schulleitung
        expect(
            again.nodes.filter((n) => n.type === NODE.ROLE && n.parentId === 'sec_schulleitung').length
        ).toBeGreaterThanOrEqual(1);
    });

    it('Nest-Drop: befüllten Bereich auf anderen Bereichskopf ablegen', () => {
        // A ist „befüllt“ (Kind C), Ziel B ebenfalls (Kind D) – Ablegen auf B-Kopf → nest in B
        const nodes = [
            { id: 'A', type: 'area', parentId: null },
            { id: 'B', type: 'area', parentId: null },
            { id: 'C', type: 'area', parentId: 'A' },
            { id: 'D', type: 'area', parentId: 'B' }
        ];
        const boxes = [
            { kind: 'area', node: nodes[0], x: 0, y: 0, w: 200, h: 220 },
            { kind: 'area', node: nodes[1], x: 240, y: 0, w: 200, h: 220 },
            { kind: 'area', node: nodes[2], x: 20, y: 90, w: 160, h: 100 },
            { kind: 'area', node: nodes[3], x: 260, y: 90, w: 160, h: 100 }
        ];
        // Im Körper von B liegt Kind D – tiefstes Kind
        expect(pickDeepestAreaAtPoint(boxes, 300, 140, new Set(['A', 'C']), nodes)).toBe('D');
        // Im Kopf von B → B selbst (befüllter Container als Nest-Ziel)
        expect(pickDeepestAreaAtPoint(boxes, 300, 30, new Set(['A', 'C']), nodes)).toBe('B');
        expect(
            resolveLandkarteNestParent({
                boxes,
                nodes,
                worldX: 300,
                worldY: 30,
                dragId: 'A',
                blockedIds: new Set(['A', 'C']),
                currentParentId: null
            })
        ).toEqual({ parentId: 'B', mode: 'nest' });

        let layout = {
            version: 2,
            view: { x: 0, y: 0, k: 1 },
            nodes: nodes.map((n) => ({ ...n, label: n.id, order: 0 }))
        };
        layout = moveSchulOrganigrammArea(layout, 'A', 'B');
        expect(layout.nodes.find((n) => n.id === 'A')?.parentId).toBe('B');
        // Kinder von A bleiben unter A (Subtree wandert mit)
        expect(layout.nodes.find((n) => n.id === 'C')?.parentId).toBe('A');
        expect(layout.nodes.find((n) => n.id === 'D')?.parentId).toBe('B');
    });

    it('Drop über Leaf trifft Parent-Area; Nest-Container wächst um Unterbereich + Rollen', () => {
        // Leaf liegt als Geschwister über dem Area-Body – Drop muss Parent erkennen
        expect(
            pickLandkarteDropAreaId(
                [{ dropAreaId: '', parentAreaId: 'sec_verwaltung' }],
                new Set(['sec_sek'])
            )
        ).toBe('sec_verwaltung');
        expect(
            pickLandkarteDropAreaId(
                [
                    { dropAreaId: 'sec_sek', parentAreaId: '' },
                    { dropAreaId: 'sec_verwaltung', parentAreaId: '' }
                ],
                new Set(['sec_sek'])
            )
        ).toBe('sec_verwaltung');

        let layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [
                { code: 'SEK', name: 'Sekretariat', tier: 'verwaltung' },
                { code: 'IT', name: 'IT-Support', tier: 'verwaltung' },
                { code: 'BIB', name: 'Bibliothek', tier: 'verwaltung' },
                { code: 'ADM', name: 'Administration', tier: 'verwaltung' }
            ],
            admin: [],
            arges: [],
            classes: []
        });
        layout = {
            ...layout,
            nodes: layout.nodes.concat([
                {
                    id: 'sec_sek_bereich',
                    type: 'area',
                    label: 'Sekretariat',
                    parentId: null,
                    order: 99,
                    tone: 'staff',
                    x: 400,
                    y: 40
                }
            ])
        };
        layout = moveSchulOrganigrammArea(layout, 'sec_sek_bereich', 'sec_verwaltung');
        expect(layout.nodes.find((n) => n.id === 'sec_sek_bereich')?.parentId).toBe('sec_verwaltung');
        const map = computeLandkarteLayout(layout);
        const outer = map.boxes.find((b) => b.kind === 'area' && b.node.id === 'sec_verwaltung');
        const inner = map.boxes.find((b) => b.kind === 'area' && b.node.id === 'sec_sek_bereich');
        const roleLeaves = map.boxes.filter(
            (b) => b.kind === 'leaf' && b.parentAreaId === 'sec_verwaltung'
        );
        expect(outer && inner).toBeTruthy();
        expect(roleLeaves.length).toBeGreaterThanOrEqual(3);
        expect(inner.y).toBeLessThan(roleLeaves[0].y);
        expect(inner.y + inner.h).toBeLessThanOrEqual(outer.y + outer.h + 0.5);
        const lastLeaf = roleLeaves[roleLeaves.length - 1];
        expect(lastLeaf.y + lastLeaf.h).toBeLessThanOrEqual(outer.y + outer.h + 0.5);
        // Keine Überlappung Unterbereich ↔ Rollen
        roleLeaves.forEach((leaf) => {
            expect(leaf.y).toBeGreaterThanOrEqual(inner.y + inner.h - 0.5);
        });
    });

    it('legt Person-Slots und M365-Bindungen an Areas an', () => {
        let layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            classes: []
        });
        layout = upsertSchulOrganigrammAreaPerson(layout, 'sec_verwaltung', {
            email: 'sek@schule.at',
            name: 'Sekretariat'
        });
        expect(
            layout.nodes.some(
                (n) =>
                    n.type === NODE.PERSON &&
                    n.refKind === 'adminPerson' &&
                    n.refKey === 'sek@schule.at' &&
                    n.parentId === 'sec_verwaltung'
            )
        ).toBe(true);
        layout = setSchulOrganigrammAreaM365(layout, 'sec_verwaltung', {
            groupMode: 'match',
            groupId: 'gid-1',
            groupLabel: 'Verwaltung',
            mailboxMode: 'match',
            mailboxUpn: 'office@schule.at',
            mailboxLabel: 'Büro'
        });
        const area = layout.nodes.find((n) => n.id === 'sec_verwaltung');
        expect(area?.m365?.groupId).toBe('gid-1');
        expect(area?.m365?.mailboxUpn).toBe('office@schule.at');
        layout = removeSchulOrganigrammAreaPerson(layout, 'sek@schule.at', 'sec_verwaltung');
        expect(layout.nodes.some((n) => n.refKey === 'sek@schule.at')).toBe(false);
    });
});

describe('schul-organigramm canvas focus', () => {
    it('expandiert Fokus-Knoten auf demselben Canvas mit Personen-Chips', () => {
        const layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            classes: [{ code: '1A', name: '1A' }],
            studentCouncil: [],
            sga: []
        });
        const node = layout.nodes.find((n) => isClassNode(n) && n.refKey === '1A');
        const people = [
            { name: 'Lea', email: 'lea@schule.at' },
            { name: 'Sam', email: 'sam@schule.at' }
        ];
        const focus = computeFocusLayout(layout, node, people, []);
        expect(focus.boxes.some((b) => b.kind === 'expanded' && b.node.id === node.id)).toBe(true);
        expect(focus.boxes.filter((b) => b.kind === 'person-chip')).toHaveLength(2);
        expect(focus.boxes.some((b) => b.dimmed)).toBe(true);
        expect(focus.focusBox?.w).toBeGreaterThanOrEqual(680);
        expect(isDrillableNode(node)).toBe(true);
        expect(isDrillableNode({ type: 'person' })).toBe(false);
    });

    it('Bereich-Zoom zeigt Nested-Areas und Rollen wie auf der Landkarte', () => {
        let layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [
                { code: 'SL', name: 'Schulleitung', tier: 'schulleitung' },
                { code: 'KG', name: 'Kindergartenleitung', tier: 'verwaltung' },
                { code: 'GS', name: 'Grundschulleitung', tier: 'verwaltung' }
            ],
            admin: [],
            arges: [],
            classes: [],
            verwaltungAudienceGroups: [
                { id: 'va-erw', label: 'Erweiterte Schulleitung', builtin: false, order: 10 }
            ]
        });
        const erwId = audienceSectionId('va-erw');
        layout = moveSchulOrganigrammArea(layout, 'sec_schulleitung', erwId);
        for (const code of ['KG', 'GS']) {
            const role = layout.nodes.find((n) => n.refKey === code);
            if (role) layout = moveSchulOrganigrammNode(layout, role.id, erwId);
        }
        const erw = layout.nodes.find((n) => n.id === erwId);
        const kids = childrenOf(layout, erwId);
        const focus = computeFocusLayout(layout, erw, [], kids);
        expect(focus.boxes.some((b) => b.kind === 'expanded' && b.node.id === erwId)).toBe(true);
        const inside = focus.boxes.filter((b) => b.focusChild);
        expect(inside.some((b) => b.node && b.node.id === 'sec_schulleitung')).toBe(true);
        expect(inside.some((b) => b.node && b.node.refKey === 'KG')).toBe(true);
        expect(inside.some((b) => b.node && b.node.refKey === 'GS')).toBe(true);
        // Nested Schulleitung-Rollen ebenfalls sichtbar (Unterbaum)
        expect(inside.some((b) => b.node && b.node.refKey === 'SL')).toBe(true);
        const fb = focus.focusBox;
        const nested = inside.find((b) => b.node && b.node.id === 'sec_schulleitung');
        expect(nested.y).toBeGreaterThanOrEqual(fb.y);
        expect(nested.y + nested.h).toBeLessThanOrEqual(fb.y + fb.h + 1);
    });
});

describe('schul-organigramm jahrgaenge', () => {
    it('legt Jahrgänge unter Schüler an und gruppiert Klassen', () => {
        const layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            teachers: [],
            students: [],
            arges: [],
            classes: [
                { code: '1A', name: '1A' },
                { code: '1B', name: '1B' },
                { code: '2A', name: '2A' },
                { code: 'VK', name: 'Vorbereitung' }
            ]
        });
        const underSchueler = childrenOf(layout, SCHUL_ORG_SEC_SCHUELER).map((n) => n.id);
        expect(underSchueler).toContain(jahrgangAreaId('1'));
        expect(underSchueler).toContain(jahrgangAreaId('2'));
        expect(underSchueler).toContain(jahrgangAreaId('sonst'));
        expect(layout.nodes.find((n) => isClassNode(n) && n.refKey === '1A')?.parentId).toBe(
            jahrgangAreaId('1')
        );
        expect(layout.nodes.find((n) => isClassNode(n) && n.refKey === '2A')?.parentId).toBe(
            jahrgangAreaId('2')
        );
        expect(layout.nodes.find((n) => isClassNode(n) && n.refKey === 'VK')?.parentId).toBe(
            jahrgangAreaId('sonst')
        );
        const map = computeLandkarteLayout(layout);
        const jgBoxes = map.boxes.filter(
            (b) => b.kind === 'area' && String(b.node?.id || '').indexOf('sec_jahrgang_') === 0
        );
        expect(jgBoxes.length).toBeGreaterThanOrEqual(3);
        const xs = jgBoxes.map((b) => b.x).sort((a, b) => a - b);
        // Nebeneinander: aufsteigende X, nicht alle am gleichen X
        expect(xs[xs.length - 1] - xs[0]).toBeGreaterThan(200);
        for (let i = 1; i < xs.length; i += 1) {
            expect(xs[i]).toBeGreaterThan(xs[i - 1]);
        }
    });

    it('migriert Legacy sec_klassen in Jahrgänge', () => {
        const legacy = {
            version: 2,
            nodes: [
                {
                    id: SCHUL_ORG_SEC_SCHUELER,
                    type: 'area',
                    label: 'Schüler',
                    parentId: null,
                    refKind: 'sammelgruppe',
                    refKey: 'schueler'
                },
                {
                    id: 'sec_klassen',
                    type: 'area',
                    label: 'Klassen',
                    parentId: SCHUL_ORG_SEC_SCHUELER,
                    refKind: 'none',
                    refKey: ''
                },
                {
                    id: 'class_3A',
                    type: 'class',
                    label: '3A',
                    parentId: 'sec_klassen',
                    refKind: 'class',
                    refKey: '3A'
                }
            ]
        };
        const next = ensureSchulOrganigrammJahrgangStructure(legacy, {
            classes: [{ code: '3A', name: '3A' }]
        });
        expect(next.nodes.find((n) => n.id === 'sec_klassen')).toBeFalsy();
        expect(next.nodes.find((n) => n.id === jahrgangAreaId('3'))?.parentId).toBe(
            SCHUL_ORG_SEC_SCHUELER
        );
        expect(next.nodes.find((n) => n.refKey === '3A')?.parentId).toBe(jahrgangAreaId('3'));
    });
});

describe('schul-organigramm area people', () => {
    it('listet Lehrer-Sammelbereich; Schüler-Bereich zeigt keine flache Liste', () => {
        const layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            teachers: [
                { name: 'Anna', email: 'anna@schule.at' },
                { name: 'Bert', email: 'bert@schule.at' }
            ],
            students: [{ name: 'Clara', email: 'clara@schule.at', klasse: '1A' }],
            arges: [],
            classes: [{ code: '1A', name: '1A' }]
        });
        const lehrer = layout.nodes.find((n) => n.id === SCHUL_ORG_SEC_LEHRER);
        const schueler = layout.nodes.find((n) => n.id === SCHUL_ORG_SEC_SCHUELER);
        expect(isSammelgruppeArea(lehrer)).toBe(true);
        expect(isSammelgruppeArea(schueler)).toBe(true);
        expect(layout.nodes.find((n) => n.id === jahrgangAreaId('1'))?.parentId).toBe(
            SCHUL_ORG_SEC_SCHUELER
        );
        const settings = {
            schulOrganigramm: layout,
            teachers: [
                { name: 'Anna', email: 'anna@schule.at' },
                { name: 'Bert', email: 'bert@schule.at' }
            ],
            students: [{ name: 'Clara', email: 'clara@schule.at', klasse: '1A' }],
            classes: [{ code: '1A', name: '1A' }]
        };
        expect(listPeopleForSchulOrgNode(settings, lehrer).map((p) => p.email)).toEqual([
            'anna@schule.at',
            'bert@schule.at'
        ]);
        expect(listPeopleForSchulOrgNode(settings, schueler)).toHaveLength(0);
        const cls = layout.nodes.find((n) => isClassNode(n) && n.refKey === '1A');
        expect(listPeopleForSchulOrgNode(settings, cls).some((p) => p.email === 'clara@schule.at')).toBe(
            true
        );
        const blocked = addSchulOrgPersonToArea(settings, lehrer, {
            email: 'x@schule.at',
            name: 'X'
        });
        expect(blocked.ok).toBe(false);
    });

    it('ordnet Personen Audience-Areas über Membership zu', () => {
        const layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            classes: []
        });
        const area = layout.nodes.find((n) => n.id === 'sec_schulleitung');
        const settings = {
            schulOrganigramm: layout,
            admin: [],
            adminAudienceMemberships: [],
            verwaltungAudienceGroups: []
        };
        const added = addSchulOrgPersonToArea(settings, area, {
            email: 'dir@schule.at',
            name: 'Direktion'
        });
        expect(added.ok).toBe(true);
        expect(added.settings.adminAudienceMemberships).toEqual([
            { groupId: 'schulleitung', email: 'dir@schule.at', name: 'Direktion' }
        ]);
        const people = listPeopleForSchulOrgNode(added.settings, area);
        expect(people.map((p) => p.email)).toContain('dir@schule.at');
        const removed = removeSchulOrgPersonFromArea(added.settings, area, 'dir@schule.at');
        expect(removed.settings.adminAudienceMemberships).toHaveLength(0);
    });

    it('verschiebt Admin-Person auf Area-Slot (ohne Audience)', () => {
        let layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            classes: []
        });
        // Area ohne Audience
        layout.nodes.push({
            id: 'sec_sek',
            type: 'area',
            label: 'Sekretariat',
            parentId: 'sec_verwaltung',
            order: 99,
            refKind: 'none',
            refKey: ''
        });
        layout = normalizeSchulOrganigramm(layout);
        const sek = layout.nodes.find((n) => n.id === 'sec_sek');
        const settings = {
            schulOrganigramm: layout,
            admin: [{ role: 'Sek', name: 'Anna', email: 'anna@schule.at' }],
            adminRoles: [{ code: 'SEK', name: 'Sek' }],
            adminAudienceMemberships: []
        };
        const moved = moveSchulOrgPerson(
            settings,
            {
                kind: 'admin',
                name: 'Anna',
                email: 'anna@schule.at',
                sourceRefKind: 'adminRole',
                sourceRefKey: 'SEK'
            },
            sek
        );
        expect(moved.ok).toBe(true);
        expect(
            moved.settings.schulOrganigramm.nodes.some(
                (n) => n.type === 'person' && n.refKey === 'anna@schule.at' && n.parentId === 'sec_sek'
            )
        ).toBe(true);
    });

    it('Area-M365: groupKind normalisieren, matchen, Labels', () => {
        expect(normalizeSchulOrgGroupKind('mailSecurity')).toBe('mailSecurity');
        expect(normalizeSchulOrgGroupKind('dl')).toBe('distribution');
        expect(normalizeSchulOrgGroupKind('groupChat')).toBe('groupChat');
        expect(normalizeSchulOrgGroupKind('teams_chat')).toBe('groupChat');
        expect(schulOrgGroupKindLabel('security')).toBe('Sicherheitsgruppe');
        expect(schulOrgGroupKindLabel('groupChat')).toBe('Teams-Gruppenchat');
        expect(inferSchulOrgGroupKind({ securityEnabled: true, mailEnabled: true })).toBe(
            'mailSecurity'
        );
        expect(inferSchulOrgGroupKind({ groupTypes: ['Unified'] })).toBe('unified');
        expect(suggestSchulOrgMailNickname('Jahrgang 1')).toBe('jahrgang-1');

        let layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            classes: []
        });
        layout = setSchulOrganigrammAreaM365(layout, 'sec_verwaltung', {
            groupMode: 'match',
            groupKind: 'distribution',
            groupId: 'gid-dl-1',
            groupLabel: 'Verwaltung Verteiler',
            groupMail: 'verwaltung@schule.at',
            mailboxMode: 'match',
            mailboxUpn: 'verwaltung-pf@schule.at',
            mailboxLabel: 'Postfach Verwaltung'
        });
        const area = layout.nodes.find((n) => n.id === 'sec_verwaltung');
        expect(area?.m365?.groupKind).toBe('distribution');
        expect(area?.m365?.groupId).toBe('gid-dl-1');
        expect(area?.m365?.mailboxUpn).toBe('verwaltung-pf@schule.at');
        const normalized = normalizeSchulOrganigramm(layout);
        expect(normalized.nodes.find((n) => n.id === 'sec_verwaltung')?.m365?.groupKind).toBe(
            'distribution'
        );

        layout.nodes = layout.nodes.map((n) =>
            n.id === 'sec_verwaltung'
                ? Object.assign({}, n, {
                      externalIds: { bip: 'bip-42', webuntis: 'wu-9', empty: '' }
                  })
                : n
        );
        const withExt = normalizeSchulOrganigramm(layout);
        expect(withExt.nodes.find((n) => n.id === 'sec_verwaltung')?.externalIds).toEqual({
            bip: 'bip-42',
            webuntis: 'wu-9'
        });
    });

    it('Kursteams hängen unter Klassen; Mehrklassen-Team in beiden Klassen', () => {
        const settings = {
            adminRoles: [],
            admin: [],
            arges: [],
            teachers: [{ code: 'MUEL', name: 'Müller', email: 'mueller@schule.at' }],
            classes: [
                { code: '1AK', name: '1AK', year: '2034' },
                { code: '1BK', name: '1BK', year: '2034' },
                { code: '2AK', name: '2AK', year: '2033' }
            ],
            unterrichtsbelegung: {
                rows: [
                    {
                        klasse: '1AK1BK',
                        fach: 'MAT',
                        lehrerCode: 'MUEL',
                        lehrerEmail: 'mueller@schule.at',
                        teamName: 'MAT 1AK/1BK',
                        gruppenmail: 'jg2034-1ak1bk-mat',
                        graphGroupId: 'gid-shared-mat'
                    },
                    {
                        klasse: '2AK',
                        fach: 'DEU',
                        lehrerCode: 'MUEL',
                        lehrerEmail: 'mueller@schule.at',
                        teamName: 'DEU 2AK',
                        gruppenmail: 'jg2033-2ak-deu'
                    }
                ]
            }
        };
        const places = kursteamPlacementsFromBelegung(settings.unterrichtsbelegung, settings.classes);
        expect(places.filter((p) => p.stableKey === kursteamStableKey(settings.unterrichtsbelegung.rows[0]))).toHaveLength(
            2
        );
        expect(places.map((p) => p.classCode).sort()).toEqual(['1AK', '1BK', '2AK']);

        const layout = seedSchulOrganigrammFromStammdaten(settings);
        const kurs = layout.nodes.filter((n) => n.type === NODE.KURSTEAM);
        expect(kurs).toHaveLength(3);
        expect(kurs.filter((n) => n.parentId === 'class_1AK')).toHaveLength(1);
        expect(kurs.filter((n) => n.parentId === 'class_1BK')).toHaveLength(1);
        expect(kurs.filter((n) => n.parentId === 'class_2AK')).toHaveLength(1);
        const shared = kurs.filter((n) => n.meta && n.meta.graphGroupId === 'gid-shared-mat');
        expect(shared).toHaveLength(2);
        expect(shared.every((n) => n.parentId === 'class_1AK' || n.parentId === 'class_1BK')).toBe(true);
        expect(isJahrgangArea(layout.nodes.find((n) => n.id === jahrgangAreaId('1')))).toBe(true);
        expect(layout.nodes.find((n) => n.id === 'class_1AK')?.parentId).toBe(jahrgangAreaId('1'));
        expect(schulOrganigrammStats(layout).kursteams).toBe(3);

        // Landkarte-Übersicht: Klasse als Bereich (ohne einzelne Unterrichtsteam-Leaves)
        const map = computeLandkarteLayout(layout);
        const classBox = map.boxes.find(
            (b) => b.kind === 'area' && b.node && b.node.id === 'class_1AK'
        );
        const kursLeavesOverview = map.boxes.filter(
            (b) => b.kind === 'leaf' && b.node && b.node.type === NODE.KURSTEAM
        );
        expect(classBox).toBeTruthy();
        expect(isClassNode(classBox.node)).toBe(true);
        expect(classBox.classNest).toBeFalsy();
        expect(kursLeavesOverview).toHaveLength(0);

        // Reinzoomen in die Klasse → Unterrichtsteams sichtbar
        const classNode = layout.nodes.find((n) => n.id === 'class_1AK');
        const kursKids = layout.nodes.filter(
            (n) => n.parentId === 'class_1AK' && n.type === NODE.KURSTEAM
        );
        const focus = computeFocusLayout(layout, classNode, [], kursKids);
        const kursInFocus = focus.boxes.filter(
            (b) => b.focusChild && b.node && b.node.type === NODE.KURSTEAM
        );
        expect(kursInFocus).toHaveLength(1);
        expect(kursInFocus[0].node.label).toMatch(/MAT/i);

        // Mit Personen: links KV/Schüler, rechts Unterrichtsteams
        const focusSplit = computeFocusLayout(
            layout,
            classNode,
            [
                {
                    kind: 'head',
                    seatLabel: 'Klassenvorstand',
                    name: 'KV Test',
                    email: 'kv@schule.at'
                },
                { kind: 'member', seatLabel: 'Schüler', name: 'SuS', email: 'sus@schule.at' }
            ],
            kursKids
        );
        const kvChip = focusSplit.boxes.find((b) => b.kind === 'person-chip' && b.leadSeat);
        const schuelerChip = focusSplit.boxes.find(
            (b) => b.kind === 'person-chip' && !b.leadSeat
        );
        const kursChip = focusSplit.boxes.find(
            (b) => b.focusChild && b.node && b.node.type === NODE.KURSTEAM
        );
        expect(kvChip).toBeTruthy();
        expect(schuelerChip).toBeTruthy();
        expect(kursChip).toBeTruthy();
        expect(schuelerChip.y).toBeGreaterThan(kvChip.y);
        expect(kursChip.x).toBeGreaterThan(kvChip.x + kvChip.w - 1);
        expect(kursChip.y).toBeLessThanOrEqual(kvChip.y + 1);
        // Schüler-Chip so breit wie die linke Spalte (Name/E-Mail sichtbar)
        expect(schuelerChip.w).toBeGreaterThanOrEqual(kvChip.w - 1);

        // Sync entfernt veraltete Kursteams
        const slim = Object.assign({}, settings, {
            unterrichtsbelegung: {
                rows: [settings.unterrichtsbelegung.rows[1]]
            }
        });
        const synced = syncSchulOrganigrammWithStammdaten(layout, slim);
        expect(synced.layout.nodes.filter((n) => n.type === NODE.KURSTEAM)).toHaveLength(1);
        expect(synced.layout.nodes.find((n) => n.type === NODE.KURSTEAM)?.parentId).toBe('class_2AK');
    });

    it('keine doppelten Klassen unter Jahrgang und Sonstige', () => {
        const settings = {
            adminRoles: [],
            admin: [],
            arges: [],
            teachers: [],
            classes: [
                { code: '5HM', name: 'Klasse 5HM - SJ2026/27', year: '2027' },
                { code: '4HMA', name: 'Klasse 4HMA - SJ2026/27', year: '2028' },
                // Duplikat über Anzeigename ohne Kürzel (früher → Sonstige)
                { code: '', name: 'Klasse 5HM - SJ2026/27', year: '2027' },
                { code: 'Klasse 4HMB - SJ2026/27', name: 'Klasse 4HMB - SJ2026/27', year: '2028' },
                { code: 'Meisterklasse', name: 'Klasse Meisterklasse - SJ2026/27', year: '' }
            ]
        };
        const layout = seedSchulOrganigrammFromStammdaten(settings);
        const classes = layout.nodes.filter((n) => isClassNode(n));
        const codes = classes.map((n) => n.refKey).sort();
        expect(codes).toEqual(['4HMA', '4HMB', '5HM', 'MEISTERKLASSE']);
        expect(layout.nodes.find((n) => isClassNode(n) && n.refKey === '5HM')?.parentId).toBe(
            jahrgangAreaId('5')
        );
        expect(layout.nodes.find((n) => isClassNode(n) && n.refKey === '4HMA')?.parentId).toBe(
            jahrgangAreaId('4')
        );
        expect(layout.nodes.find((n) => isClassNode(n) && n.refKey === '4HMB')?.parentId).toBe(
            jahrgangAreaId('4')
        );
        expect(
            layout.nodes.find((n) => isClassNode(n) && n.refKey === 'MEISTERKLASSE')?.parentId
        ).toBe(jahrgangAreaId('sonst'));
        expect(classes.filter((n) => /5HM/i.test(n.label || '')).length).toBe(1);
    });

    it('mergeDuplicateSchulOrganigrammClasses vereinigt Jahrgang+Sonstige und behält Kursteams', () => {
        const messy = {
            version: 2,
            nodes: [
                {
                    id: 'sec_schueler',
                    type: 'area',
                    label: 'Schüler',
                    parentId: null,
                    order: 0,
                    refKind: 'sammelgruppe',
                    refKey: 'schueler'
                },
                {
                    id: jahrgangAreaId('5'),
                    type: 'area',
                    label: '5. Jahrgang',
                    parentId: 'sec_schueler',
                    order: 1,
                    refKind: 'jahrgang',
                    refKey: '5'
                },
                {
                    id: jahrgangAreaId('sonst'),
                    type: 'area',
                    label: 'Sonstige Klassen',
                    parentId: 'sec_schueler',
                    order: 2,
                    refKind: 'jahrgang',
                    refKey: 'sonst'
                },
                {
                    id: 'class_5HM',
                    type: 'class',
                    label: 'Klasse 5HM - SJ2026/27',
                    parentId: jahrgangAreaId('5'),
                    order: 3,
                    refKind: 'class',
                    refKey: '5HM'
                },
                {
                    id: 'class_Klasse 5HM - SJ2026/27',
                    type: 'class',
                    label: 'Klasse 5HM - SJ2026/27',
                    parentId: jahrgangAreaId('sonst'),
                    order: 4,
                    refKind: 'class',
                    refKey: 'Klasse 5HM - SJ2026/27'
                },
                {
                    id: 'kursteam_mat@Klasse 5HM - SJ2026/27',
                    type: 'kursteam',
                    label: 'MAT',
                    parentId: 'class_Klasse 5HM - SJ2026/27',
                    order: 5,
                    refKind: 'kursteam',
                    refKey: 'mat@Klasse 5HM - SJ2026/27'
                }
            ]
        };
        const { layout, merged } = mergeDuplicateSchulOrganigrammClasses(messy);
        expect(merged).toBeGreaterThanOrEqual(1);
        const classNodes = layout.nodes.filter((n) => isClassNode(n));
        expect(classNodes).toHaveLength(1);
        expect(classNodes[0].refKey).toBe('5HM');
        expect(classNodes[0].parentId).toBe(jahrgangAreaId('5'));
        const kurs = layout.nodes.filter((n) => n.type === NODE.KURSTEAM);
        expect(kurs).toHaveLength(1);
        expect(kurs[0].parentId).toBe('class_5HM');
    });

    it('Kursteams matchen auch „Klasse 2HMB“ / Punkt-Schreibweise', () => {
        const classes = [{ code: '2HMB', name: 'Klasse 2HMB - SJ2026/27' }];
        const places = kursteamPlacementsFromBelegung(
            {
                rows: [
                    { klasse: 'Klasse 2HMB', fach: 'MAT', gruppenmail: 'jg-2hmb-mat', teamName: 'MAT' },
                    { klasse: '2.HMB', fach: 'DEU', gruppenmail: 'jg-2hmb-deu', teamName: 'DEU' },
                    { klasse: '2HMB', fach: 'ENG', gruppenmail: 'jg-2hmb-eng', teamName: 'ENG' }
                ]
            },
            classes
        );
        expect(places.filter((p) => p.classCode === '2HMB')).toHaveLength(3);
    });

    it('klassenübergreifende Teams (2HMA/2HMB) erscheinen unter beiden Klassen', () => {
        const settings = {
            adminRoles: [],
            admin: [],
            arges: [],
            teachers: [],
            classes: [
                { code: '2HMA', name: 'Klasse 2HMA - SJ2026/27' },
                { code: '2HMB', name: 'Klasse 2HMB - SJ2026/27' }
            ],
            unterrichtsbelegung: {
                rows: [
                    {
                        klasse: '2HMA/2HMB',
                        fach: 'MAT',
                        teamName: 'MAT klassenübergreifend',
                        gruppenmail: 'jg-2hm-mat-shared'
                    },
                    {
                        klasse: '2HMA,2HMB',
                        fach: 'DEU',
                        teamName: 'DEU shared',
                        gruppenmail: 'jg-2hm-deu-shared'
                    }
                ]
            }
        };
        const places = kursteamPlacementsFromBelegung(
            settings.unterrichtsbelegung,
            settings.classes
        );
        expect(places.filter((p) => p.classCode === '2HMA')).toHaveLength(2);
        expect(places.filter((p) => p.classCode === '2HMB')).toHaveLength(2);

        const layout = seedSchulOrganigrammFromStammdaten(settings);
        expect(layout.nodes.filter((n) => n.parentId === 'class_2HMA' && n.type === NODE.KURSTEAM)).toHaveLength(
            2
        );
        expect(layout.nodes.filter((n) => n.parentId === 'class_2HMB' && n.type === NODE.KURSTEAM)).toHaveLength(
            2
        );
    });

    it('entfernt Geister-Areas unter Sonstige, die wie Klassen aussehen', () => {
        const settings = {
            adminRoles: [],
            admin: [],
            arges: [],
            teachers: [],
            classes: [
                { code: '4HMA', name: 'Klasse 4HMA - SJ2026/27' },
                { code: '4HMB', name: 'Klasse 4HMB - SJ2026/27' },
                { code: '5HM', name: 'Klasse 5HM - SJ2026/27' },
                { code: 'Meisterklasse', name: 'Meisterklasse - SJ2026/27' }
            ]
        };
        const messy = {
            version: 2,
            nodes: [
                {
                    id: 'sec_schueler',
                    type: 'area',
                    label: 'Schüler',
                    parentId: null,
                    order: 0,
                    refKind: 'sammelgruppe',
                    refKey: 'schueler'
                },
                {
                    id: jahrgangAreaId('4'),
                    type: 'area',
                    label: '4. Jahrgang',
                    parentId: 'sec_schueler',
                    order: 1,
                    refKind: 'jahrgang',
                    refKey: '4'
                },
                {
                    id: jahrgangAreaId('5'),
                    type: 'area',
                    label: '5. Jahrgang',
                    parentId: 'sec_schueler',
                    order: 2,
                    refKind: 'jahrgang',
                    refKey: '5'
                },
                {
                    id: jahrgangAreaId('sonst'),
                    type: 'area',
                    label: 'Sonstige Klassen',
                    parentId: 'sec_schueler',
                    order: 3,
                    refKind: 'jahrgang',
                    refKey: 'sonst'
                },
                {
                    id: 'class_4HMA',
                    type: 'class',
                    label: 'Klasse 4HMA - SJ2026/27',
                    parentId: jahrgangAreaId('4'),
                    order: 10,
                    refKind: 'class',
                    refKey: '4HMA'
                },
                {
                    id: 'class_4HMB',
                    type: 'class',
                    label: 'Klasse 4HMB - SJ2026/27',
                    parentId: jahrgangAreaId('4'),
                    order: 11,
                    refKind: 'class',
                    refKey: '4HMB'
                },
                {
                    id: 'class_5HM',
                    type: 'class',
                    label: 'Klasse 5HM - SJ2026/27',
                    parentId: jahrgangAreaId('5'),
                    order: 12,
                    refKind: 'class',
                    refKey: '5HM'
                },
                // Geister unter Sonstige (Area mit Klassen-Label) – das sah der User
                {
                    id: 'og_ghost_5hm',
                    type: 'area',
                    label: 'Klasse 5HM - SJ2026/27',
                    parentId: jahrgangAreaId('sonst'),
                    order: 20,
                    refKind: 'none',
                    tone: 'class'
                },
                {
                    id: 'og_ghost_4hma',
                    type: 'area',
                    label: 'Klasse 4HMA - SJ2026/27',
                    parentId: jahrgangAreaId('sonst'),
                    order: 21,
                    refKind: 'none',
                    tone: 'class'
                },
                {
                    id: 'og_ghost_4hmb',
                    type: 'area',
                    label: 'Klasse 4HMB - SJ2026/27',
                    parentId: jahrgangAreaId('sonst'),
                    order: 22,
                    refKind: 'class',
                    refKey: '4HMB',
                    tone: 'class'
                },
                {
                    id: 'class_MEISTERKLASSE',
                    type: 'class',
                    label: 'Meisterklasse - SJ2026/27',
                    parentId: jahrgangAreaId('sonst'),
                    order: 23,
                    refKind: 'class',
                    refKey: 'MEISTERKLASSE'
                }
            ]
        };

        expect(
            isPhantomSchulOrgClassNode(messy.nodes.find((n) => n.id === 'og_ghost_5hm'), new Set(['5HM']))
        ).toBe(true);

        const synced = syncSchulOrganigrammWithStammdaten(messy, settings);
        const classNodes = synced.layout.nodes.filter((n) => isClassNode(n));
        const codes = classNodes.map((n) => n.refKey).sort();
        expect(codes).toEqual(['4HMA', '4HMB', '5HM', 'MEISTERKLASSE']);
        expect(classNodes.filter((n) => n.refKey === '4HMA')).toHaveLength(1);
        expect(classNodes.filter((n) => n.refKey === '4HMB')).toHaveLength(1);
        expect(classNodes.filter((n) => n.refKey === '5HM')).toHaveLength(1);
        expect(
            synced.layout.nodes.find((n) => isClassNode(n) && n.refKey === '4HMA')?.parentId
        ).toBe(jahrgangAreaId('4'));
        expect(
            synced.layout.nodes.find((n) => isClassNode(n) && n.refKey === '5HM')?.parentId
        ).toBe(jahrgangAreaId('5'));
        // Keine Geister-Areas mehr unter Sonstige
        const sonstKids = synced.layout.nodes.filter((n) => n.parentId === jahrgangAreaId('sonst'));
        expect(sonstKids.every((n) => isClassNode(n))).toBe(true);
        expect(sonstKids.map((n) => n.refKey)).toEqual(['MEISTERKLASSE']);
        expect(synced.healed).toBe(true);
    });
});

describe('schul-organigramm save keeps landkarte positions', () => {
    it('applySchulOrganigrammToStammdaten + merge behält verschobene Root-Orte', () => {
        const settings = {
            adminRoles: [{ code: 'SL', name: 'Direktion', tier: 'schulleitung' }],
            admin: [],
            arges: Array.from({ length: 8 }, (_, i) => ({ code: 'A' + i, name: 'ARGE ' + i })),
            classes: [
                { code: '1A', name: '1A' },
                { code: '2A', name: '2A' },
                { code: '3A', name: '3A' }
            ],
            teachers: [],
            students: [],
            sga: [],
            studentCouncil: []
        };
        let layout = seedSchulOrganigrammFromStammdaten(settings);
        layout = setSchulOrganigrammNodePosition(layout, 'sec_schulleitung', 40, 30);
        layout = setSchulOrganigrammNodePosition(layout, 'sec_verwaltung', 420, 30);
        layout = setSchulOrganigrammNodePosition(layout, 'sec_lehrer', 900, 30);
        layout = setSchulOrganigrammNodePosition(layout, 'sec_schueler', 40, 520);
        layout = setSchulOrganigrammNodePosition(layout, 'sec_arge', 1500, 180);
        layout = setSchulOrganigrammNodePosition(layout, SCHUL_ORG_SGA_BOARD_ID, 700, 520);

        const map = computeLandkarteLayout(layout);
        layout = applyLandkartePositionsFromBoxes(layout, map.boxes, { rootsOnly: true });
        const applied = applySchulOrganigrammToStammdaten(settings, layout);
        const merged = mergeSchulOrganigrammLivePositions(applied.layout, layout);

        expect(merged.nodes.find((n) => n.id === 'sec_arge')?.x).toBe(1500);
        expect(merged.nodes.find((n) => n.id === 'sec_arge')?.y).toBe(180);
        expect(merged.nodes.find((n) => n.id === 'sec_schueler')?.x).toBe(40);
        expect(merged.nodes.find((n) => n.id === SCHUL_ORG_SGA_BOARD_ID)?.x).toBe(700);
        // ARGEs darf nicht auf dem Default-Slot über Schüler landen
        expect(merged.nodes.find((n) => n.id === 'sec_arge')?.x).toBeGreaterThan(1200);
    });
});

describe('schul-organigramm area window resize', () => {
    it('setzt und clampt Fenstergröße nur für Root-Areas', () => {
        let layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [{ code: 'SL', name: 'Schulleitung', tier: 'verwaltung' }],
            admin: [],
            arges: [],
            classes: []
        });
        const nested = {
            id: 'sec_nested_x',
            type: 'area',
            label: 'Nested',
            parentId: 'sec_verwaltung',
            order: 50,
            tone: 'staff'
        };
        layout = {
            ...layout,
            nodes: layout.nodes.concat([nested])
        };
        layout = setSchulOrganigrammNodeSize(layout, 'sec_verwaltung', 100, 50);
        const root = layout.nodes.find((n) => n.id === 'sec_verwaltung');
        expect(root?.w).toBe(SCHUL_ORG_AREA_MIN_W);
        expect(root?.h).toBe(SCHUL_ORG_AREA_MIN_H);

        layout = setSchulOrganigrammNodeSize(layout, 'sec_verwaltung', 640, 380);
        expect(layout.nodes.find((n) => n.id === 'sec_verwaltung')?.w).toBe(640);
        expect(layout.nodes.find((n) => n.id === 'sec_verwaltung')?.h).toBe(380);

        const beforeNested = layout.nodes.find((n) => n.id === 'sec_nested_x');
        layout = setSchulOrganigrammNodeSize(layout, 'sec_nested_x', 500, 300);
        const afterNested = layout.nodes.find((n) => n.id === 'sec_nested_x');
        expect(afterNested?.w).toBeUndefined();
        expect(afterNested?.h).toBeUndefined();
        expect(beforeNested?.parentId).toBe('sec_verwaltung');

        const clamped = clampSchulOrganigrammAreaSize(50, 9000);
        expect(clamped.w).toBe(SCHUL_ORG_AREA_MIN_W);
        expect(clamped.h).toBeLessThanOrEqual(4000);
    });

    it('persistiert w/h beim Speichern und mergen, erfindet keine Größe ohne Resize', () => {
        let layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            classes: []
        });
        const map0 = computeLandkarteLayout(layout);
        layout = applyLandkartePositionsFromBoxes(layout, map0.boxes, { rootsOnly: true });
        expect(layout.nodes.find((n) => n.id === 'sec_verwaltung')?.w).toBeUndefined();

        layout = setSchulOrganigrammNodeSize(layout, 'sec_verwaltung', 520, 260);
        const map1 = computeLandkarteLayout(layout);
        const box = map1.boxes.find((b) => b.kind === 'area' && b.node.id === 'sec_verwaltung');
        expect(box?.w).toBe(520);
        expect(box?.resizable).toBe(true);
        layout = applyLandkartePositionsFromBoxes(layout, map1.boxes, { rootsOnly: true });
        expect(layout.nodes.find((n) => n.id === 'sec_verwaltung')?.w).toBe(520);
        expect(layout.nodes.find((n) => n.id === 'sec_verwaltung')?.h).toBeGreaterThanOrEqual(260);

        const merged = mergeSchulOrganigrammLivePositions(
            seedSchulOrganigrammFromStammdaten({
                adminRoles: [],
                admin: [],
                arges: [],
                classes: []
            }),
            layout
        );
        expect(merged.nodes.find((n) => n.id === 'sec_verwaltung')?.w).toBe(520);
    });

    it('ordnet Inhalte bei schmalerer Fensterbreite neu an (Leaves schmaler, Jahrgänge umbrechen)', () => {
        let layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [
                { code: 'SL', name: 'Schulleitung', tier: 'verwaltung' },
                { code: 'AV', name: 'Admin', tier: 'verwaltung' },
                { code: 'IT', name: 'IT', tier: 'verwaltung' }
            ],
            admin: [],
            arges: [],
            classes: [
                { code: '1A', name: '1A' },
                { code: '2A', name: '2A' },
                { code: '3A', name: '3A' },
                { code: '4A', name: '4A' }
            ]
        });
        layout = setSchulOrganigrammNodeSize(layout, 'sec_verwaltung', 300, 400);
        const mapAdmin = computeLandkarteLayout(layout);
        const adminBox = mapAdmin.boxes.find(
            (b) => b.kind === 'area' && b.node.id === 'sec_verwaltung'
        );
        const leaves = mapAdmin.boxes.filter(
            (b) => b.kind === 'leaf' && b.parentAreaId === 'sec_verwaltung'
        );
        expect(adminBox?.w).toBe(300);
        expect(leaves.length).toBeGreaterThanOrEqual(3);
        leaves.forEach((leaf) => {
            expect(leaf.w).toBeLessThanOrEqual(300 - 20);
        });

        layout = setSchulOrganigrammNodeSize(layout, 'sec_schueler', 320, 500);
        const mapSch = computeLandkarteLayout(layout);
        const schBox = mapSch.boxes.find((b) => b.kind === 'area' && b.node.id === 'sec_schueler');
        const jahrgangBoxes = mapSch.boxes.filter(
            (b) => b.kind === 'area' && b.node && isJahrgangArea(b.node)
        );
        expect(schBox?.w).toBe(320);
        expect(jahrgangBoxes.length).toBeGreaterThanOrEqual(2);
        // Bei schmalem Fenster müssen Jahrgänge untereinander rutschen (nicht nur eine Zeile)
        const tops = [...new Set(jahrgangBoxes.map((b) => Math.round(b.y)))].sort((a, b) => a - b);
        expect(tops.length).toBeGreaterThanOrEqual(2);
        jahrgangBoxes.forEach((jg) => {
            expect(jg.x + jg.w).toBeLessThanOrEqual(schBox.x + schBox.w + 1);
        });
    });

    it('Eltern- und KlassenChats-Fenster: Resize + Jahrgänge umbrechen wie bei Schüler', () => {
        let layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            classes: [
                { code: '1A', name: '1A' },
                { code: '2A', name: '2A' },
                { code: '3A', name: '3A' },
                { code: '4A', name: '4A' }
            ],
            students: []
        });

        layout = setSchulOrganigrammNodeSize(layout, SCHUL_ORG_SEC_ELTERN, 320, 480);
        layout = setSchulOrganigrammNodeSize(layout, SCHUL_ORG_SEC_KLASSENCHATS, 320, 480);
        const map = computeLandkarteLayout(layout);

        const elternBox = map.boxes.find(
            (b) => b.kind === 'area' && b.node.id === SCHUL_ORG_SEC_ELTERN
        );
        const chatBox = map.boxes.find(
            (b) => b.kind === 'area' && b.node.id === SCHUL_ORG_SEC_KLASSENCHATS
        );
        expect(elternBox?.resizable).toBe(true);
        expect(chatBox?.resizable).toBe(true);
        expect(elternBox?.w).toBe(320);
        expect(chatBox?.w).toBe(320);

        const elternJg = map.boxes.filter(
            (b) => b.kind === 'area' && b.node && isElternJahrgangArea(b.node)
        );
        const chatJg = map.boxes.filter(
            (b) => b.kind === 'area' && b.node && isKlassenchatJahrgangArea(b.node)
        );
        expect(elternJg.length).toBeGreaterThanOrEqual(2);
        expect(chatJg.length).toBeGreaterThanOrEqual(2);

        const elternTops = [...new Set(elternJg.map((b) => Math.round(b.y)))].sort((a, b) => a - b);
        const chatTops = [...new Set(chatJg.map((b) => Math.round(b.y)))].sort((a, b) => a - b);
        expect(elternTops.length).toBeGreaterThanOrEqual(2);
        expect(chatTops.length).toBeGreaterThanOrEqual(2);

        elternJg.forEach((jg) => {
            expect(jg.x + jg.w).toBeLessThanOrEqual(elternBox.x + elternBox.w + 1);
        });
        chatJg.forEach((jg) => {
            expect(jg.x + jg.w).toBeLessThanOrEqual(chatBox.x + chatBox.w + 1);
        });

        // Breite Fenster: Jahrgänge wieder nebeneinander
        layout = setSchulOrganigrammNodeSize(layout, SCHUL_ORG_SEC_ELTERN, 1400, 500);
        const mapWide = computeLandkarteLayout(layout);
        const elternWide = mapWide.boxes.find(
            (b) => b.kind === 'area' && b.node.id === SCHUL_ORG_SEC_ELTERN
        );
        const elternJgWide = mapWide.boxes.filter(
            (b) => b.kind === 'area' && b.node && isElternJahrgangArea(b.node)
        );
        const lefts = [...new Set(elternJgWide.map((b) => Math.round(b.x)))].sort((a, b) => a - b);
        expect(elternWide?.w).toBe(1400);
        expect(lefts.length).toBeGreaterThanOrEqual(2);
    });

    it('verschachtelte Bereiche bleiben innerhalb der Elternbreite (Bereich-in-Bereich)', () => {
        let layout = normalizeSchulOrganigramm({
            version: 2,
            nodes: [
                { id: 'root_a', type: 'area', label: 'Root', parentId: null, order: 1, x: 40, y: 40 },
                { id: 'nest_b', type: 'area', label: 'Nest B', parentId: 'root_a', order: 2 },
                { id: 'nest_c', type: 'area', label: 'Nest C', parentId: 'root_a', order: 3 },
                { id: 'nest_d', type: 'area', label: 'Nest D', parentId: 'nest_b', order: 4 },
                { id: 'nest_e', type: 'area', label: 'Nest E', parentId: 'nest_b', order: 5 },
                {
                    id: 'role_1',
                    type: 'role',
                    label: 'Rolle 1',
                    parentId: 'nest_d',
                    order: 6,
                    refKind: 'adminRole',
                    refKey: 'R1'
                },
                {
                    id: 'role_2',
                    type: 'role',
                    label: 'Rolle 2',
                    parentId: 'nest_e',
                    order: 7,
                    refKind: 'adminRole',
                    refKey: 'R2'
                }
            ]
        });
        layout = setSchulOrganigrammNodeSize(layout, 'root_a', 300, 400);
        const map = computeLandkarteLayout(layout);
        const areas = map.boxes.filter((b) => b.kind === 'area');
        const root = areas.find((b) => b.node.id === 'root_a');
        expect(root?.w).toBe(300);
        areas.forEach((a) => {
            if (!a.parentAreaId) return;
            const p = areas.find((x) => x.node.id === a.parentAreaId);
            expect(p).toBeTruthy();
            expect(a.x + a.w).toBeLessThanOrEqual(p.x + p.w + 1.5);
        });
        map.boxes
            .filter((b) => b.kind === 'leaf')
            .forEach((leaf) => {
                const p = areas.find((x) => x.node.id === leaf.parentAreaId);
                expect(p).toBeTruthy();
                expect(leaf.x + leaf.w).toBeLessThanOrEqual(p.x + p.w + 1.5);
            });

        // Schüler: Klassen-Nester müssen in den Jahrgangs-Fenstern bleiben
        layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            classes: [
                { code: '1A' },
                { code: '1B' },
                { code: '1C' },
                { code: '2A' },
                { code: '2B' },
                { code: '3A' }
            ],
            students: []
        });
        layout = setSchulOrganigrammNodeSize(layout, SCHUL_ORG_SEC_SCHUELER, 340, 500);
        const mapSch = computeLandkarteLayout(layout);
        const schAreas = mapSch.boxes.filter((b) => b.kind === 'area');
        schAreas.forEach((a) => {
            if (!a.parentAreaId) return;
            const p = schAreas.find((x) => x.node.id === a.parentAreaId);
            expect(p).toBeTruthy();
            expect(a.x + a.w).toBeLessThanOrEqual(p.x + p.w + 1.5);
        });
    });
});

describe('schul-organigramm ARGE-Leiter', () => {
    it('ensure legt Bereich unter ARGEs an und übernimmt manuellen Namen', () => {
        const layout = ensureSchulOrganigrammArgeLeiterStructure({
            nodes: [
                {
                    id: 'sec_arge',
                    type: 'area',
                    label: 'ARGEs / Fachschaften',
                    parentId: null,
                    order: 1
                },
                {
                    id: 'area_custom',
                    type: 'area',
                    label: 'ARGE-Leiter',
                    parentId: 'sec_arge',
                    order: 2,
                    refKind: 'none'
                }
            ]
        });
        const leiter = layout.nodes.find((n) => n.id === SCHUL_ORG_SEC_ARGE_LEITER);
        expect(leiter).toBeTruthy();
        expect(leiter.parentId).toBe('sec_arge');
        expect(leiter.refKind).toBe('argeLeiter');
        expect(layout.nodes.filter((n) => /arge[\s_-]*leiter/i.test(n.label || '')).length).toBe(1);
    });

    it('collectArgeLeiterPeopleFromSettings dedupliziert nach E-Mail', () => {
        const people = collectArgeLeiterPeopleFromSettings({
            arges: [
                { code: 'MAT', name: 'Mathe', headName: 'Anna', headEmail: 'anna@schule.at' },
                { code: 'MUS', name: 'Musik', headName: 'Bert', headEmail: 'bert@schule.at' },
                { code: 'SPO', name: 'Sport', headName: 'Anna', headEmail: 'Anna@schule.at' }
            ]
        });
        expect(people.map((p) => p.email)).toEqual(['anna@schule.at', 'bert@schule.at']);
    });

    it('listArgeCatalogGroupRefs liest catalogLinks kind=arge', () => {
        expect(
            listArgeCatalogGroupRefs({
                catalogLinks: [
                    { kind: 'arge', code: 'MAT', graphGroupId: 'g1', displayName: 'Mathe' },
                    { kind: 'class', code: '1A', graphGroupId: 'g2' },
                    { kind: 'arge', code: 'MAT', graphGroupId: 'g1' },
                    { kind: 'arge', code: 'MUS', graphGroupId: 'g3' }
                ]
            }).map((r) => r.groupId)
        ).toEqual(['g1', 'g3']);
    });

    it('listPeople zeigt ARGE-Leitungen im Bereich', () => {
        const layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [{ code: 'MAT', name: 'Mathe', headName: 'Anna', headEmail: 'anna@schule.at' }],
            classes: []
        });
        const settings = {
            arges: [{ code: 'MAT', name: 'Mathe', headName: 'Anna', headEmail: 'anna@schule.at' }],
            schulOrganigramm: layout
        };
        const people = listPeopleForSchulOrgNode(
            settings,
            layout.nodes.find((n) => n.id === SCHUL_ORG_SEC_ARGE_LEITER)
        );
        expect(people).toHaveLength(1);
        expect(people[0].email).toBe('anna@schule.at');
        expect(people[0].seatLabel).toBe('ARGE-Leitung');
    });
});

describe('schul-organigramm pages', () => {
    it('Tool-Seite und Stammdaten-Tab existieren', () => {
        const tool = readFileSync(join(projectRoot, 'tools/schul-organigramm.html'), 'utf8');
        expect(tool).toContain('schulOrgApp');
        expect(tool).toContain('mountSchulOrganigramm');
        const tenant = readFileSync(join(projectRoot, 'tenant.html'), 'utf8');
        expect(tenant).toContain('id="tabMainOrganigramm"');
        expect(tenant).toContain('tenantSchulOrganigrammApp');
        const index = readFileSync(join(projectRoot, 'index.html'), 'utf8');
        expect(index).toContain('data-tool-id="schul-organigramm"');
    });
});
