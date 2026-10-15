import { describe, expect, it } from 'vitest';
import { buildVerwaltungStammdatenTree } from '../src/shared/verwaltung-stammdaten-tree.js';
import { moveSchulOrgRole } from '../src/tools/schul-organigramm/schul-organigramm-edit.js';
import {
    audienceSectionId,
    rebuildSchulOrganigrammFromStammdaten,
    syncSchulOrganigrammWithStammdaten
} from '../src/tools/schul-organigramm/schul-organigramm-logic.js';

describe('verwaltung-stammdaten-tree', () => {
    it('baut verschachtelten Baum Bereich → Unterbereich → Funktion → Person', () => {
        const aud = audienceSectionId('va-erw');
        const tree = buildVerwaltungStammdatenTree({
            adminRoles: [{ code: 'QMS', name: 'Qualitätsmanagement', seatMode: 'single' }],
            admin: [
                {
                    role: 'Qualitätsmanagement',
                    name: 'Clara',
                    email: 'clara@schule.at',
                    defaultKey: 'Qualitätsmanagement'
                }
            ],
            adminAudienceMemberships: [],
            verwaltungAudienceGroups: [
                { id: 'va-erw', label: 'Erw. Schulleitung', builtin: false, order: 10 }
            ],
            schulOrganigramm: {
                version: 2,
                nodes: [
                    {
                        id: 'sec_verwaltung',
                        type: 'area',
                        label: 'Verwaltung (Personal)',
                        parentId: null,
                        order: 0,
                        refKind: 'none'
                    },
                    {
                        id: aud,
                        type: 'area',
                        label: 'Erw. Schulleitung',
                        parentId: 'sec_verwaltung',
                        order: 1,
                        refKind: 'audience',
                        refKey: 'va-erw'
                    },
                    {
                        id: 'role_QMS',
                        type: 'role',
                        label: 'Qualitätsmanagement',
                        parentId: aud,
                        order: 0,
                        refKind: 'adminRole',
                        refKey: 'QMS'
                    }
                ]
            }
        });

        expect(tree.source).toBe('organigramm');
        expect(tree.counts.areas).toBeGreaterThanOrEqual(2);
        expect(tree.counts.roles).toBe(1);
        expect(tree.counts.people).toBe(1);

        const root = tree.roots.find((r) => r.id === 'sec_verwaltung');
        expect(root).toBeTruthy();
        const nested = root.children.find((c) => c.id === aud);
        expect(nested?.kind).toBe('area');
        expect(nested?.label).toBe('Erw. Schulleitung');
        const role = nested.children.find((c) => c.kind === 'role');
        expect(role?.label).toMatch(/Qualität/);
        expect(role?.children[0]?.email).toBe('clara@schule.at');
    });

    it('liefert leeren Baum ohne Organigramm-Bereiche', () => {
        const tree = buildVerwaltungStammdatenTree({
            adminRoles: [],
            admin: [],
            schulOrganigramm: { version: 2, nodes: [] }
        });
        expect(tree.source).toBe('empty');
        expect(tree.roots).toEqual([]);
    });

    it('verschiebt Funktion (z. B. IT-Support) in Unterbereich IT-Team', () => {
        const settings = {
            adminRoles: [
                { code: 'IT-SUPPORT', name: 'IT-Support', seatMode: 'single', tier: 'verwaltung' }
            ],
            admin: [
                {
                    role: 'IT-Support',
                    name: 'IT Admin',
                    email: 'demo-it-admin@ms365.schule',
                    defaultKey: 'IT-Support'
                }
            ],
            adminAudienceMemberships: [],
            schulOrganigramm: {
                version: 2,
                nodes: [
                    {
                        id: 'sec_verwaltung',
                        type: 'area',
                        label: 'Verwaltung (Personal)',
                        parentId: null,
                        order: 0,
                        refKind: 'none'
                    },
                    {
                        id: 'area_it_team',
                        type: 'area',
                        label: 'IT-Team',
                        parentId: 'sec_verwaltung',
                        order: 0,
                        refKind: 'none'
                    },
                    {
                        id: 'role_IT-SUPPORT',
                        type: 'role',
                        label: 'IT-Support',
                        parentId: 'sec_verwaltung',
                        order: 1,
                        refKind: 'adminRole',
                        refKey: 'IT-SUPPORT'
                    }
                ]
            }
        };

        const before = buildVerwaltungStammdatenTree(settings);
        const vw = before.roots.find((r) => r.id === 'sec_verwaltung');
        expect(vw.children.some((c) => c.id === 'role_IT-SUPPORT')).toBe(true);
        expect(
            vw.children.find((c) => c.id === 'area_it_team')?.children.some(
                (c) => c.id === 'role_IT-SUPPORT'
            )
        ).toBeFalsy();

        const moved = moveSchulOrgRole(settings, 'role_IT-SUPPORT', 'area_it_team');
        expect(moved.ok).toBe(true);
        expect(
            moved.settings.schulOrganigramm.nodes.find((n) => n.id === 'role_IT-SUPPORT')?.parentId
        ).toBe('area_it_team');

        const after = buildVerwaltungStammdatenTree(moved.settings);
        const root = after.roots.find((r) => r.id === 'sec_verwaltung');
        const itTeam = root.children.find((c) => c.id === 'area_it_team');
        expect(itTeam?.kind).toBe('area');
        expect(itTeam.children.some((c) => c.id === 'role_IT-SUPPORT')).toBe(true);
        expect(root.children.some((c) => c.id === 'role_IT-SUPPORT')).toBe(false);

        // Sync/Heal darf IT-Team und die neue Parent-Zuordnung nicht verwerfen
        const synced = syncSchulOrganigrammWithStammdaten(
            moved.settings.schulOrganigramm,
            moved.settings
        );
        expect(synced.layout.nodes.find((n) => n.id === 'area_it_team')).toBeTruthy();
        expect(
            synced.layout.nodes.find((n) => n.id === 'role_IT-SUPPORT')?.parentId
        ).toBe('area_it_team');

        const rebuilt = rebuildSchulOrganigrammFromStammdaten(
            moved.settings.schulOrganigramm,
            moved.settings
        );
        expect(rebuilt.nodes.find((n) => n.id === 'area_it_team')).toBeTruthy();
        expect(rebuilt.nodes.find((n) => n.id === 'role_IT-SUPPORT')?.parentId).toBe(
            'area_it_team'
        );
        const treeAfterHeal = buildVerwaltungStammdatenTree({
            ...moved.settings,
            schulOrganigramm: rebuilt
        });
        const itAfterHeal = treeAfterHeal.roots
            .find((r) => r.id === 'sec_verwaltung')
            ?.children.find((c) => c.id === 'area_it_team');
        expect(itAfterHeal?.children.some((c) => c.id === 'role_IT-SUPPORT')).toBe(true);
    });
});
