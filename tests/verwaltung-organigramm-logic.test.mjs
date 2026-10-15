import { describe, expect, it } from 'vitest';
import {
    buildVerwaltungOrganigramm,
    organigramRoleNode,
    seedOrganigrammFromRoles,
    syncOrganigrammWithRoles,
    movePersonToRole,
    moveNode,
    childrenOf,
    updateOrgNodeLabel,
    ORG_NODE_ROLE
} from '../src/tools/verwaltung/verwaltung-organigramm-logic.js';

describe('verwaltung-organigramm-logic', () => {
    it('groups roles by schulleitung vs verwaltung', () => {
        const roles = [
            { code: 'DIREKTION', name: 'Direktion', tier: 'schulleitung' },
            { code: 'SEK', name: 'Sekretariat', tier: 'verwaltung' },
            { code: 'IT', name: 'IT-Support', tier: 'verwaltung' }
        ];
        const people = {
            DIREKTION: [{ name: 'Holtmann', email: 'h@schule.de' }],
            SEK: [
                { name: 'Ganz', email: 'g@schule.de' },
                { name: 'Bermudez', email: 'b@schule.de' }
            ],
            IT: []
        };
        const tree = buildVerwaltungOrganigramm(roles, (r) => people[r.code] || []);
        expect(tree.sections.map((s) => s.id)).toEqual(['schulleitung', 'verwaltung']);
        expect(tree.sections[0].roles).toHaveLength(1);
        expect(tree.sections[1].roles).toHaveLength(2);
        expect(tree.stats.people).toBe(3);
        expect(tree.stats.vacant).toBe(1);
    });

    it('infers direktion as schulleitung without explicit tier', () => {
        const node = organigramRoleNode({ code: 'DIR', name: 'Direktion' }, []);
        expect(node.tier).toBe('schulleitung');
        expect(node.vacant).toBe(true);
    });

    it('filters by name or email', () => {
        const roles = [
            { code: 'DIREKTION', name: 'Direktion', tier: 'schulleitung' },
            { code: 'SEK', name: 'Sekretariat', tier: 'verwaltung' }
        ];
        const tree = buildVerwaltungOrganigramm(
            roles,
            (r) => (r.code === 'SEK' ? [{ name: 'Fr. Ganz', email: 'ganz@schule.de' }] : []),
            { filter: 'ganz' }
        );
        expect(tree.sections).toHaveLength(1);
        expect(tree.sections[0].roles[0].code).toBe('SEK');
    });

    it('seeds layout with two root groups and role nodes', () => {
        const layout = seedOrganigrammFromRoles([
            { code: 'DIREKTION', name: 'Direktion', tier: 'schulleitung' },
            { code: 'SEK', name: 'Sekretariat', tier: 'verwaltung' }
        ]);
        const roots = childrenOf(layout, null);
        expect(roots).toHaveLength(2);
        const roles = layout.nodes.filter((n) => n.type === ORG_NODE_ROLE);
        expect(roles).toHaveLength(2);
    });

    it('sync adds missing role nodes', () => {
        const seeded = seedOrganigrammFromRoles([{ code: 'SEK', name: 'Sekretariat', tier: 'verwaltung' }]);
        const synced = syncOrganigrammWithRoles(seeded, [
            { code: 'SEK', name: 'Sekretariat', tier: 'verwaltung' },
            { code: 'IT', name: 'IT-Support', tier: 'verwaltung' }
        ]);
        expect(synced.nodes.filter((n) => n.type === ORG_NODE_ROLE).map((n) => n.roleCode).sort()).toEqual([
            'IT',
            'SEK'
        ]);
    });

    it('sync updates role labels from stammdaten', () => {
        const seeded = seedOrganigrammFromRoles([{ code: 'SEK', name: 'Sekretariat', tier: 'verwaltung' }]);
        const synced = syncOrganigrammWithRoles(seeded, [
            { code: 'SEK', name: 'Sekretariat GS', tier: 'verwaltung' }
        ]);
        const role = synced.nodes.find((n) => n.type === ORG_NODE_ROLE && n.roleCode === 'SEK');
        expect(role.label).toBe('Sekretariat GS');
    });

    it('renames group nodes in layout', () => {
        const layout = seedOrganigrammFromRoles([]);
        const lead = layout.nodes.find((n) => n.tone === 'lead');
        const next = updateOrgNodeLabel(layout, lead.id, 'Erweiterte Schulleitung');
        expect(next.nodes.find((n) => n.id === lead.id).label).toBe('Erweiterte Schulleitung');
    });

    it('moves person to another role by email', () => {
        const rows = [
            { role: 'Sekretariat', name: 'Anna', email: 'a@x.at' },
            { role: 'IT-Support', name: 'Ben', email: 'b@x.at' }
        ];
        const out = movePersonToRole(rows, 'a@x.at', 'IT-Support', 'IT');
        expect(out.moved).toBe(true);
        expect(out.rows[0].role).toBe('IT-Support');
        expect(out.rows[1].role).toBe('IT-Support');
    });

    it('moves node under another parent without cycles', () => {
        const layout = seedOrganigrammFromRoles([{ code: 'SEK', name: 'Sekretariat', tier: 'verwaltung' }]);
        const role = layout.nodes.find((n) => n.type === ORG_NODE_ROLE);
        const lead = layout.nodes.find((n) => n.tone === 'lead');
        const next = moveNode(layout, role.id, lead.id);
        const moved = next.nodes.find((n) => n.id === role.id);
        expect(moved.parentId).toBe(lead.id);
        const cyclic = moveNode(next, lead.id, role.id);
        expect(cyclic.nodes.find((n) => n.id === lead.id).parentId).toBe(null);
    });
});
