import { describe, it, expect } from 'vitest';
import { mergeGraphLayers, stammdatenMembersForClass, resolveClassGraphGroupId } from '../src/tools/schulgraph/schulgraph-explore-logic.js';

describe('schulgraph-explore-logic', () => {
    it('mergeGraphLayers vereinigt Knoten und Kanten', () => {
        const merged = mergeGraphLayers(
            {
                nodes: [{ id: 'a', kind: 'class', label: 'A' }],
                edges: [{ kind: 'belongs_to', source: 'a', target: 'school:root' }]
            },
            {
                nodes: [{ id: 'g1', kind: 'm365group', label: 'Gruppe' }],
                edges: [{ kind: 'member_of', source: 'teacher:T', target: 'g1' }]
            }
        );
        expect(merged.nodes).toHaveLength(2);
        expect(merged.edges).toHaveLength(2);
    });

    it('stammdatenMembersForClass filtert nach Klasse', () => {
        const rows = stammdatenMembersForClass(
            '3LA',
            {},
            { students: [{ klasse: '3LA', name: 'Max', email: 'm@schule.at' }, { klasse: '2B', name: 'X' }] }
        );
        expect(rows).toHaveLength(1);
        expect(rows[0].email).toBe('m@schule.at');
    });

    it('resolveClassGraphGroupId aus catalogLinks', () => {
        const id = resolveClassGraphGroupId('3LA', {
            catalogLinks: [{ kind: 'class', code: '3LA', graphGroupId: 'abc-guid' }]
        });
        expect(id).toBe('abc-guid');
    });
});
