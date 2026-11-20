import { describe, expect, it } from 'vitest';
import {
    normalizeEntraGroupEntries,
    mergeLegacyEntraGroup,
    formatEntraGroupsSummary
} from '../src/shared/entra-group-normalize.js';
import { normalizeSrdpPermissionsConfig } from '../src/tools/sharepoint/sharepoint-liste-srdp-permissions-ui.js';

const G1 = '11111111-1111-1111-1111-111111111111';
const G2 = '22222222-2222-2222-2222-222222222222';

describe('entra-group-normalize', () => {
    it('dedupliziert Gruppen und verwirft ungültige IDs', () => {
        expect(
            normalizeEntraGroupEntries([
                { groupId: G1, groupLabel: 'A' },
                { id: G1, label: 'A-dup' },
                { groupId: 'kein-guid', groupLabel: 'x' },
                { groupId: G2, groupLabel: 'B' }
            ])
        ).toEqual([
            { groupId: G1, groupLabel: 'A' },
            { groupId: G2, groupLabel: 'B' }
        ]);
    });

    it('merged Legacy-Einzelgruppe', () => {
        expect(mergeLegacyEntraGroup([{ groupId: G2, groupLabel: 'B' }], G1, 'A')).toEqual([
            { groupId: G1, groupLabel: 'A' },
            { groupId: G2, groupLabel: 'B' }
        ]);
        expect(mergeLegacyEntraGroup([{ groupId: G1, groupLabel: 'A' }], G1, 'A-alt')).toEqual([
            { groupId: G1, groupLabel: 'A' }
        ]);
    });

    it('formatiert Kurzfassung', () => {
        expect(formatEntraGroupsSummary([{ groupId: G1, groupLabel: 'Demo' }])).toBe('Demo');
        expect(
            formatEntraGroupsSummary([
                { groupId: G1, groupLabel: 'Demo' },
                { groupId: G2, groupLabel: 'Klasse 5B' }
            ])
        ).toBe('Demo +1 Gruppe(n)');
    });
});

describe('normalizeSrdpPermissionsConfig multi groups', () => {
    it('hebt Legacy group*Id in Arrays und spiegelt Primary', () => {
        const n = normalizeSrdpPermissionsConfig({
            groupAdminId: G1,
            groupAdmin: 'Schulleitung',
            kandidatenGroups: [
                { groupId: G1, groupLabel: '5A' },
                { groupId: G2, groupLabel: '5B' }
            ]
        });
        expect(n.adminGroups).toEqual([{ groupId: G1, groupLabel: 'Schulleitung' }]);
        expect(n.groupAdminId).toBe(G1);
        expect(n.kandidatenGroups).toHaveLength(2);
        expect(n.groupKandidatenId).toBe(G1);
        expect(n.groupKandidaten).toBe('5A');
    });
});
