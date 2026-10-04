import { describe, it, expect } from 'vitest';
import { mapFreistellungFromItem } from '../src/tools/freistellung-planer/freistellung-planer-graph.js';

describe('freistellung-planer-graph', () => {
    it('mapFreistellungFromItem reads audit columns', () => {
        const row = mapFreistellungFromItem({
            id: '42',
            createdBy: { user: { email: 'schueler@schule.at', displayName: 'Max Muster' } },
            fields: {
                Title: 'Max Muster',
                Beginn: '2026-10-10',
                Ende: '2026-10-12',
                Status: 'Genehmigt',
                Klasse: '3AK',
                Kategorie: 'Sonstiges',
                GenehmigtVonKV: 'KV Name <kv@schule.at>',
                GenehmigtAmKV: '2026-10-08',
                GenehmigtVonDirektion: 'Dir <dir@schule.at>',
                GenehmigtAmDirektion: '2026-10-09'
            }
        });
        expect(row.genehmigtVonKv).toBe('KV Name <kv@schule.at>');
        expect(row.genehmigtAmKv).toBe('2026-10-08');
        expect(row.genehmigtVonDirektion).toContain('Dir');
        expect(row.multiDay).toBe(true);
    });
});
