import { describe, it, expect } from 'vitest';
import {
    readPersonField,
    enrichFreistellungItemsKvFromSiteUserList
} from '../src/tools/freistellung-planer/freistellung-planer-graph.js';

describe('freistellung KV enrich', () => {
    it('readPersonField liest KlassenvorstandEmail und LookupId', () => {
        const kv = readPersonField(
            {
                KlassenvorstandLookupId: '42',
                KlassenvorstandEmail: 'brian.may@kurtrocks.com'
            },
            'Klassenvorstand'
        );
        expect(kv.email).toBe('brian.may@kurtrocks.com');
        expect(kv.lookupId).toBe('42');
    });

    it('enrichFreistellungItemsKvFromSiteUserList setzt kvEmail aus LookupId', () => {
        const map = new Map([['42', { email: 'brian.may@kurtrocks.com', name: 'Brian May' }]]);
        const out = enrichFreistellungItemsKvFromSiteUserList(
            [{ klasse: '1A', kvLookupId: '42', kvName: '', kvEmail: '' }],
            map
        );
        expect(out[0].kvEmail).toBe('brian.may@kurtrocks.com');
        expect(out[0].kvName).toBe('Brian May');
    });
});
