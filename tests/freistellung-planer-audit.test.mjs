import { describe, it, expect } from 'vitest';
import {
    formatFreistellungAuditorLabel,
    buildFreistellungStatusFields
} from '../src/tools/freistellung-planer/freistellung-planer-audit.js';

describe('freistellung-planer-audit', () => {
    it('formatFreistellungAuditorLabel', () => {
        expect(formatFreistellungAuditorLabel('Brian May', 'brian@schule.at')).toBe(
            'Brian May <brian@schule.at>'
        );
    });

    it('buildFreistellungStatusFields KV genehmigt', () => {
        const f = buildFreistellungStatusFields({
            status: 'Genehmigt',
            role: 'kv',
            actorName: 'Brian May',
            actorEmail: 'brian@schule.at',
            today: '2026-10-06'
        });
        expect(f.Status).toBe('Genehmigt');
        expect(f.GenehmigtVonKV).toBe('Brian May <brian@schule.at>');
        expect(f.GenehmigtAmKV).toBe('2026-10-06');
    });

    it('buildFreistellungStatusFields abgelehnt', () => {
        const f = buildFreistellungStatusFields({
            status: 'Abgelehnt',
            role: 'kv',
            actorName: 'X',
            actorEmail: 'x@schule.at',
            today: '2026-10-06'
        });
        expect(f.AbgelehntVon).toBe('X <x@schule.at>');
        expect(f.AbgelehntAm).toBe('2026-10-06');
    });
});
