import { describe, it, expect } from 'vitest';
import { buildSchulGraph, neighborhood, layoutClusterGraph } from '../src/tools/schulgraph/schulgraph-logic.js';
import { aspectsForPreset, detectPreset } from '../src/tools/schulgraph/schulgraph-aspects.js';

describe('schulgraph-logic', () => {
    it('verknüpft Fach, ARGE, Klasse und KV', () => {
        const g = buildSchulGraph({
            settings: {
                schoolName: 'Demo HAK',
                classes: [{ code: '3LA', name: '3LA', headEmail: 'kv@schule.at' }],
                subjects: [{ code: 'D', name: 'Deutsch' }],
                arges: [{ code: 'FVV', name: 'FVV', subjects: ['D', 'ENG'] }],
                teachers: [{ code: 'MUEK', name: 'Müller', email: 'kv@schule.at' }]
            },
            yearBucket: {},
            setup: { catalogLinks: [{ kind: 'subject', code: 'D', displayName: 'Fach D', graphGroupId: 'g-d' }] },
            options: { aspects: aspectsForPreset('overview'), preset: 'overview', klasseFilter: '' }
        });

        expect(g.nodes.some((n) => n.id === 'class:3LA')).toBe(true);
        expect(g.nodes.some((n) => n.id === 'subject:D')).toBe(true);
        expect(g.nodes.some((n) => n.id === 'arge:FVV')).toBe(true);
        expect(g.edges.some((e) => e.kind === 'kv_of' && e.target === 'class:3LA')).toBe(true);
        expect(g.edges.some((e) => e.kind === 'subject_in_arge' && e.source === 'subject:D')).toBe(true);
        expect(g.edges.some((e) => e.kind === 'm365_link' && e.source === 'subject:D')).toBe(true);
    });

    it('Preset fach_arge ohne KV und ohne M365', () => {
        const g = buildSchulGraph({
            settings: {
                classes: [{ code: '3LA', name: '3LA', headEmail: 'kv@schule.at' }],
                subjects: [{ code: 'D', name: 'Deutsch' }],
                arges: [{ code: 'FVV', name: 'FVV', subjects: ['D'] }],
                teachers: [{ code: 'MUEK', name: 'Müller', email: 'kv@schule.at' }]
            },
            yearBucket: {},
            setup: { catalogLinks: [{ kind: 'subject', code: 'D', graphGroupId: 'g-d' }] },
            options: { aspects: aspectsForPreset('fach_arge'), preset: 'fach_arge' }
        });
        expect(g.edges.some((e) => e.kind === 'kv_of')).toBe(false);
        expect(g.edges.some((e) => e.kind === 'm365_link')).toBe(false);
        expect(g.edges.some((e) => e.kind === 'subject_in_arge')).toBe(true);
    });

    it('baut Schüler–Eltern–Klasse bei Familie-Preset', () => {
        const g = buildSchulGraph({
            settings: { classes: [{ code: '1A', name: '1A' }] },
            yearBucket: {
                students: [{ id: 's1', klasse: '1A', name: 'Anna', guardianIds: ['g1'] }],
                guardians: [{ id: 'g1', name: 'Eltern Anna', email: 'eltern@example.com' }]
            },
            setup: {},
            options: {
                aspects: aspectsForPreset('familie'),
                preset: 'familie',
                peopleAutoOffThreshold: 999
            }
        });
        expect(g.edges.some((e) => e.kind === 'in_class' && e.source === 'student:s1')).toBe(true);
        expect(g.edges.some((e) => e.kind === 'guardian_of' && e.target === 'student:s1')).toBe(true);
    });

    it('neighborhood und Layout liefern Positionen', () => {
        const g = buildSchulGraph({
            settings: {
                classes: [{ code: 'A', name: 'A' }, { code: 'B', name: 'B' }],
                subjects: [{ code: 'M', name: 'Mathe' }]
            },
            yearBucket: {},
            setup: {},
            options: { aspects: aspectsForPreset('klassen') }
        });
        const hood = neighborhood('class:A', g.edges);
        expect(hood.has('school:root')).toBe(true);
        const pos = layoutClusterGraph(g.nodes, g.edges, { width: 400, height: 300 });
        expect(pos.get('class:A')).toMatchObject({ x: expect.any(Number), y: expect.any(Number) });
    });

    it('detectPreset erkennt Gesamtüberblick', () => {
        expect(detectPreset(aspectsForPreset('overview'))).toBe('overview');
    });
});
