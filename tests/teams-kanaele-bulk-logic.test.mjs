import { describe, it, expect } from 'vitest';
import { loadScript } from './kursteams-vm.mjs';

describe('teams-kanaele-bulk-logic', () => {
    function logic() {
        const ctx = loadScript('src/tools/teams-kanaele-bulk/teams-kanaele-bulk-logic.js');
        return ctx.ms365TeamsKanaeleBulkLogic;
    }

    it('buildFachschaftPresetChannels liefert 01–04 mit Bindestrich', () => {
        const { buildFachschaftPresetChannels } = logic();
        const rows = buildFachschaftPresetChannels('hyphen');
        expect(rows.map((r) => r.displayName)).toEqual([
            '01-Fachkonferenz',
            '02-Curriculum',
            '03-Material',
            '04-Verschiedenes'
        ]);
    });

    it('buildFachschaftPresetChannels unterstützt Leerzeichen-Stil', () => {
        const { buildFachschaftPresetChannels } = logic();
        expect(buildFachschaftPresetChannels('space')[0].displayName).toBe('01 Fachkonferenz');
    });

    it('normalizeChannelSpecs entfernt Leerzeilen und Duplikate (case-insensitive)', () => {
        const { normalizeChannelSpecs } = logic();
        const specs = normalizeChannelSpecs([
            { displayName: '01-Fachkonferenz', description: 'A' },
            { displayName: '  ' },
            { displayName: '01-fachkonferenz', description: 'B' },
            '02-Curriculum'
        ]);
        expect(specs).toEqual([
            { displayName: '01-Fachkonferenz', description: 'A' },
            { displayName: '02-Curriculum', description: '' }
        ]);
    });

    it('filterAndSortGroupsByTerm findet Fachschaften und sortiert', () => {
        const { filterAndSortGroupsByTerm } = logic();
        const groups = [
            { id: '2', displayName: 'Fachschaft Physik', mailNickname: 'fs-phy' },
            { id: '1', displayName: 'Fachschaft Deutsch', mailNickname: 'fs-d' },
            { id: '3', displayName: 'Kursteam Mathe', mailNickname: 'kurs-m' }
        ];
        const hits = filterAndSortGroupsByTerm(groups, 'Fachschaft');
        expect(hits.map((h) => h.id)).toEqual(['1', '2']);
    });

    it('filterAndSortGroupsByTerm ohne Suchbegriff → leer', () => {
        const { filterAndSortGroupsByTerm } = logic();
        expect(filterAndSortGroupsByTerm([{ id: '1', displayName: 'X' }], '')).toEqual([]);
    });

    it('planChannelsForTeam unterscheidet create und exists', () => {
        const { planChannelsForTeam } = logic();
        const plan = planChannelsForTeam(
            [
                { displayName: '01-Fachkonferenz' },
                { displayName: '02-Curriculum' }
            ],
            [{ id: 'ch1', displayName: '01-Fachkonferenz' }]
        );
        expect(plan[0]).toMatchObject({ action: 'exists', existingId: 'ch1' });
        expect(plan[1]).toMatchObject({ action: 'create', displayName: '02-Curriculum' });
    });

    it('buildBulkChannelSummary zählt created/existed/fail', () => {
        const { buildBulkChannelSummary } = logic();
        expect(
            buildBulkChannelSummary([
                { ok: true, action: 'create' },
                { ok: true, action: 'exists', skipped: true },
                { ok: false }
            ])
        ).toEqual({ total: 3, created: 1, existed: 1, fail: 1 });
    });
});
