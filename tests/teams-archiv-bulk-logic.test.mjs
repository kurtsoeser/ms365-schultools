import { describe, it, expect } from 'vitest';
import { loadScript } from './kursteams-vm.mjs';

describe('teams-archiv-bulk-logic', () => {
    it('filterAndSortGroupsByTerm findet Treffer im Anzeigenamen oder mailNickname (case-insensitive)', () => {
        const ctx = loadScript('src/tools/teams-archiv/teams-archiv-bulk-logic.js');
        const { filterAndSortGroupsByTerm } = ctx.ms365TeamsArchivBulkLogic;

        const groups = [
            { id: '1', displayName: 'SJ25 | 1A | D', mailNickname: 'sj25-1a-d' },
            { id: '2', displayName: 'SJ26 | 1A | D', mailNickname: 'sj26-1a-d' },
            { id: '3', displayName: '1B | Mathe | sj25', mailNickname: 'mathe-1b-sj25' },
            { id: '4', displayName: 'Elternverein', mailNickname: 'elternverein' }
        ];

        const hits = filterAndSortGroupsByTerm(groups, ' sj25 ');
        expect(hits.map((h) => h.id).sort()).toEqual(['1', '3']);
    });

    it('filterAndSortGroupsByTerm liefert leeres Array ohne Suchbegriff (verhindert versehentliches „alle“)', () => {
        const ctx = loadScript('src/tools/teams-archiv/teams-archiv-bulk-logic.js');
        const { filterAndSortGroupsByTerm } = ctx.ms365TeamsArchivBulkLogic;

        expect(filterAndSortGroupsByTerm([{ id: '1', displayName: 'X' }], '')).toEqual([]);
        expect(filterAndSortGroupsByTerm([{ id: '1', displayName: 'X' }], '   ')).toEqual([]);
    });

    it('filterAndSortGroupsByTerm sortiert nach Anzeigename (de)', () => {
        const ctx = loadScript('src/tools/teams-archiv/teams-archiv-bulk-logic.js');
        const { filterAndSortGroupsByTerm } = ctx.ms365TeamsArchivBulkLogic;

        const groups = [
            { id: '1', displayName: 'SJ25 | 2B | Deutsch', mailNickname: 'sj25-2b-d' },
            { id: '2', displayName: 'SJ25 | 1A | Deutsch', mailNickname: 'sj25-1a-d' }
        ];
        const hits = filterAndSortGroupsByTerm(groups, 'sj25');
        expect(hits.map((h) => h.id)).toEqual(['2', '1']);
    });

    it('groupMatchesTerm ignoriert Gruppen ohne Treffer', () => {
        const ctx = loadScript('src/tools/teams-archiv/teams-archiv-bulk-logic.js');
        const { groupMatchesTerm } = ctx.ms365TeamsArchivBulkLogic;

        expect(groupMatchesTerm({ displayName: 'SJ26 | 1A' }, 'sj25')).toBe(false);
        expect(groupMatchesTerm({ mailNickname: 'sj25-1a-d' }, 'SJ25')).toBe(true);
        expect(groupMatchesTerm(null, 'sj25')).toBe(false);
    });

    it('buildBulkActionSummary zählt ok/fail', () => {
        const ctx = loadScript('src/tools/teams-archiv/teams-archiv-bulk-logic.js');
        const { buildBulkActionSummary } = ctx.ms365TeamsArchivBulkLogic;

        const s = buildBulkActionSummary([{ ok: true }, { ok: false }, { ok: true }]);
        expect(s).toEqual({ total: 3, ok: 2, fail: 1 });
        expect(buildBulkActionSummary([])).toEqual({ total: 0, ok: 0, fail: 0 });
        expect(buildBulkActionSummary(undefined)).toEqual({ total: 0, ok: 0, fail: 0 });
    });
});
