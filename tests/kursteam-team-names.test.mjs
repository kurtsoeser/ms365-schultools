import { describe, it, expect } from 'vitest';
import { loadScript } from './kursteams-vm.mjs';

describe('kursteam-team-names', () => {
    it('normalizePattern und buildTeamNameFromPattern', () => {
        const ctx = loadScript('src/tools/kursteams/kursteam-team-names.js');
        const { normalizePattern, buildTeamNameFromPattern, defaultTeamNamePattern } = ctx.ms365KursteamTeamNames;

        const p = normalizePattern([{ type: 'yearPrefix' }, { type: 'text', value: ' | ' }, { type: 'klasse' }]);
        expect(p.length).toBeGreaterThan(0);

        const name = buildTeamNameFromPattern(p, {
            yearPrefix: 'SJ26',
            klasse: '1AK',
            fach: 'D',
            gruppe: 'G1'
        });
        expect(name).toContain('SJ26');
        expect(name).toContain('1AK');

        const def = defaultTeamNamePattern();
        expect(Array.isArray(def)).toBe(true);
    });
});
