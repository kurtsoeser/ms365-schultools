import { describe, expect, it, beforeEach, afterEach, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import vm from 'node:vm';

const corePath = path.join(path.dirname(fileURLToPath(import.meta.url)), '../src/shared/tenant-settings-core.js');

function loadTenantSettingsCore() {
    const code = readFileSync(corePath, 'utf8');
    const ctx = { window: {}, console };
    vm.createContext(ctx);
    vm.runInContext(code, ctx);
    return ctx.window;
}

describe('tenant-settings admin groups tier', () => {
    /** @type {ReturnType<typeof loadTenantSettingsCore>} */
    let win;

    beforeEach(() => {
        win = loadTenantSettingsCore();
    });

    afterEach(() => {
        vi.restoreAllMocks();
    });

    it('liest und schreibt 5. Feld schulleitung|verwaltung', () => {
        const parse = win.ms365TenantSettingsParseAdminGroupsLines;
        const toLines = win.ms365TenantSettingsAdminGroupsToLines;
        const text =
            'Direktion;DIREKTION;Max Direktor;dir@schule.at;schulleitung\n' +
            'Sekretariat;SEKRETARIAT;Anna Sek;sek@schule.at;verwaltung';
        const groups = parse(text);
        expect(groups.find((g) => g.name === 'Direktion').tier).toBe('schulleitung');
        expect(groups.find((g) => g.name === 'Sekretariat').tier).toBe('verwaltung');
        const round = parse(toLines(groups));
        expect(round.find((g) => g.name === 'Direktion').tier).toBe('schulleitung');
    });

    it('ohne 5. Feld: Direktion → Schulleitung, Rest Verwaltung', () => {
        const parse = win.ms365TenantSettingsParseAdminGroupsLines;
        const groups = parse(
            ['Direktion;DIR;Max;m@schule.at', 'Bibliothek;BIB;Anna;a@schule.at'].join('\n')
        );
        expect(groups.find((g) => g.name === 'Direktion').tier).toBe('schulleitung');
        expect(groups.find((g) => g.name === 'Bibliothek').tier).toBe('verwaltung');
    });
});
