import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createContext, runInContext } from 'node:vm';
import { describe, expect, it, beforeEach } from 'vitest';

const root = dirname(fileURLToPath(import.meta.url));
const projectRoot = join(root, '..');

function loadCore(store) {
    const full = join(projectRoot, 'src/shared/tenant-settings-core.js');
    const code = readFileSync(full, 'utf8');
    const sandbox = { console };
    sandbox.window = sandbox;
    sandbox.localStorage = {
        getItem(k) {
            return store.has(k) ? store.get(k) : null;
        },
        setItem(k, v) {
            store.set(k, String(v));
        },
        removeItem(k) {
            store.delete(k);
        }
    };
    createContext(sandbox);
    runInContext(code, sandbox, { filename: full });
    return sandbox;
}

describe('SGA board partition', () => {
    let store;

    beforeEach(() => {
        store = new Map();
    });

    it('teilt Zeilen in 3 Slots pro Gruppe und Overflow', () => {
        const ctx = loadCore(store);
        const rows = ctx.ms365TenantSettingsParseSgaLines(
            'Lehrer;A;A@x.at\nLehrer;B;B@x.at\nLehrer;C;C@x.at\nLehrer;D;D@x.at\nSchueler;S1;s1@x.at\nExtern;P;p@ext.at'
        );
        const { board, overflow } = ctx.ms365TenantSettingsSgaPartition(rows);
        expect(board.teachers).toHaveLength(3);
        expect(board.teachers[2].email).toBe('c@x.at');
        expect(overflow).toHaveLength(1);
        expect(overflow[0].email).toBe('d@x.at');
        expect(board.students[0].email).toBe('s1@x.at');
        expect(board.externals[0].email).toBe('p@ext.at');
    });

    it('baut Zeilen aus Board und behält Overflow', () => {
        const ctx = loadCore(store);
        const board = {
            teachers: [{ name: 'T1', email: 't1@x.at' }, { name: '', email: '' }, { name: '', email: '' }],
            students: [{ name: 'S1', email: 's1@x.at' }, { name: '', email: '' }, { name: '', email: '' }],
            externals: [{ name: 'E1', email: 'e1@x.at' }, { name: '', email: '' }, { name: '', email: '' }]
        };
        const overflow = [{ scope: 'teacher', name: 'Extra', email: 'extra@x.at' }];
        const out = ctx.ms365TenantSettingsSgaRowsFromBoard(board, overflow);
        expect(out).toHaveLength(4);
        expect(out[0].scope).toBe('teacher');
        expect(out[3].email).toBe('extra@x.at');
        expect(ctx.ms365TenantSettingsSgaScopeLabel('external')).toBe('Extern (Eltern)');
    });
});
