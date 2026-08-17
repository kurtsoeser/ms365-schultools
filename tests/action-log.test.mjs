import { describe, it, expect, beforeEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createContext, runInContext } from 'node:vm';

const root = dirname(fileURLToPath(import.meta.url));
const projectRoot = join(root, '..');

function load(store) {
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
    const ad = join(projectRoot, 'src/shared/app-data-v2.js');
    const al = join(projectRoot, 'src/shared/action-log.js');
    runInContext(readFileSync(ad, 'utf8'), sandbox, { filename: ad });
    runInContext(readFileSync(al, 'utf8'), sandbox, { filename: al });
    return sandbox;
}

describe('Aktionsprotokoll', () => {
    let store;

    beforeEach(() => {
        store = new Map();
    });

    it('append speichert Einträge im Setup und list liefert neueste zuerst', () => {
        const ctx = load(store);
        ctx.ms365ActionLog.append({ tool: 'graph', action: 'create-group', summary: 'Testgruppe', target: 'jg1' });
        const rows = ctx.ms365ActionLog.list();
        expect(rows.length).toBe(1);
        expect(rows[0].summary).toBe('Testgruppe');
        expect(rows[0].tool).toBe('graph');
        const setup = ctx.ms365AppDataV2.getSetup();
        expect(setup.actionLog).toHaveLength(1);
        const json = ctx.ms365ActionLog.exportJson();
        expect(json).toMatch(/Testgruppe/);
        ctx.ms365ActionLog.clear();
        expect(ctx.ms365ActionLog.list()).toHaveLength(0);
    });
});
