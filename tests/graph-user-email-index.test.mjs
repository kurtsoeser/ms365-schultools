import { describe, it, expect, beforeEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createContext, runInContext } from 'node:vm';

const root = dirname(fileURLToPath(import.meta.url));
const projectRoot = join(root, '..');

function loadGraph() {
    const sandbox = { console };
    sandbox.window = sandbox;
    createContext(sandbox);
    const full = join(projectRoot, 'src/shared/graph-unified-groups.js');
    runInContext(readFileSync(full, 'utf8'), sandbox, { filename: full });
    return sandbox;
}

describe('graph user email index', () => {
    let api;

    beforeEach(() => {
        api = loadGraph().ms365GraphUnifiedGroups;
    });

    it('indexUsersByEmail indexiert mail, UPN und otherMails', () => {
        const a = {
            id: '1',
            displayName: 'Anna',
            mail: 'anna@schule.at',
            userPrincipalName: 'anna@schule.onmicrosoft.com',
            otherMails: ['anna.alias@schule.at']
        };
        const b = {
            id: '2',
            displayName: 'Ben',
            mail: '',
            userPrincipalName: 'ben@schule.at'
        };
        const map = api.indexUsersByEmail([a, b]);
        expect(map.get('anna@schule.at').id).toBe('1');
        expect(map.get('anna@schule.onmicrosoft.com').id).toBe('1');
        expect(map.get('anna.alias@schule.at').id).toBe('1');
        expect(map.get('ben@schule.at').id).toBe('2');
        expect(map.has('nobody@schule.at')).toBe(false);
    });
});
