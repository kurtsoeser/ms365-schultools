import { execSync } from 'node:child_process';
import { describe, expect, it } from 'vitest';

describe('check-tracked-secrets', () => {
    it('ms365-config.js hat keinen eingetragenen functionKey', () => {
        expect(() => {
            execSync('node scripts/check-tracked-secrets.mjs', {
                cwd: process.cwd(),
                stdio: 'pipe'
            });
        }).not.toThrow();
    });
});
