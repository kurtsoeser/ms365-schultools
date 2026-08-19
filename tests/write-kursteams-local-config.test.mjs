import fs from 'node:fs';
import path from 'node:path';
import { execSync } from 'node:child_process';
import { describe, expect, it } from 'vitest';

describe('write-kursteams-local-config', () => {
    it('schreibt dist/ms365-config.local.js aus KURSTEAMS_FUNCTION_KEY', () => {
        const dist = path.join(process.cwd(), 'dist');
        const out = path.join(dist, 'ms365-config.local.js');
        const hadDist = fs.existsSync(dist);
        if (!hadDist) fs.mkdirSync(dist);
        const hadOut = fs.existsSync(out);
        const previous = hadOut ? fs.readFileSync(out, 'utf8') : null;

        try {
            execSync('node scripts/write-kursteams-local-config.mjs', {
                cwd: process.cwd(),
                env: { ...process.env, KURSTEAMS_FUNCTION_KEY: 'test-key-abc' },
                stdio: 'pipe'
            });
            const content = fs.readFileSync(out, 'utf8');
            expect(content).toContain('test-key-abc');
            expect(content).toContain('MS365_CONFIG_LOCAL');
        } finally {
            if (previous !== null) fs.writeFileSync(out, previous);
            else if (fs.existsSync(out)) fs.unlinkSync(out);
        }
    });
});
