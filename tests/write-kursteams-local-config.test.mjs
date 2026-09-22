import fs from 'node:fs';
import path from 'node:path';
import { execSync } from 'node:child_process';
import { describe, expect, it } from 'vitest';

describe('write-kursteams-local-config', () => {
    it('schreibt den Kursteams-Function-Key nicht nach dist', () => {
        const dist = path.join(process.cwd(), 'dist');
        const out = path.join(dist, 'ms365-config.local.js');
        const hadDist = fs.existsSync(dist);
        if (!hadDist) fs.mkdirSync(dist);
        const hadOut = fs.existsSync(out);
        const previous = hadOut ? fs.readFileSync(out, 'utf8') : null;
        if (hadOut) fs.unlinkSync(out);

        try {
            execSync('node scripts/write-kursteams-local-config.mjs', {
                cwd: process.cwd(),
                env: {
                    ...process.env,
                    KURSTEAMS_FUNCTION_KEY: 'test-key-abc',
                    LICENSE_API_BASE_URL: '',
                    LICENSE_FUNCTION_KEY: ''
                },
                stdio: 'pipe'
            });
            expect(fs.existsSync(out)).toBe(false);
        } finally {
            if (previous !== null) fs.writeFileSync(out, previous);
        }
    });

    it('schreibt die Lizenz-URL und lässt den Kursteams-Key weg', () => {
        const dist = path.join(process.cwd(), 'dist');
        const out = path.join(dist, 'ms365-config.local.js');
        const hadDist = fs.existsSync(dist);
        if (!hadDist) fs.mkdirSync(dist);
        const hadOut = fs.existsSync(out);
        const previous = hadOut ? fs.readFileSync(out, 'utf8') : null;

        try {
            execSync('node scripts/write-kursteams-local-config.mjs', {
                cwd: process.cwd(),
                env: {
                    ...process.env,
                    KURSTEAMS_FUNCTION_KEY: 'test-key-abc',
                    LICENSE_API_BASE_URL: 'https://license.example/api/license',
                    LICENSE_FUNCTION_KEY: ''
                },
                stdio: 'pipe'
            });
            const content = fs.readFileSync(out, 'utf8');
            expect(content).toContain('https://license.example/api/license');
            expect(content).toContain('MS365_LICENSE_API');
            expect(content).not.toContain('test-key-abc');
            expect(content).not.toContain('MS365_KURSTEAMS_API');
        } finally {
            if (previous !== null) fs.writeFileSync(out, previous);
            else if (fs.existsSync(out)) fs.unlinkSync(out);
        }
    });
});
