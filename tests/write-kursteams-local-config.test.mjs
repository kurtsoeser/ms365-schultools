import fs from 'node:fs';
import path from 'node:path';
import { execSync } from 'node:child_process';
import { describe, expect, it } from 'vitest';

describe('write-kursteams-local-config (a9)', () => {
    function run(env) {
        return execSync('node scripts/write-kursteams-local-config.mjs', {
            cwd: process.cwd(),
            env: { ...process.env, ...env },
            stdio: 'pipe',
            encoding: 'utf8'
        });
    }

    it('schreibt keinen Function Key und keinen Kursteams-Key nach dist', () => {
        const dist = path.join(process.cwd(), 'dist');
        const out = path.join(dist, 'ms365-config.local.js');
        const hadDist = fs.existsSync(dist);
        if (!hadDist) fs.mkdirSync(dist);
        const hadOut = fs.existsSync(out);
        const previous = hadOut ? fs.readFileSync(out, 'utf8') : null;
        if (hadOut) fs.unlinkSync(out);

        try {
            run({
                LICENSE_API_BASE_URL: 'https://license.example/api/license',
                LICENSE_FUNCTION_KEY: 'should-never-appear',
                KURSTEAMS_FUNCTION_KEY: 'test-key-abc',
                MS365_REQUIRE_PROD_URLS: ''
            });
            const content = fs.readFileSync(out, 'utf8');
            expect(content).toContain('https://license.example/api/license');
            expect(content).toContain('MS365_LICENSE_API');
            expect(content).not.toContain('should-never-appear');
            expect(content).not.toContain('functionKey');
            expect(content).not.toContain('test-key-abc');
        } finally {
            if (previous !== null) fs.writeFileSync(out, previous);
            else if (fs.existsSync(out)) fs.unlinkSync(out);
        }
    });

    it('warnt bei -dev, bricht nur mit MS365_REQUIRE_PROD_URLS ab', () => {
        const dist = path.join(process.cwd(), 'dist');
        const out = path.join(dist, 'ms365-config.local.js');
        if (!fs.existsSync(dist)) fs.mkdirSync(dist);
        const previous = fs.existsSync(out) ? fs.readFileSync(out, 'utf8') : null;

        try {
            // Aktuell nur *-dev-Apps in Azure: Deploy darf noch schreiben (Warnung).
            run({
                LICENSE_API_BASE_URL: 'https://func-ms365-license-dev.azurewebsites.net/api/license',
                MS365_REQUIRE_PROD_URLS: '',
                MS365_DEPLOY_ENV: 'production'
            });
            expect(fs.readFileSync(out, 'utf8')).toContain('license-dev');

            expect(() =>
                run({
                    LICENSE_API_BASE_URL:
                        'https://func-ms365-license-dev.azurewebsites.net/api/license',
                    MS365_REQUIRE_PROD_URLS: 'true'
                })
            ).toThrow();
        } finally {
            if (previous !== null) fs.writeFileSync(out, previous);
            else if (fs.existsSync(out)) fs.unlinkSync(out);
        }
    });
});
