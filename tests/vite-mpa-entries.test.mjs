import { readdirSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import viteConfig from '../vite.config.js';

describe('vite MPA entries', () => {
    it('nimmt jede Tool-HTML in den GitHub-Pages-Build auf', async () => {
        const resolved =
            typeof viteConfig === 'function'
                ? await viteConfig({ command: 'build', mode: 'production' })
                : viteConfig;
        const input = resolved.build.rollupOptions.input;
        const inputPaths = Object.values(input).map((p) => String(p).replace(/\\/g, '/'));

        for (const dir of ['tools', 'tools/archiv']) {
            for (const name of readdirSync(dir)) {
                if (!name.endsWith('.html')) continue;
                const needle = `/${dir}/${name}`;
                const hit = inputPaths.some((p) => p.endsWith(needle) || p.endsWith(`${dir}/${name}`));
                expect(hit, `${dir}/${name} fehlt in vite.config.js`).toBe(true);
            }
        }
    });
});
