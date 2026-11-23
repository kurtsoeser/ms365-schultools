import fs from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';

const toolsDir = path.resolve('tools');
const livePages = [
    'playbook-daten-import-verknuepfen.html',
    'playbook-schuljahresstart.html',
    'playbook-intranet.html',
    'playbook-schularbeiten-planer.html',
    'playbook-freistellungen.html',
    'playbook-kursteams.html',
    'playbook-eltern.html',
    'playbook-elternsprechtag.html',
    'cleanup-playbook.html'
];
const stubPages = [
    'playbook-schulaktivitaeten.html',
    'playbook-projektwochen.html',
    'playbook-einsammler.html',
    'playbook-elternbrief.html',
    'playbook-krankmeldung.html',
    'playbook-srdp.html',
    'playbook-wahlen.html'
];

describe('playbook-shell wave 1', () => {
    it('stellt shared CSS bereit', () => {
        const css = fs.readFileSync(path.resolve('src/shared/playbook-shell.css'), 'utf8');
        expect(css).toMatch(/\.pb-step/);
        expect(css).toMatch(/pb-step--ready-ok/);
        expect(css).toMatch(/pb-phase__head/);
    });

    it('migriert Live-Playbooks ohne Inline-cp-Styles', () => {
        livePages.forEach((file) => {
            const html = fs.readFileSync(path.join(toolsDir, file), 'utf8');
            expect(html, file).toContain('playbook-shell.css');
            expect(html, file).toContain('design-tokens.css');
            expect(html, file).toContain('shell-fluent.css');
            expect(html, file).toContain('pb-page');
            expect(html, file).toContain('pb-step');
            expect(html, file).not.toMatch(/<style>[\s\S]*\.cp-step\s*\{/);
            expect(html, file).toContain('wirePlaybookPage');
            expect(html, file).toContain('data-cp-check');
        });
    });

    it('bindet App-Playbooks an dieselbe Shell (Wave 3 ausgearbeitet)', () => {
        stubPages.forEach((file) => {
            const html = fs.readFileSync(path.join(toolsDir, file), 'utf8');
            expect(html, file).toContain('playbook-shell.css');
            expect(html, file).toContain('pb-page');
            expect(html, file).toContain('pb-step');
            expect(html, file).toContain('wirePlaybookPage');
            expect(html, file).not.toContain('pb-stub-banner');
            expect(html, file).not.toMatch(/Stub/);
        });
    });

    it('bewahrt Blaupause-Phasen und SA-Header-Actions', () => {
        const blaupause = fs.readFileSync(
            path.join(toolsDir, 'playbook-daten-import-verknuepfen.html'),
            'utf8'
        );
        expect(blaupause).toContain('pb-phases');
        expect(blaupause).toContain('pb-step__actions');
        expect(blaupause).toContain('pb-pill--must');

        const sa = fs.readFileSync(path.join(toolsDir, 'playbook-schularbeiten-planer.html'), 'utf8');
        expect(sa).toContain('playbook-intranet.html');
        expect(sa).toContain('schularbeiten-planer.html?view=regeln');
    });
});
