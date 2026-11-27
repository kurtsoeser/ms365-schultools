import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
    PRIVACY_COPY,
    PRIVACY_FORBIDDEN_PHRASES
} from '../src/shared/privacy-copy.js';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');

function read(rel) {
    return readFileSync(join(root, rel), 'utf8');
}

describe('privacy-copy (Single Source of Truth)', () => {
    it('hat die Kernaussagen für Schulen', () => {
        const blob = JSON.stringify(PRIVACY_COPY);
        expect(blob).toMatch(/Schul-Tenant|Microsoft-365-Tenant/i);
        expect(blob).toMatch(/Minimal/i);
        expect(blob).toMatch(/Lizenz/i);
        expect(blob).toMatch(/keine.*Klassen|Keine.*Klassen/i);
        expect(PRIVACY_COPY.landingTeaser.points.length).toBeGreaterThanOrEqual(3);
        expect(PRIVACY_COPY.core.operatorServices.list.length).toBeGreaterThanOrEqual(3);
    });

    it('speist Hilfe und Landing über Marker', () => {
        const help = read('hilfe.html');
        const land = read('landing/datenschutz.html');
        const idx = read('landing/index.html');
        const readme = read('README.md');

        expect(help).toContain('<!-- privacy-copy:help-core:start -->');
        expect(help).toContain(PRIVACY_COPY.core.operatorServices.titleHelp);
        expect(help).toContain('keinen Anwendungsserver des Betreibers für Ihre Schul-Listen');
        expect(help).toContain('Minimalprinzip');

        expect(land).toContain('<!-- privacy-copy:landing-core:start -->');
        expect(land).toContain('<!-- privacy-copy:landing-core-rest:start -->');
        expect(land).toContain(PRIVACY_COPY.core.operatorServices.titleLanding);
        expect(land).toContain('src/shared/privacy-copy.js');

        expect(idx).toContain('<!-- privacy-copy:landing-teaser:start -->');
        expect(idx).toContain(PRIVACY_COPY.landingTeaser.title);

        expect(readme).toContain('<!-- privacy-copy:readme:start -->');
        expect(readme).toContain('Klassenlisten-Datenbank beim Betreiber');
    });

    it('enthält die alten Widerspruchs-Phrasen nicht mehr in Hilfe/Landing/README', () => {
        const surfaces = [
            read('hilfe.html'),
            read('landing/datenschutz.html'),
            read('landing/index.html'),
            read('README.md')
        ].join('\n');

        PRIVACY_FORBIDDEN_PHRASES.forEach((phrase) => {
            expect(surfaces, phrase).not.toContain(phrase);
        });
        expect(surfaces).not.toMatch(/reine Browser-App\*\* \(ohne Server\)/);
    });

    it('hält Hilfe und Landing beim Betreiber-Absatz inhaltlich deckungsgleich', () => {
        const help = read('hilfe.html');
        const land = read('landing/datenschutz.html');
        const sample = PRIVACY_COPY.core.operatorServices.list[0];
        expect(help).toContain(sample);
        expect(land).toContain(sample);
        expect(help).toContain(PRIVACY_COPY.core.summary.paragraphs[0]);
        expect(land).toContain(PRIVACY_COPY.core.summary.paragraphs[0]);
    });

    it('Trust-Seite hat Datenfluss, Persona-Scopes und Build-Nachweis', () => {
        const trust = read('landing/trust.html');
        expect(trust).toContain('id="flow"');
        expect(trust).toContain('Datenflusskarte');
        expect(trust).toContain('id="scopes"');
        expect(trust).toContain('User.Read');
        expect(trust).toContain('name="ms365-landing-build"');
        expect(trust).toContain('id="landing-last-updated"');
        expect(trust).toContain('avv-vorlage.md');
        expect(trust).toContain('sicherheitspruefung-bericht.md');
        expect(trust).toContain('LICENSE');
        expect(read('LICENSE')).toMatch(/NICHT FREI NUTZBAR|PROPRIETARY/i);
        expect(read('README.md')).toMatch(/nicht frei nutzbar/i);
        expect(read('landing/datenschutz.html')).toContain('trust.html');
        expect(read('landing/index.html')).toContain('trust.html');
    });

    it('zeigt Impressum, Datenschutz und Nutzungsbedingungen im Footer', () => {
        for (const rel of [
            'landing/index.html',
            'landing/impressum.html',
            'landing/datenschutz.html',
            'landing/nutzungsbedingungen.html',
            'landing/neuigkeiten.html',
            'landing/trust.html'
        ]) {
            const html = read(rel);
            expect(html, rel).toContain('>Impressum<');
            expect(html, rel).toContain('>Datenschutz<');
            expect(html, rel).toContain('>Nutzungsbedingungen<');
        }
        expect(read('landing/impressum.html')).toContain('29320646');
        expect(read('landing/impressum.html')).toContain('9110004879074');
        expect(read('landing/impressum.html')).not.toMatch(/Stripe|Replay Pass|Magic Link/i);
        expect(read('landing/nutzungsbedingungen.html')).toContain('kein entgeltlicher Vertrag');
    });
});

