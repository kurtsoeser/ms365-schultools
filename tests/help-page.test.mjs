import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createContext, runInContext } from 'node:vm';
import { describe, expect, it } from 'vitest';

const projectRoot = join(dirname(fileURLToPath(import.meta.url)), '..');

function read(rel) {
    return readFileSync(join(projectRoot, rel), 'utf8');
}

function loadHelpApi() {
    const code = read('src/shared/help-page.js');
    const sandbox = {
        console,
        document: {
            readyState: 'complete',
            body: null,
            getElementById() {
                return null;
            },
            addEventListener() {}
        }
    };
    sandbox.window = sandbox;
    const ctx = createContext(sandbox);
    runInContext(code, ctx);
    return sandbox.ms365HelpPage;
}

describe('Hilfe-Seite', () => {
    it('bewahrt die bestehenden Sprungmarken und ergänzt Suche plus Datenschutz', () => {
        const html = read('hilfe.html');
        const requiredIds = [
            'grundprinzip',
            'datenspeicher',
            'voraussetzungen',
            'tool-tenant',
            'tool-kursteams',
            'tool-jahrgang',
            'tool-arge',
            'tool-slg',
            'tool-verwaltung',
            'tool-wtg',
            'tool-personen-verwaltung',
            'tool-eltern-verteiler',
            'tool-teams-archiv',
            'tool-schuljahr',
            'tool-klassen-umbenennen',
            'tool-gruppenerstellung',
            'tool-gaeste-verwalten',
            'windows-cmd',
            'datenschutz',
            'schnellstart',
            'einrichtung',
            'faq',
            'hinweise',
            'tool-postfaecher',
            'tool-verteilerlisten',
            'tool-intranet',
            'tool-lehrerliste',
            'tool-schultermine',
            'tool-sharepoint-website',
            'tool-sharepoint-teilen',
            'tool-leere-gruppen',
            'tool-schulstruktur',
            'helpSearch',
            'helpRoot'
        ];
        requiredIds.forEach(function (id) {
            expect(html, 'fehlende id ' + id).toContain('id="' + id + '"');
        });
        expect(html).toContain('src="src/shared/help-page.js"');
        expect(html).toContain('data-help-article');
        expect(html).toContain('data-help-faq');
        expect(html).toContain('Browser-Backup');
        expect(html).toContain('Verantwortliche im Sinne der DSGVO');
        expect(html).toContain('Mark of the Web');
    });

    it('deckt die Dashboard-Werkzeuge in der Hilfe ab', () => {
        const html = read('hilfe.html');
        const index = read('index.html');
        const toolIds = [...index.matchAll(/data-tool-id="([^"]+)"/g)].map((m) => m[1]);
        expect(toolIds.length).toBeGreaterThan(10);
        const helpAnchors = {
            jahrgang: 'tool-jahrgang',
            kursteams: 'tool-kursteams',
            'arge-fachgruppen': 'tool-arge',
            'personen-verwaltung': 'tool-personen-verwaltung',
            'gaeste-verwalten': 'tool-gaeste-verwalten',
            slg: 'tool-slg',
            verwaltung: 'tool-verwaltung',
            'organisations-assistent': 'tool-schuljahr',
            postfaecher: 'tool-postfaecher',
            verteilerlisten: 'tool-verteilerlisten',
            'eltern-verteiler': 'tool-eltern-verteiler',
            'sharepoint-intranet-hub': 'tool-intranet',
            'sharepoint-liste-lehrer': 'tool-lehrerliste',
            'sharepoint-liste-schultermine': 'tool-schultermine',
            gruppenerstellung: 'tool-gruppenerstellung',
            'sharepoint-mandant-website': 'tool-sharepoint-website',
            'sharepoint-mandant-teilen': 'tool-sharepoint-teilen',
            'schulstruktur-sync': 'tool-schulstruktur',
            'leere-gruppen-report': 'tool-leere-gruppen'
        };
        toolIds.forEach(function (id) {
            const anchor = helpAnchors[id];
            expect(anchor, 'kein Hilfe-Anker für Dashboard-Tool ' + id).toBeTruthy();
            expect(html).toContain('id="' + anchor + '"');
        });
    });

    it('filtert Abschnitte über die Volltextsuche', () => {
        const api = loadHelpApi();
        expect(api.normalizeQuery('  Eltern   Verteiler ')).toBe('eltern verteiler');

        const eltern = {
            getAttribute(name) {
                return name === 'data-search' ? 'eltern verteiler gal' : '';
            },
            textContent: 'Eltern-Verteiler Exchange'
        };
        const kurse = {
            getAttribute(name) {
                return name === 'data-search' ? 'kursteam stundenplan' : '';
            },
            textContent: 'Unterrichtsteams anlegen'
        };
        expect(api.matchesQuery(eltern, 'eltern')).toBe(true);
        expect(api.matchesQuery(kurse, 'eltern')).toBe(false);
        expect(api.matchesQuery(kurse, '')).toBe(true);
    });

    it('öffnet Hilfe im selben Tab, ohne target=_blank', () => {
        const files = [
            'index.html',
            'tenant.html',
            'einrichtung.html',
            'tools/kursteams.html',
            'tools/organisations-assistent.html'
        ];
        files.forEach(function (rel) {
            const html = read(rel);
            const helpLinks = [...html.matchAll(/<a href="[^"]*hilfe\.html[^"]*"[^>]*>/g)].map((m) => m[0]);
            expect(helpLinks.length, 'kein Hilfe-Link in ' + rel).toBeGreaterThan(0);
            helpLinks.forEach(function (tag) {
                expect(tag, rel + ': ' + tag).not.toMatch(/target="_blank"/);
            });
        });
        const gate = read('src/shared/pin-gate.js');
        expect(gate).toContain('isHelpPage');
        expect(gate).toContain('hilfe\\.html');
    });
});
