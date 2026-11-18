import { describe, it, expect } from 'vitest';
import {
    buildIcs,
    buildHtmlDescription,
    toWebcalUrl
} from '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-planer-export.js';
import { preferDownloadUrl } from '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-planer-feed.js';

const stammdaten = {
    classes: [{ code: '1AK', name: '1AK' }],
    teachers: [
        { code: 'MU', name: 'Müller', email: 'mueller@schule.at' },
        { code: 'HO', name: 'Huber', email: 'huber@schule.at' }
    ]
};

const sample = {
    aktivitaetId: 'akt-feed-1',
    titel: 'Museumsbesuch',
    typ: 'Exkursion',
    status: 'genehmigt',
    klasseCode: '1AK',
    klasseCodes: '1AK',
    teilnahmeModus: 'klasse',
    lehrerCode: 'MU',
    lehrerEmail: 'mueller@schule.at',
    begleitungCodes: 'HO',
    ort: 'Landesmuseum',
    startdatum: '2026-11-18',
    enddatum: '2026-11-18',
    startUe: 1,
    endUe: 4,
    unterrichtsbezug: 'Geschichte',
    verkehrsmittel: 'Bus'
};

describe('schulaktivitaeten ICS export / feed content', () => {
    it('embeds tabular details (DESCRIPTION + X-ALT-DESC) and organizer', () => {
        const ics = buildIcs([sample], stammdaten, 'Schulaktivitäten', { onlyApproved: true });
        expect(ics).toContain('X-WR-CALNAME:Schulaktivitäten');
        expect(ics).toContain('X-PUBLISHED-TTL:PT6H');
        expect(ics).toContain('SUMMARY:Exkursion: Museumsbesuch · 1AK');
        expect(ics).toContain('LOCATION:Landesmuseum');
        expect(ics).toContain('ORGANIZER;CN=Müller:mailto:mueller@schule.at');
        expect(ics).toContain('CATEGORIES:Exkursion');
        expect(ics).toContain('Verantwortlich: Müller');
        expect(ics).toContain('Begleitung: Huber');
        expect(ics).toContain('Teilnahme: 1AK');
        expect(ics).toContain('X-ALT-DESC;FMTTYPE=text/html:');
        expect(ics).toContain('<th');
        expect(ics).toContain('Begleitung');
    });

    it('onlyApproved skips beantragt/abgelehnt', () => {
        const ics = buildIcs(
            [
                sample,
                { ...sample, aktivitaetId: 'akt-2', status: 'beantragt', titel: 'Offen' },
                { ...sample, aktivitaetId: 'akt-3', status: 'abgelehnt', titel: 'Nein' }
            ],
            stammdaten,
            'Schulaktivitäten',
            { onlyApproved: true }
        );
        expect(ics).toContain('Museumsbesuch');
        expect(ics).not.toContain('SUMMARY:Exkursion: Offen');
        expect(ics).not.toContain('SUMMARY:Exkursion: Nein');
        expect((ics.match(/BEGIN:VEVENT/g) || []).length).toBe(1);
    });

    it('default export excludes only abgelehnt', () => {
        const ics = buildIcs(
            [
                sample,
                { ...sample, aktivitaetId: 'akt-2', status: 'beantragt', titel: 'Offen' },
                { ...sample, aktivitaetId: 'akt-3', status: 'abgelehnt', titel: 'Nein' }
            ],
            stammdaten
        );
        expect(ics).toContain('Museumsbesuch');
        expect(ics).toContain('Offen');
        expect(ics).not.toContain('Nein');
        expect((ics.match(/BEGIN:VEVENT/g) || []).length).toBe(2);
    });

    it('buildHtmlDescription renders table rows', () => {
        const html = buildHtmlDescription({
            titel: 'Test',
            typ: 'Exkursion',
            status: 'Genehmigt',
            teilnahme: '1AK',
            verantwortlich: 'Müller',
            begleitung: 'Huber',
            zeitraum: '18.11.2026'
        });
        expect(html).toContain('<table');
        expect(html).toContain('Verantwortlich');
        expect(html).toContain('Huber');
        expect(buildHtmlDescription({ titel: 'A < B', typ: 'x', status: 'y', teilnahme: 'z', verantwortlich: 'v', zeitraum: 't' })).toContain(
            'A &lt; B'
        );
    });

    it('toWebcalUrl and preferDownloadUrl', () => {
        expect(toWebcalUrl('https://contoso.sharepoint.com/sites/x/file.ics')).toBe(
            'webcal://contoso.sharepoint.com/sites/x/file.ics'
        );
        expect(toWebcalUrl('webcal://x')).toBe('webcal://x');
        expect(preferDownloadUrl('https://contoso.sharepoint.com/sites/x/a.ics')).toBe(
            'https://contoso.sharepoint.com/sites/x/a.ics?download=1'
        );
        expect(preferDownloadUrl('https://contoso.sharepoint.com/sites/x/a.ics?download=1')).toBe(
            'https://contoso.sharepoint.com/sites/x/a.ics?download=1'
        );
        expect(preferDownloadUrl('https://example.com/a.ics')).toBe('https://example.com/a.ics');
    });
});
