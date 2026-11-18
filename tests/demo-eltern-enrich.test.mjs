import { describe, expect, it } from 'vitest';
import {
    demoParentEmail,
    demoParentName,
    enrichStudentLinesWithDemoParents,
    enrichStudentWithDemoParents,
    isBrokenDemoParentName
} from '../src/shared/demo-eltern-enrich.js';

describe('demo-eltern-enrich', () => {
    it('baut stabile Demo-Eltern-Mails', () => {
        expect(demoParentEmail('schueler01.demo@ms365.schule', 'Schüler01 DEMO', 1)).toBe(
            'eltern.schueler01.demo.a@example.com'
        );
        expect(demoParentEmail('schueler01.demo@ms365.schule', 'Schüler01 DEMO', 2)).toBe(
            'eltern.schueler01.demo.b@example.com'
        );
    });

    it('liefert nie undefined in Elternnamen', () => {
        for (let i = 1; i <= 40; i++) {
            const id = String(i).padStart(2, '0');
            const email = 'schueler' + id + '.demo@ms365.schule';
            const a = demoParentName(email, 'Schüler' + id + ' DEMO', 1);
            const b = demoParentName(email, 'Schüler' + id + ' DEMO', 2);
            expect(a).not.toMatch(/undefined/i);
            expect(b).not.toMatch(/undefined/i);
            expect(a.split(' ')).toHaveLength(2);
            expect(b.split(' ')).toHaveLength(2);
            expect(isBrokenDemoParentName(a)).toBe(false);
            expect(isBrokenDemoParentName(b)).toBe(false);
        }
    });

    it('erkennt kaputte Demo-Namen', () => {
        expect(isBrokenDemoParentName('undefined undefined')).toBe(true);
        expect(isBrokenDemoParentName('Michael undefined')).toBe(true);
        expect(isBrokenDemoParentName('undefined Demo')).toBe(true);
        expect(isBrokenDemoParentName('Anna Gruber')).toBe(false);
    });

    it('repariert undefined-Namen ohne force', () => {
        const input =
            '1A;Schüler03 DEMO;schueler03.demo@ms365.schule;undefined undefined;eltern.schueler03.demo.a@example.com';
        const r = enrichStudentLinesWithDemoParents(input);
        expect(r.repaired).toBe(1);
        expect(r.text).not.toMatch(/undefined/i);
        expect(r.text).toMatch(/eltern\.schueler03\.demo\.a@example\.com/);
    });

    it('ergänzt fehlende Eltern in Zeilen', () => {
        const input =
            '1A;Schüler01 DEMO;schueler01.demo@ms365.schule\n' +
            '1A;Schüler02 DEMO;schueler02.demo@ms365.schule;Maria Gruber;schon.da@example.com';
        const r = enrichStudentLinesWithDemoParents(input);
        expect(r.enriched).toBe(1);
        expect(r.skipped).toBe(1);
        expect(r.text.split('\n')[0]).toMatch(
            /^1A;Schüler01 DEMO;schueler01\.demo@ms365\.schule;.+;eltern\.schueler01\.demo\.a@example\.com/
        );
        expect(r.text.split('\n')[1]).toContain('schon.da@example.com');
    });

    it('force überschreibt bestehende Eltern', () => {
        const row = enrichStudentWithDemoParents(
            {
                name: 'Anna',
                email: 'anna@ms365.schule',
                parentPairs: [{ name: 'Alt', email: 'alt@example.com' }]
            },
            { force: true }
        );
        expect(row.parentPairs[0].email).not.toBe('alt@example.com');
        expect(row.parentPairs[0].email).toContain('@example.com');
        expect(row.parentPairs[0].name).not.toMatch(/undefined/i);
    });
});
