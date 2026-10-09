import { describe, expect, it } from 'vitest';
import { parseStudentsTableAoa } from '../src/tools/schueler-lifecycle/schueler-lifecycle-import.js';

describe('schueler-lifecycle-import', () => {
    it('parseStudentsTableAoa erkennt einfache Header-Tabelle', () => {
        const aoa = [
            ['Name', 'E-Mail', 'Klasse'],
            ['Ada Muster', 'ada@s.at', '1A'],
            ['Ben Test', 'ben@s.at', '2B']
        ];
        const rows = parseStudentsTableAoa(aoa);
        expect(rows).toHaveLength(2);
        expect(rows[0]).toEqual({ name: 'Ada Muster', email: 'ada@s.at', klasse: '1A' });
    });

    it('parseStudentsTableAoa mit Vor-/Nachname', () => {
        const aoa = [
            ['Vorname', 'Nachname', 'Mail', 'Klasse'],
            ['Ada', 'Muster', 'ada@s.at', '3C']
        ];
        const rows = parseStudentsTableAoa(aoa);
        expect(rows[0].name).toBe('Ada Muster');
        expect(rows[0].email).toBe('ada@s.at');
    });
});
