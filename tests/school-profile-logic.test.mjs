import { describe, it, expect } from 'vitest';
import { normalizeSchoolProfile, emptySchoolProfile } from '../src/shared/school-profile-logic.js';

describe('school-profile-logic', () => {
    it('normalisiert leeres Profil', () => {
        const p = normalizeSchoolProfile({});
        expect(p.country).toBe('Österreich');
        expect(p.logoDataUrl).toBe('');
    });

    it('lehnt ungültige Logo-Data-URLs ab', () => {
        const p = normalizeSchoolProfile({ logoDataUrl: 'not-an-image' });
        expect(p.logoDataUrl).toBe('');
    });

    it('behält gültige Kontaktfelder', () => {
        const p = normalizeSchoolProfile({
            schoolCode: '302123',
            street: 'Hauptstraße 1',
            city: 'Wien',
            email: 'Office@Schule.AT'
        });
        expect(p.schoolCode).toBe('302123');
        expect(p.email).toBe('office@schule.at');
        expect(emptySchoolProfile().website).toBe('');
    });
});
