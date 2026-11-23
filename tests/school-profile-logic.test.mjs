import { describe, it, expect } from 'vitest';
import {
    normalizeSchoolProfile,
    emptySchoolProfile,
    formatSchoolAddressLine,
    normalizeWebsiteHref,
    normalizeSchoolLogoUrl,
    resolveSchoolLogoSrc,
    buildDashboardSchoolCard,
    splitRolePersonFields,
    SCHOOL_LOGO_DATA_URL_MAX_CHARS
} from '../src/shared/school-profile-logic.js';

describe('school-profile-logic', () => {
    it('normalisiert leeres Profil', () => {
        const p = normalizeSchoolProfile({});
        expect(p.country).toBe('Österreich');
        expect(p.logoDataUrl).toBe('');
        expect(p.logoUrl).toBe('');
    });

    it('lehnt ungültige Logo-Data-URLs ab', () => {
        const p = normalizeSchoolProfile({ logoDataUrl: 'not-an-image' });
        expect(p.logoDataUrl).toBe('');
    });

    it('lehnt zu lange Logo-Data-URLs ab (Backup/localStorage)', () => {
        const huge = 'data:image/jpeg;base64,' + 'A'.repeat(SCHOOL_LOGO_DATA_URL_MAX_CHARS);
        expect(normalizeSchoolProfile({ logoDataUrl: huge }).logoDataUrl).toBe('');
        const ok = 'data:image/png;base64,iVBORw0KGgo=';
        expect(normalizeSchoolProfile({ logoDataUrl: ok }).logoDataUrl).toBe(ok);
    });

    it('akzeptiert nur öffentliche HTTPS-Logo-URLs', () => {
        expect(normalizeSchoolLogoUrl('http://schule.at/logo.png')).toBe('');
        expect(normalizeSchoolLogoUrl('ftp://schule.at/logo.png')).toBe('');
        expect(normalizeSchoolLogoUrl('/relative/logo.png')).toBe('');
        expect(normalizeSchoolLogoUrl('javascript:alert(1)')).toBe('');
        expect(normalizeSchoolLogoUrl('https://schule.at/media/logo.png')).toBe(
            'https://schule.at/media/logo.png'
        );
    });

    it('bevorzugt HTTPS-URL vor Data-URL für die Anzeige', () => {
        const src = resolveSchoolLogoSrc({
            logoUrl: 'https://cdn.schule.at/logo.png',
            logoDataUrl: 'data:image/png;base64,aaa'
        });
        expect(src).toBe('https://cdn.schule.at/logo.png');
    });

    it('behält gültige Kontaktfelder', () => {
        const p = normalizeSchoolProfile({
            schoolCode: '302123',
            street: 'Hauptstraße 1',
            city: 'Wien',
            email: 'Office@Schule.AT',
            logoUrl: 'https://example.com/logo.webp'
        });
        expect(p.schoolCode).toBe('302123');
        expect(p.email).toBe('office@schule.at');
        expect(p.logoUrl).toBe('https://example.com/logo.webp');
        expect(emptySchoolProfile().website).toBe('');
    });

    it('trennt Rollen in Name und E-Mail', () => {
        expect(splitRolePersonFields('Kurt Söser', 'kurt@schule.at')).toEqual({
            name: 'Kurt Söser',
            email: 'kurt@schule.at'
        });
        expect(splitRolePersonFields('kurt@schule.at', '')).toEqual({
            name: '',
            email: 'kurt@schule.at'
        });
        expect(splitRolePersonFields('Kurt Söser <kurt@schule.at>', '')).toEqual({
            name: 'Kurt Söser',
            email: 'kurt@schule.at'
        });
        const p = normalizeSchoolProfile({
            direktion: 'DEMO Direktor',
            itM365Admin: 'it@ms365.schule'
        });
        expect(p.direktion).toBe('DEMO Direktor');
        expect(p.direktionEmail).toBe('');
        expect(p.itM365Admin).toBe('');
        expect(p.itM365AdminEmail).toBe('it@ms365.schule');
    });

    it('formatiert Adresse und Website für Dashboard-Karte', () => {
        expect(
            formatSchoolAddressLine({
                street: 'Schulgasse 2',
                postalCode: '1010',
                city: 'Wien'
            })
        ).toBe('Schulgasse 2, 1010 Wien');
        expect(normalizeWebsiteHref('www.schule.at')).toBe('https://www.schule.at');
        const card = buildDashboardSchoolCard({
            schoolName: 'Demo-HAK',
            domain: 'demo.at',
            schoolProfile: {
                street: 'Schulgasse 2',
                city: 'Wien',
                website: 'www.schule.at',
                logoUrl: 'https://demo.at/logo.png'
            }
        });
        expect(card.schoolName).toBe('Demo-HAK');
        expect(card.hasLogo).toBe(true);
        expect(card.logoSrc).toBe('https://demo.at/logo.png');
        expect(card.websiteHref).toBe('https://www.schule.at');
        expect(card.hasIdentity).toBe(true);
    });
});
