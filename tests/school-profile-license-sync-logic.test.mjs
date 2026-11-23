import { describe, it, expect } from 'vitest';
import {
    buildLicenseProfilePatchFromLocal,
    licenseProfilePatchIsEmpty,
    mergeLicenseIntoLocalFillEmpty,
    normalizeLicenseProfileSnapshot,
    pickLicenseSchoolStammdaten
} from '../src/shared/school-profile-license-sync-logic.js';

describe('school-profile-license-sync-logic', () => {
    it('pickLicenseSchoolStammdaten nimmt nur Sync-Felder', () => {
        expect(
            pickLicenseSchoolStammdaten({
                Strasse: 'Berggasse 1',
                Ort: 'Steyr',
                WorkshopQuelle: 'geheim',
                Telefon: '+43 1'
            })
        ).toEqual({
            Strasse: 'Berggasse 1',
            Ort: 'Steyr',
            Telefon: '+43 1'
        });
    });

    it('mergeLicenseIntoLocalFillEmpty füllt nur Leeres', () => {
        const merged = mergeLicenseIntoLocalFillEmpty(
            {
                schoolName: '',
                domain: 'schule.at',
                profile: { street: 'Lokal 1', city: '', phone: '' }
            },
            normalizeLicenseProfileSnapshot({
                schoolName: 'HAK Demo',
                primaryDomain: 'andere.at',
                schoolStammdaten: {
                    Strasse: 'Admin-Str. 9',
                    Ort: 'Linz',
                    Telefon: '+43 732',
                    Direktion: 'Frau X'
                }
            })
        );
        expect(merged.schoolName).toBe('HAK Demo');
        expect(merged.domain).toBe('schule.at');
        expect(merged.profile.street).toBe('Lokal 1');
        expect(merged.profile.city).toBe('Linz');
        expect(merged.profile.phone).toBe('+43 732');
        expect(merged.profile.direktion).toBe('Frau X');
        expect(merged.changed).toBe(true);
    });

    it('buildLicenseProfilePatchFromLocal sendet nur nicht-leere Sync-Felder', () => {
        const patch = buildLicenseProfilePatchFromLocal({
            schoolName: 'MS Test',
            profile: {
                street: 'A 1',
                postalCode: '4400',
                city: 'Steyr',
                phone: '',
                direktion: 'Dir.',
                direktionEmail: 'dir@schule.at',
                website: 'https://x.at'
            }
        });
        expect(patch.schoolName).toBe('MS Test');
        expect(patch.schoolStammdaten).toEqual({
            Strasse: 'A 1',
            PLZ: '4400',
            Ort: 'Steyr',
            Standort: 'Steyr',
            Direktion: 'Dir.',
            DirektionEmail: 'dir@schule.at'
        });
        expect(licenseProfilePatchIsEmpty(patch)).toBe(false);
        expect(licenseProfilePatchIsEmpty({})).toBe(true);
    });

    it('zieht E-Mail aus altem Direktion-Feld in direktionEmail', () => {
        const partial = mergeLicenseIntoLocalFillEmpty(
            { schoolName: 'X', domain: '', profile: {} },
            normalizeLicenseProfileSnapshot({
                schoolName: 'X',
                schoolStammdaten: { Direktion: 'leitung@schule.at', ItM365Admin: 'Max' }
            })
        );
        expect(partial.profile.direktion).toBeUndefined();
        expect(partial.profile.direktionEmail).toBe('leitung@schule.at');
        expect(partial.profile.itM365Admin).toBe('Max');
        expect(partial.profile.itM365AdminEmail).toBeUndefined();
    });
});
