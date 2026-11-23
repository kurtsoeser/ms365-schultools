import assert from 'node:assert/strict';
import test from 'node:test';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const {
    pickSchoolStammdatenFromFields,
    normalizeSchoolProfileMePatch
} = require('../src/lib/school-profile-sync.js');

test('pickSchoolStammdatenFromFields', () => {
    assert.deepEqual(
        pickSchoolStammdatenFromFields({
            Title: 'Schule',
            Strasse: 'Hauptstr. 1',
            PLZ: '4560',
            Ort: 'Kirchdorf',
            Direktion: 'Trinko',
            WorkshopQuelle: 'intern'
        }),
        {
            Strasse: 'Hauptstr. 1',
            PLZ: '4560',
            Ort: 'Kirchdorf',
            Direktion: 'Trinko'
        }
    );
});

test('normalizeSchoolProfileMePatch erlaubt Sync-Felder', () => {
    const r = normalizeSchoolProfileMePatch({
        schoolName: 'HAK Kirchdorf',
        schoolStammdaten: {
            Strasse: 'x',
            Ort: 'Kirchdorf',
            WorkshopQuelle: 'soll weg',
            Telefon: '+43',
            DirektionEmail: 'dir@schule.at'
        }
    });
    assert.equal(r.ok, true);
    assert.equal(r.patch.schoolName, 'HAK Kirchdorf');
    assert.deepEqual(r.patch.extra, {
        Strasse: 'x',
        Ort: 'Kirchdorf',
        Telefon: '+43',
        DirektionEmail: 'dir@schule.at'
    });
});

test('normalizeSchoolProfileMePatch lehnt leeren Patch ab', () => {
    assert.equal(normalizeSchoolProfileMePatch({ schoolStammdaten: {} }).ok, false);
    assert.equal(normalizeSchoolProfileMePatch({ schoolName: '   ' }).reason, 'empty_school_name');
});
