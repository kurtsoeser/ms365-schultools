import { createHash } from 'node:crypto';
import { describe, expect, it } from 'vitest';
import {
    EINSCHREIBER_HEADERS,
    payloadToEinschreiberRow,
    recordsToEinschreiberAoa
} from '../src/tools/schueler-stammdaten-portal/stammdaten-portal-export.js';
import { FIELD_PRESETS, resolveFieldProfile, visibleFields } from '../src/tools/schueler-stammdaten-portal/stammdaten-portal-fields.js';
import {
    applyGeprueftStamp,
    applyImportedStamp,
    canAdminTransition,
    canMarkGeprueft,
    canMarkImported,
    createAccessPin,
    createEnrollmentRecord,
    createUniqueGuid,
    hashPin,
    isParentEditLocked,
    normalizePayload,
    validatePayloadAgainstProfile,
    validateSvnr,
    verifyPin,
    attemptUnlock,
    STATUS
} from '../src/tools/schueler-stammdaten-portal/stammdaten-portal-logic.js';
import { buildDemoEnrollmentBlueprints } from '../src/tools/schueler-stammdaten-portal/stammdaten-portal-demo.js';

/** Gültige AT-SVNR: Stellen 1–3 Laufnr., 4 Prüfziffer, 5–10 TTMMJJ. */
function makeValidSvnr(ddmmjj = '010115') {
    const weights = [3, 7, 9, 5, 8, 4, 2, 1, 6];
    const birth = String(ddmmjj).padStart(6, '0').slice(0, 6);
    for (let serial = 100; serial < 999; serial += 1) {
        const head = String(serial).padStart(3, '0');
        const nine = [...head, ...birth].map(Number);
        let sum = 0;
        for (let i = 0; i < 9; i += 1) sum += nine[i] * weights[i];
        const rem = sum % 11;
        if (rem === 10) continue;
        return `${head}${rem}${birth}`;
    }
    throw new Error('no svnr');
}

describe('schueler-stammdaten-portal logic', () => {
    it('erzeugt GUID und 6-stelligen PIN', () => {
        expect(createUniqueGuid()).toMatch(
            /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i
        );
        expect(createAccessPin()).toMatch(/^\d{6}$/);
    });

    it('hasht und verifiziert PIN', async () => {
        const h = await hashPin('849201', 'pepper');
        expect(h).toHaveLength(64);
        expect(await verifyPin('849201', h, 'pepper')).toBe(true);
        expect(await verifyPin('000000', h, 'pepper')).toBe(false);
        // Node-Referenz: gleicher Algorithmus
        const ref = createHash('sha256').update('pepper:849201', 'utf8').digest('hex');
        expect(h).toBe(ref);
    });

    it('validiert SVNR (Prüfziffer Stelle 4)', () => {
        // Screenshot-Beispiel Lorenz: 4946300512 / Geb. 30.05.2012
        expect(validateSvnr('4946300512').ok).toBe(true);
        const svnr = makeValidSvnr('010115');
        expect(validateSvnr(svnr).ok).toBe(true);
        expect(validateSvnr('4444010115').ok).toBe(false);
        expect(validateSvnr('123').ok).toBe(false);
    });

    it('sperrt nach Fehlversuchen', async () => {
        const { record, accessPin } = await createEnrollmentRecord({ pepper: 'x' });
        let cur = record;
        for (let i = 0; i < 5; i += 1) {
            const r = await attemptUnlock(cur, '000000', { pepper: 'x', maxAttempts: 5 });
            expect(r.ok).toBe(false);
            cur = r.record;
        }
        expect(cur.lockedUntil).toBeTruthy();
        const blocked = await attemptUnlock(cur, accessPin, { pepper: 'x' });
        expect(blocked.ok).toBe(false);
        expect(blocked.code).toBe('locked');
    });

    it('kennt Status Geprüft für Sekretariats-Stempel', () => {
        expect(STATUS.GEPRUEFT).toBe('Geprüft');
        expect(STATUS.VOLLSTAENDIG).toBe('Vollständig');
    });

    it('sperrt Eltern-Schreiben nach Geprüft/Importiert', () => {
        expect(isParentEditLocked(STATUS.VOLLSTAENDIG)).toBe(false);
        expect(isParentEditLocked(STATUS.GEPRUEFT)).toBe(true);
        expect(isParentEditLocked(STATUS.IMPORTIERT)).toBe(true);
        expect(canMarkGeprueft(STATUS.VOLLSTAENDIG)).toBe(true);
        expect(canMarkGeprueft(STATUS.IN_BEARBEITUNG)).toBe(false);
        expect(canMarkImported(STATUS.GEPRUEFT)).toBe(true);
        expect(canMarkImported(STATUS.VOLLSTAENDIG)).toBe(false);
    });

    it('erlaubt nur Pipeline-Statusübergänge', () => {
        expect(canAdminTransition(STATUS.VOLLSTAENDIG, STATUS.GEPRUEFT).ok).toBe(true);
        expect(canAdminTransition(STATUS.GEPRUEFT, STATUS.IMPORTIERT).ok).toBe(true);
        expect(canAdminTransition(STATUS.VORERFASST, STATUS.GEPRUEFT).ok).toBe(false);
        expect(canAdminTransition(STATUS.GEPRUEFT, STATUS.VOLLSTAENDIG).ok).toBe(false);
        expect(canAdminTransition(STATUS.IMPORTIERT, STATUS.GEPRUEFT).ok).toBe(false);
    });

    it('setzt Ownership-Stempel beim Prüfen', () => {
        const base = {
            uniqueGuid: 'x',
            status: STATUS.VOLLSTAENDIG,
            title: 'Test',
            payload: {}
        };
        const stamped = applyGeprueftStamp(base, {
            actorEmail: 'sek@schule.at',
            actorName: 'Sekretariat',
            now: '2026-03-09T10:00:00.000Z'
        });
        expect(stamped.ok).toBe(true);
        expect(stamped.record.status).toBe(STATUS.GEPRUEFT);
        expect(stamped.record.reviewedBy).toBe('sek@schule.at');
        expect(stamped.record.reviewedAt).toBe('2026-03-09T10:00:00.000Z');
        expect(applyGeprueftStamp({ ...base, status: STATUS.GEPRUEFT }).ok).toBe(false);
        const imported = applyImportedStamp(stamped.record);
        expect(imported.ok).toBe(true);
        expect(imported.record.status).toBe(STATUS.IMPORTIERT);
        expect(applyImportedStamp(base).ok).toBe(false);
    });

    it('prüft Pflichtfelder nur beim Absenden', () => {
        const profile = resolveFieldProfile('minimal');
        const empty = normalizePayload({});
        expect(validatePayloadAgainstProfile(empty, profile, { mode: 'draft' }).ok).toBe(true);
        const submit = validatePayloadAgainstProfile(empty, profile, { mode: 'submit' });
        expect(submit.ok).toBe(false);
        expect(submit.errors.length).toBeGreaterThan(0);
    });
});

describe('feldprofile', () => {
    it('Kern zeigt weniger als Erweitert', () => {
        const kern = visibleFields(resolveFieldProfile('kern'));
        const erw = visibleFields(resolveFieldProfile('erweitert'));
        expect(kern.length).toBeLessThan(erw.length);
        expect(kern.length).toBeGreaterThan(10);
        expect(FIELD_PRESETS.kern['schueler.svnr']).toBe('pflicht');
    });
});

describe('demo local-first', () => {
    it('liefert opt-in Blueprints ohne Produktiv-Zwang', () => {
        const rows = buildDemoEnrollmentBlueprints();
        expect(rows.length).toBeGreaterThanOrEqual(3);
        expect(rows.some((r) => r.status === STATUS.VOLLSTAENDIG)).toBe(true);
        expect(rows.some((r) => r.status === STATUS.GEPRUEFT)).toBe(true);
        expect(rows.every((r) => r.payload?.schueler?.familienname)).toBe(true);
    });
});

describe('einschreiber export', () => {
    it('hat 67 Spalten und mappt Kernfelder', () => {
        expect(EINSCHREIBER_HEADERS).toHaveLength(67);
        const svnr = makeValidSvnr();
        const row = payloadToEinschreiberRow({
            schueler: {
                familienname: 'Muster',
                vorname: 'Max',
                geschlecht: 'männlich',
                geburtsdatum: '2000-12-01',
                svnr
            },
            wohnadresse: { plz: '1020', ort: 'Wien', strasse: 'Alpenstraße', hausnummer: '2', land: 'A' },
            erziehungsberechtigte: [
                { adressart: 'Mutter', vorname: 'Heidemarie', familienname: 'Mustermann', mail: 'a@b.at' },
                { adressart: 'Vater', vorname: 'Peter', familienname: 'Mustermann' }
            ],
            sameAddressAsStudent: true
        });
        expect(row).toHaveLength(67);
        expect(row[0]).toBe('Muster');
        expect(row[1]).toBe('Max');
        expect(row[4]).toBe('01.12.2000');
        expect(row[12]).toBe(svnr);
        expect(row[21]).toBe('Mutter');
        expect(row[44]).toBe('Vater');
        const aoa = recordsToEinschreiberAoa([{ payload: { schueler: { familienname: 'A', vorname: 'B' } } }]);
        expect(aoa[0]).toEqual(EINSCHREIBER_HEADERS);
        expect(aoa).toHaveLength(2);
    });
});
