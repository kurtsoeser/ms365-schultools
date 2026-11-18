/**
 * Opt-in Demo-Seed für Schüler-Stammdaten-Portal (local-first).
 * Nie still Produktiv-/Remote-Daten befüllen.
 */

import { STATUS, createEnrollmentRecord, normalizePayload } from './stammdaten-portal-logic.js';
import {
    __loadRootForDemo,
    __saveRootForDemo,
    loadPortalSettings,
    savePortalSettings
} from './stammdaten-portal-store.js';

const DEMO_TAG = 'ssp-demo-v1';

function wantsDemoFromUrl() {
    try {
        const q = new URLSearchParams(window.location.search);
        return q.get('demo') === '1' || q.get('demo') === 'true';
    } catch {
        return false;
    }
}

/**
 * @returns {Array<{ title: string, status: string, payload: object }>}
 */
export function buildDemoEnrollmentBlueprints() {
    return [
        {
            title: 'Demo, Anna (offen)',
            status: STATUS.IN_BEARBEITUNG,
            payload: normalizePayload({
                schueler: {
                    familienname: 'Demo',
                    vorname: 'Anna',
                    geschlecht: 'weiblich',
                    geburtsdatum: '2012-05-30',
                    svnr: '4946300512',
                    staatsbuergerschaft: 'Österreich'
                },
                wohnadresse: {
                    land: 'A',
                    plz: '4400',
                    ort: 'Steyr',
                    strasse: 'Demostraße',
                    hausnummer: '1'
                },
                erziehungsberechtigte: [
                    {
                        adressart: 'Mutter',
                        anrede: 'Frau',
                        vorname: 'Maria',
                        familienname: 'Demo',
                        mail: 'maria.demo@example.at',
                        mobil: '+43 664 0000001'
                    },
                    { adressart: 'Vater', anrede: 'Herr', vorname: 'Thomas', familienname: 'Demo' }
                ]
            })
        },
        {
            title: 'Beispiel, Ben (vollständig)',
            status: STATUS.VOLLSTAENDIG,
            payload: normalizePayload({
                schueler: {
                    familienname: 'Beispiel',
                    vorname: 'Ben',
                    geschlecht: 'männlich',
                    geburtsdatum: '2011-03-15',
                    geburtsort: 'Linz',
                    svnr: '1006150311',
                    staatsbuergerschaft: 'Österreich',
                    alltagssprache: 'Deutsch',
                    erstsprache: 'Deutsch'
                },
                wohnadresse: {
                    land: 'A',
                    plz: '4020',
                    ort: 'Linz',
                    strasse: 'Beispielweg',
                    hausnummer: '12a'
                },
                erziehungsberechtigte: [
                    {
                        adressart: 'Mutter',
                        anrede: 'Frau',
                        vorname: 'Clara',
                        familienname: 'Beispiel',
                        mail: 'clara.beispiel@example.at',
                        mobil: '+43 664 0000002',
                        erziehungsberechtigt: 'ja',
                        zahlungspflichtig: 'ja'
                    },
                    {
                        adressart: 'Vater',
                        anrede: 'Herr',
                        vorname: 'David',
                        familienname: 'Beispiel',
                        mail: 'david.beispiel@example.at',
                        erziehungsberechtigt: 'ja'
                    }
                ],
                sameAddressAsStudent: true
            })
        },
        {
            title: 'Probe, Pia (geprüft)',
            status: STATUS.GEPRUEFT,
            payload: normalizePayload({
                schueler: {
                    familienname: 'Probe',
                    vorname: 'Pia',
                    geschlecht: 'weiblich',
                    geburtsdatum: '2010-11-02',
                    svnr: '1004021110',
                    staatsbuergerschaft: 'Österreich'
                },
                wohnadresse: {
                    land: 'A',
                    plz: '1010',
                    ort: 'Wien',
                    strasse: 'Probegasse',
                    hausnummer: '7'
                },
                erziehungsberechtigte: [
                    {
                        adressart: 'Mutter',
                        anrede: 'Frau',
                        vorname: 'Eva',
                        familienname: 'Probe',
                        mail: 'eva.probe@example.at'
                    }
                ]
            })
        }
    ];
}

/**
 * @param {{ force?: boolean, switchToLocal?: boolean }} [opts]
 */
export async function seedLocalDemoEnrollments(opts = {}) {
    const settings = loadPortalSettings();
    if (settings.mode === 'remote' && !opts.force) {
        if (opts.switchToLocal) {
            savePortalSettings({ mode: 'local' });
        } else {
            return {
                ok: false,
                seeded: 0,
                skipped: true,
                reason:
                    'Remote-Modus aktiv – Demo füllt keine Produktivdaten. Unter Einrichtung auf „local“ wechseln oder Demo mit Bestätigung laden.'
            };
        }
    }

    const root = __loadRootForDemo();
    if (opts.switchToLocal) root.mode = 'local';

    const existingDemo = (root.records || []).filter((r) => r.demoSeed === true || r.demoTag === DEMO_TAG);
    if (existingDemo.length && !opts.force) {
        return {
            ok: true,
            seeded: 0,
            skipped: true,
            reason: 'Demo-Daten bereits vorhanden.',
            pins: []
        };
    }

    if (opts.force) {
        root.records = (root.records || []).filter((r) => r.demoSeed !== true && r.demoTag !== DEMO_TAG);
    }

    const pins = [];
    const blueprints = buildDemoEnrollmentBlueprints();
    for (const bp of blueprints) {
        const { record, accessPin } = await createEnrollmentRecord({
            title: bp.title,
            payload: bp.payload,
            pepper: root.pepper
        });
        const now = new Date().toISOString();
        const stamped = {
            ...record,
            status: bp.status,
            demoSeed: true,
            demoTag: DEMO_TAG,
            updatedAt: now
        };
        if (bp.status === STATUS.GEPRUEFT) {
            stamped.reviewedAt = now;
            stamped.reviewedBy = 'demo@local';
            stamped.reviewedByName = 'Demo Sekretariat';
        }
        root.records.unshift(stamped);
        pins.push({ id: stamped.uniqueGuid, title: stamped.title, pin: accessPin, status: stamped.status });
    }
    __saveRootForDemo(root);
    return { ok: true, seeded: pins.length, skipped: false, pins };
}

export function clearLocalDemoEnrollments() {
    const root = __loadRootForDemo();
    const before = root.records.length;
    root.records = (root.records || []).filter((r) => r.demoSeed !== true && r.demoTag !== DEMO_TAG);
    __saveRootForDemo(root);
    return { removed: before - root.records.length };
}

export function isDemoOptInRequested() {
    return wantsDemoFromUrl();
}
