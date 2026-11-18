/**
 * Demo-Daten Elternbrief (local-first, Quelle „demo“).
 */

import { draftFromVorlage } from './elternbrief-templates.js';
import {
    ELTERNBRIEF_DEMO_SEED_TAG,
    LESE_STATUS,
    STATUS,
    normalizeBrief,
    normalizeEmpfaenger,
    toIsoDateOnly
} from './elternbrief-model.js';

export { ELTERNBRIEF_DEMO_SEED_TAG };

export function buildElternbriefDemoState() {
    const in7 = (() => {
        const d = new Date();
        d.setDate(d.getDate() + 7);
        return toIsoDateOnly(d);
    })();

    const d1 = draftFromVorlage('schularbeitenplan');
    const d2 = draftFromVorlage('elternabend');

    const b1 = normalizeBrief({
        ...d1,
        status: STATUS.GESENDET,
        klasseCodes: '3A',
        faelligAm: in7,
        lesebestaetigung: true,
        gesendetAm: new Date().toISOString(),
        erstellerEmail: 'demo@schule.at',
        erstellerName: 'Demo Lehrkraft',
        absenderEmail: 'demo@schule.at',
        absenderName: 'Demo Lehrkraft',
        quelle: 'demo',
        seedTag: ELTERNBRIEF_DEMO_SEED_TAG
    });
    const b2 = normalizeBrief({
        ...d2,
        status: STATUS.ENTWURF,
        klasseCodes: '3A;2B',
        lesebestaetigung: true,
        erstellerEmail: 'demo@schule.at',
        erstellerName: 'Demo Lehrkraft',
        quelle: 'demo',
        seedTag: ELTERNBRIEF_DEMO_SEED_TAG
    });

    const empfaenger = [
        normalizeEmpfaenger({
            briefId: b1.id,
            schuelerName: 'Anna Berger',
            klasseCode: '3A',
            elternEmail: 'eltern.anna@example.com',
            status: LESE_STATUS.GELESEN,
            gelesenAm: new Date().toISOString(),
            quelle: 'demo',
            seedTag: ELTERNBRIEF_DEMO_SEED_TAG
        }),
        normalizeEmpfaenger({
            briefId: b1.id,
            schuelerName: 'Ben Hofer',
            klasseCode: '3A',
            elternEmail: 'eltern.ben@example.com',
            status: LESE_STATUS.OFFEN,
            quelle: 'demo',
            seedTag: ELTERNBRIEF_DEMO_SEED_TAG
        }),
        normalizeEmpfaenger({
            briefId: b1.id,
            schuelerName: 'Clara Mayr',
            klasseCode: '3A',
            elternEmail: 'eltern.clara@example.com',
            status: LESE_STATUS.OFFEN,
            quelle: 'demo',
            seedTag: ELTERNBRIEF_DEMO_SEED_TAG
        }),
        normalizeEmpfaenger({
            briefId: b2.id,
            schuelerName: 'David Steiner',
            klasseCode: '2B',
            elternEmail: 'eltern.david@example.com',
            status: LESE_STATUS.OFFEN,
            quelle: 'demo',
            seedTag: ELTERNBRIEF_DEMO_SEED_TAG
        })
    ];

    return {
        briefe: [b1, b2],
        empfaenger,
        updatedAt: new Date().toISOString()
    };
}
