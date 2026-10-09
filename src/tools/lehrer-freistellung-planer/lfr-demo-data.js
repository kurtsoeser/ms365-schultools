/**
 * Demo-Daten für Lehrer-Freistellungen (ohne SharePoint).
 */
import { newAntragId } from './lfr-schema.js';

export function buildDemoItems(accountEmail, accountName) {
    const em = String(accountEmail || 'lehrer@schule.at').toLowerCase();
    const name = String(accountName || 'Demo Lehrkraft').trim();
    const y = new Date().getFullYear();
    return [
        {
            itemId: 'demo-1',
            antragId: newAntragId(),
            titel: 'Fortbildung Informatik',
            beginn: `${y}-11-12T08:00`,
            ende: `${y}-11-12T16:00`,
            status: 'Genehmigt',
            kategorie: 'Fortbildung',
            beschreibung: 'Workshop an der Pädagogischen Hochschule.',
            lehrerName: name,
            lehrerEmail: em,
            genehmigtVon: 'Direktion',
            genehmigtAm: `${y}-11-01`
        },
        {
            itemId: 'demo-2',
            antragId: newAntragId(),
            titel: 'Arzttermin',
            beginn: `${y}-12-03T10:00`,
            ende: `${y}-12-03T12:00`,
            status: 'Ausstehend',
            kategorie: 'Arzttermin',
            beschreibung: '',
            lehrerName: name,
            lehrerEmail: em
        }
    ];
}
