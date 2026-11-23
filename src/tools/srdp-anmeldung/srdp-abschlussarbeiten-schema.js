/**
 * SharePoint-Schema für ABA-Portal-Import (Abschlussarbeiten / Diplomarbeiten).
 */
import { getProfile } from './srdp-anmeldung-profiles.js';

/**
 * @param {string|number} terminJahr
 * @param {string} [profileId]
 */
export function abschlussarbeitenListTitleForYear(terminJahr, profileId) {
    const y = String(terminJahr || '').trim();
    if (!/^\d{4}$/.test(y)) {
        throw new Error('Terminjahr muss vierstellig sein (z. B. 2027).');
    }
    const profile = getProfile(profileId);
    const code = profile.listTitleCode || profile.label || 'HAK';
    return 'sRDP-Abschlussarbeiten ' + code + ' ' + y;
}

/**
 * @param {string|number} terminJahr
 * @param {string} [profileId]
 */
export function abschlussarbeitenListTitleCandidates(terminJahr, profileId) {
    const y = String(terminJahr || '').trim();
    const primary = abschlussarbeitenListTitleForYear(y, profileId);
    const profile = getProfile(profileId);
    const alts = [primary];
    if (profile.id === 'hak') {
        alts.push('sRDP-Diplomarbeiten HAK ' + y);
        alts.push('sRDP-Diplomarbeiten ' + y);
    }
    if (profile.id === 'ahs') {
        alts.push('sRP-Abschlussarbeiten AHS ' + y);
        alts.push('sRDP-ABA AHS ' + y);
    }
    return [...new Set(alts)];
}

/**
 * @returns {import('./srdp-anmeldung-schema.js').GraphColumnDef[]}
 */
export function buildAbschlussarbeitenColumns() {
    return [
        {
            name: 'Titel',
            displayName: 'Titel',
            text: { allowMultipleLines: true, maxLength: 4000 }
        },
        {
            name: 'SchuelerName',
            displayName: 'Schüler/in',
            text: { allowMultipleLines: false, maxLength: 200 }
        },
        {
            name: 'KlasseRoh',
            displayName: 'Klasse (Roh)',
            text: { allowMultipleLines: false, maxLength: 120 }
        },
        {
            name: 'KlasseCode',
            displayName: 'Klasse',
            text: { allowMultipleLines: false, maxLength: 40 }
        },
        {
            name: 'Schuljahr',
            displayName: 'Schuljahr',
            text: { allowMultipleLines: false, maxLength: 20 }
        },
        {
            name: 'Betreuer',
            displayName: 'Betreuer/in',
            text: { allowMultipleLines: false, maxLength: 200 }
        },
        {
            name: 'Status',
            displayName: 'Status',
            text: { allowMultipleLines: false, maxLength: 120 }
        },
        {
            name: 'Schule',
            displayName: 'Schule',
            text: { allowMultipleLines: false, maxLength: 200 }
        },
        {
            name: 'Quelle',
            displayName: 'Quelle',
            text: { allowMultipleLines: false, maxLength: 80 }
        },
        {
            name: 'ImportBatch',
            displayName: 'Import-Batch',
            text: { allowMultipleLines: false, maxLength: 80 }
        },
        {
            name: 'TerminJahr',
            displayName: 'Terminjahr',
            text: { allowMultipleLines: false, maxLength: 10 }
        },
        {
            name: 'Schulform',
            displayName: 'Schulform',
            text: { allowMultipleLines: false, maxLength: 20 }
        }
    ];
}
