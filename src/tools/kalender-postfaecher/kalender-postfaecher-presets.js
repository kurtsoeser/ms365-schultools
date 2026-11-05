/**
 * Vorlagen für schulische Kalender-Postfächer (Shared Mailboxes).
 */

import { suggestSharedMailboxIdentity } from '../../shared/exchange-shared-mailbox-script.js';

/** @typedef {{ id: string, label: string, aliasHint: string, description: string, publishDefault: boolean }} CalendarPreset */

/** @type {CalendarPreset[]} */
export const CALENDAR_PRESETS = [
    {
        id: 'schulkalender',
        label: 'Schulkalender',
        aliasHint: 'schulkalender',
        description: 'Öffentliche Schultermine (Elternabende, Schulfeste, Ferienhinweise).',
        publishDefault: true
    },
    {
        id: 'lehrerkalender',
        label: 'Lehrerkalender (Postfach)',
        aliasHint: 'lehrerkalender',
        description:
            'Alternative zum Gruppenkalender: freigegebenes Postfach für Kollegiums-Termine.',
        publishDefault: false
    },
    {
        id: 'schularbeiten',
        label: 'Schularbeitenkalender',
        aliasHint: 'schularbeiten-kalender',
        description: 'Schularbeiten und Prüfungstermine zentral sichtbar.',
        publishDefault: false
    },
    {
        id: 'matura',
        label: 'Matura-Kalender',
        aliasHint: 'matura-kalender',
        description: 'Termine rund um die Matura / Abschlussprüfungen.',
        publishDefault: false
    },
    {
        id: 'custom',
        label: 'Eigener Kalender',
        aliasHint: '',
        description: 'Frei benennbares Kalender-Postfach.',
        publishDefault: false
    }
];

/**
 * @param {string} presetId
 * @returns {CalendarPreset | null}
 */
export function getCalendarPreset(presetId) {
    const id = String(presetId || '').trim();
    return CALENDAR_PRESETS.find((p) => p.id === id) || null;
}

/**
 * @param {string} displayName
 * @returns {string}
 */
export function slugAliasFromName(displayName) {
    return suggestSharedMailboxIdentity(displayName).alias;
}

/**
 * @param {string} presetId
 * @param {string} [domain]
 * @param {{ displayName?: string }} [overrides]
 */
export function buildCalendarFromPreset(presetId, domain, overrides) {
    const preset = getCalendarPreset(presetId) || getCalendarPreset('custom');
    const o = overrides && typeof overrides === 'object' ? overrides : {};
    const displayName = String(o.displayName || (preset && preset.label) || 'Schulkalender').trim();
    const aliasHint =
        preset && preset.aliasHint
            ? preset.aliasHint
            : slugAliasFromName(displayName);
    const sug = suggestSharedMailboxIdentity(aliasHint || displayName, domain);
    return {
        id: 'cal-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7),
        presetId: preset ? preset.id : 'custom',
        displayName: displayName,
        alias: sug.alias,
        primarySmtp: sug.primarySmtp,
        description: preset ? preset.description : '',
        mode: /** @type {'create'} */ ('create'),
        editors: /** @type {string[]} */ ([]),
        readers: /** @type {string[]} */ ([]),
        sendAs: true,
        publish: preset ? !!preset.publishDefault : false,
        publishUrlNote: ''
    };
}

/**
 * Heuristik: Name/Alias/UPN wirkt wie Kalender-Postfach.
 * @param {{ name?: string, displayName?: string, upn?: string, mail?: string, alias?: string }} row
 */
export function looksLikeCalendarMailbox(row) {
    const blob = [
        row && row.name,
        row && row.displayName,
        row && row.upn,
        row && row.mail,
        row && row.alias
    ]
        .map((x) => String(x || '').toLowerCase())
        .join(' ');
    return (
        blob.indexOf('kalender') !== -1 ||
        blob.indexOf('calendar') !== -1 ||
        blob.indexOf('termine') !== -1 ||
        blob.indexOf('matura') !== -1 ||
        blob.indexOf('schularbeit') !== -1
    );
}
