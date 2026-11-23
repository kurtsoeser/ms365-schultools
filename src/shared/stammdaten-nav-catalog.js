/**
 * Stammdaten-Untermenü für IT-Leiste & Dashboard:
 * Primär = Register-Bereiche (Schulprofil, Fächer, …),
 * nachrangig = Import/Blaupause-Werkzeuge.
 */
import { buildRegisterHash } from './stammdaten-register-mode.js';

/** @typedef {{ tabBtnId: string, label: string, icon: string, group: string, groupLabel: string }} StammdatenRegisterNavItem */

/** @type {StammdatenRegisterNavItem[]} */
export const STAMMDATEN_REGISTER_NAV = [
    {
        tabBtnId: 'tabMainStammdaten',
        label: 'Schulprofil',
        icon: 'bi bi-building',
        group: 'schule',
        groupLabel: 'Schule'
    },
    {
        tabBtnId: 'tabMainVerwaltung',
        label: 'Verwaltung',
        icon: 'bi bi-building-gear',
        group: 'schule',
        groupLabel: 'Schule'
    },
    {
        tabBtnId: 'tabMainSubjects',
        label: 'Fächer',
        icon: 'bi bi-journal-text',
        group: 'unterricht',
        groupLabel: 'Unterricht'
    },
    {
        tabBtnId: 'tabMainArges',
        label: 'ARGE',
        icon: 'bi bi-diagram-3',
        group: 'unterricht',
        groupLabel: 'Unterricht'
    },
    {
        tabBtnId: 'tabMainKlassen',
        label: 'Klassen',
        icon: 'bi bi-mortarboard',
        group: 'unterricht',
        groupLabel: 'Unterricht'
    },
    {
        tabBtnId: 'tabMainKlassenlehrer',
        label: 'Lehrer je Klasse',
        icon: 'bi bi-person-badge',
        group: 'unterricht',
        groupLabel: 'Unterricht'
    },
    {
        tabBtnId: 'tabMainLehrer',
        label: 'Lehrer',
        icon: 'bi bi-person-workspace',
        group: 'personen',
        groupLabel: 'Personen'
    },
    {
        tabBtnId: 'tabMainSchueler',
        label: 'Schüler',
        icon: 'bi bi-people',
        group: 'personen',
        groupLabel: 'Personen'
    },
    {
        tabBtnId: 'tabMainSchuelervertretung',
        label: 'Schülervertretung',
        icon: 'bi bi-megaphone',
        group: 'vertretung',
        groupLabel: 'Vertretung'
    },
    {
        tabBtnId: 'tabMainSga',
        label: 'SGA',
        icon: 'bi bi-people-fill',
        group: 'vertretung',
        groupLabel: 'Vertretung'
    },
    {
        tabBtnId: 'tabMainDatenlandkarte',
        label: 'Datenlandkarte',
        icon: 'bi bi-grid-3x3-gap',
        group: 'uebersicht',
        groupLabel: 'Übersicht'
    }
];

function isTenantHref(href) {
    return /(?:^|\/)tenant\.html(?:[?#]|$)/i.test(String(href || '').trim());
}

function toolFileKey(href) {
    return String(href || '')
        .trim()
        .replace(/^\.\.\//, '')
        .replace(/^\.\//, '')
        .replace(/#.*$/, '')
        .split('?')[0]
        .toLowerCase();
}

/**
 * @returns {{ href: string, label: string, icon: string, group: string, groupLabel: string, tabBtnId: string, expert: boolean }[]}
 */
export function stammdatenRegisterNavTools() {
    return STAMMDATEN_REGISTER_NAV.map(function (item) {
        const hash = buildRegisterHash(item.tabBtnId) || 'stammdaten';
        return {
            href: 'tenant.html#' + hash,
            label: item.label,
            icon: item.icon,
            group: item.group,
            groupLabel: item.groupLabel,
            tabBtnId: item.tabBtnId,
            expert: false
        };
    });
}

/**
 * Import-/Einrichtungs-Links unter Stammdaten (nachrangig / Mehr…).
 * @param {{ href: string, label: string, icon?: string, expert?: boolean }[]} [extracted]
 * @returns {{ href: string, label: string, icon: string, expert: boolean }[]}
 */
export function stammdatenImportNavTools(extracted) {
    const fallback = [
        {
            href: 'tools/playbook-daten-import-verknuepfen.html',
            label: 'Blaupause',
            icon: 'bi bi-flag',
            expert: true
        },
        {
            href: 'tools/stammdaten-quelle-waehlen.html',
            label: 'Datenquelle wählen',
            icon: 'bi bi-signpost-split',
            expert: true
        },
        {
            href: 'tools/webuntis-stammdaten-import.html?from=tenant',
            label: 'Stammdaten aus WebUntis',
            icon: 'bi bi-box-arrow-in-down',
            expert: true
        },
        {
            href: 'tools/bildungsportal-stammdaten.html',
            label: 'Bildungsportal',
            icon: 'bi bi-building',
            expert: true
        }
    ];
    const raw = Array.isArray(extracted) && extracted.length ? extracted : fallback;
    const seen = Object.create(null);
    const out = [];
    raw.forEach(function (t) {
        if (!t || isTenantHref(t.href)) return;
        const key = toolFileKey(t.href);
        if (!key || seen[key]) return;
        seen[key] = true;
        out.push({
            href: t.href,
            label: t.label || key,
            icon: t.icon || 'bi bi-box-arrow-up-right',
            expert: true
        });
    });
    if (out.length) return out;
    return fallback.map(function (t) {
        return Object.assign({}, t, { expert: true });
    });
}

/**
 * Volle Stammdaten-Tool-Liste für Nav-Baum: Register zuerst, Import unter Mehr…
 * @param {{ href: string, label: string, icon?: string, expert?: boolean }[]} [extracted]
 */
export function mergeStammdatenWorkspaceTools(extracted) {
    return stammdatenRegisterNavTools().concat(stammdatenImportNavTools(extracted));
}
