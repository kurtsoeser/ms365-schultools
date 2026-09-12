/**
 * Spielwiesen- / Demo-Team Naming.
 */

export const SPIEL_PREFIX = 'spiel';

function norm(s) {
    return String(s == null ? '' : s).trim();
}

export function slugify(raw) {
    return norm(raw)
        .replace(/[äÄ]/g, 'ae')
        .replace(/[öÖ]/g, 'oe')
        .replace(/[üÜ]/g, 'ue')
        .replace(/ß/g, 'ss')
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, '-')
        .replace(/-+/g, '-')
        .replace(/^-|-$/g, '')
        .slice(0, 36);
}

/**
 * @param {{ label?: string, year?: string|number, asDemo?: boolean }} input
 */
export function buildSpielwiesenPlan(input) {
    const label = norm(input && input.label) || 'Schilf';
    const year = norm(input && input.year) || String(new Date().getFullYear());
    const asDemo = !!(input && input.asDemo);
    const issues = [];
    if (!label) issues.push('Bezeichnung fehlt');
    const displayName = (asDemo ? 'DEMO ' : 'Spielwiese ') + label + ' ' + year;
    const mailNickname = (SPIEL_PREFIX + '-' + year + '-' + slugify(label)).slice(0, 64);
    const description =
        'Schilf-/Spielwiesen-Team ohne Live-Schüler. Class Notebook und Aufgaben erst nach Schulung freigeben. ' +
        'Angelegt mit MS365-Schulverwaltung.';
    return {
        ok: issues.length === 0,
        issues,
        displayName,
        mailNickname,
        description,
        asTeam: true,
        notebookChecklist: [
            'Team ohne echte Schüler:innen anlegen (diese Seite).',
            'In Teams: Klassennotizbuch hinzufügen (Register „Klassennotizbuch“ / OneNote Class Notebook).',
            'Abschnitte: Inhaltsbibliothek, Zusammenarbeit, Lernermappe – Vorlagen erst nach Schilf.',
            'Aufgaben/Assignments erst nach Schilf freigeben – sonst landen Probe-Aufgaben bei Live-Klassen.',
            'Teamcode oder Einladungslink nur an Fortbildungsgruppe geben.',
            'Nach Schilf: Team archivieren oder in „DEMO/SPIEL“ belassen – nicht umbenennen auf echte Klasse.'
        ]
    };
}

/**
 * @param {{ displayName?: string, mailNickname?: string }} g
 */
export function isSpielwiesenGroup(g) {
    const nick = norm(g && g.mailNickname).toLowerCase();
    const dn = norm(g && g.displayName).toLowerCase();
    if (nick.indexOf(SPIEL_PREFIX + '-') === 0) return true;
    if (/^spielwiese\b/.test(dn) || /^demo\b/.test(dn)) return true;
    return false;
}

export default {
    SPIEL_PREFIX,
    slugify,
    buildSpielwiesenPlan,
    isSpielwiesenGroup
};
