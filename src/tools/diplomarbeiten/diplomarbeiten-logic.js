/**
 * Naming & Filter für Diplomarbeiten-Gruppen.
 * Standard: Display „Diplomarbeit {Jahr} – {Thema}“, Nick dipl-{jahr}-{slug}
 */

export const DIPL_PREFIX = 'dipl';

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
        .slice(0, 40);
}

/**
 * @param {{ year?: string|number, topic?: string, student?: string, mentor?: string }} input
 */
export function buildDiplomDisplayName(input) {
    const year = norm(input && input.year) || String(new Date().getFullYear());
    const topic = norm(input && input.topic) || norm(input && input.student) || 'ohne Thema';
    const mentor = norm(input && input.mentor);
    let name = 'Diplomarbeit ' + year + ' – ' + topic;
    if (mentor) name += ' (' + mentor + ')';
    return name.slice(0, 120);
}

/**
 * @param {{ year?: string|number, topic?: string, student?: string }} input
 */
export function buildDiplomMailNickname(input) {
    const year = norm(input && input.year) || String(new Date().getFullYear());
    const slug = slugify((input && (input.topic || input.student)) || 'arbeit') || 'arbeit';
    return (DIPL_PREFIX + '-' + year + '-' + slug).replace(/[^a-z0-9-]/gi, '').slice(0, 64).toLowerCase();
}

/**
 * Beschreibung mit Betreuer / Hinweis.
 */
export function buildDiplomDescription(input) {
    const bits = ['Diplomarbeit (MS365-Schul-Tools-Standard)'];
    if (input && input.mentor) bits.push('Betreuung: ' + norm(input.mentor));
    if (input && input.student) bits.push('Schüler:in: ' + norm(input.student));
    bits.push('Mail/Freigaben nur über diese Gruppe – nicht über persönliche Postfächer.');
    return bits.join(' · ');
}

/**
 * Erkennt Diplom-Gruppen an Nick oder DisplayName.
 * @param {{ displayName?: string, mailNickname?: string, mail?: string }} g
 */
export function isDiplomGroup(g) {
    const nick = norm(g && (g.mailNickname || g.mail)).toLowerCase();
    const dn = norm(g && g.displayName).toLowerCase();
    if (nick.indexOf(DIPL_PREFIX + '-') === 0) return true;
    if (/^diplomarbeit\b/.test(dn)) return true;
    if (/\bdipl[-_ ]?\d{4}\b/.test(nick) || /\bdipl[-_ ]?\d{4}\b/.test(dn)) return true;
    return false;
}

/**
 * @param {array} groups
 * @param {string} [query]
 */
export function filterDiplomGroups(groups, query) {
    const q = norm(query).toLowerCase();
    return (Array.isArray(groups) ? groups : [])
        .filter(isDiplomGroup)
        .filter(function (g) {
            if (!q) return true;
            const blob = [g.displayName, g.mailNickname, g.mail, g.description].map(norm).join(' ').toLowerCase();
            return blob.indexOf(q) !== -1;
        });
}

/**
 * Plan für Vorschau / Tests.
 */
export function buildDiplomPlan(input) {
    const year = norm(input && input.year) || String(new Date().getFullYear());
    const topic = norm(input && input.topic);
    const student = norm(input && input.student);
    const mentor = norm(input && input.mentor);
    const issues = [];
    if (!topic && !student) issues.push('Thema oder Schülername fehlt');
    const displayName = buildDiplomDisplayName({ year, topic, student, mentor });
    const mailNickname = buildDiplomMailNickname({ year, topic, student });
    const description = buildDiplomDescription({ year, topic, student, mentor });
    return {
        ok: issues.length === 0,
        issues,
        year,
        topic,
        student,
        mentor,
        displayName,
        mailNickname,
        description,
        asTeam: !!(input && input.asTeam !== false)
    };
}

export default {
    DIPL_PREFIX,
    slugify,
    buildDiplomDisplayName,
    buildDiplomMailNickname,
    buildDiplomDescription,
    isDiplomGroup,
    filterDiplomGroups,
    buildDiplomPlan
};
