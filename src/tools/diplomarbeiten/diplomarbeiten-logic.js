/**
 * Naming & Filter für Diplomarbeiten / VWA.
 * Legacy: Display „Diplomarbeit {Jahr} – {Thema}“, Nick dipl-{jahr}-{slug}
 * Standard: Dach-Team pro Maturajahrgang + private Kanäle.
 */

export const DIPL_PREFIX = 'dipl';
export const VWA_PREFIX = 'vwa';

/** @typedef {'diplom'|'vwa'} ArbeitTyp */

function norm(s) {
    return String(s == null ? '' : s).trim();
}

/**
 * @param {ArbeitTyp|string} [arbeitTyp]
 */
export function labelsForArbeitTyp(arbeitTyp) {
    const typ = String(arbeitTyp || 'diplom').toLowerCase() === 'vwa' ? 'vwa' : 'diplom';
    if (typ === 'vwa') {
        return {
            typ: 'vwa',
            singular: 'Vorwissenschaftliche Arbeit',
            short: 'VWA',
            plural: 'VWA',
            toolTitle: 'VWA',
            nickPrefix: VWA_PREFIX,
            channelPrefix: 'VWA'
        };
    }
    return {
        typ: 'diplom',
        singular: 'Diplomarbeit',
        short: 'DA',
        plural: 'Diplomarbeiten',
        toolTitle: 'Diplomarbeiten',
        nickPrefix: DIPL_PREFIX,
        channelPrefix: 'DA'
    };
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
 * Dach-Team Anzeigename für Maturajahrgang.
 * @param {{ maturajahrgang?: string, arbeitTyp?: ArbeitTyp }} input
 */
export function buildCohortTeamDisplayName(input) {
    const L = labelsForArbeitTyp(input && input.arbeitTyp);
    const jahr = norm(input && input.maturajahrgang) || String(new Date().getFullYear());
    return (L.plural + ' ' + jahr).slice(0, 120);
}

/**
 * @param {{ maturajahrgang?: string, arbeitTyp?: ArbeitTyp }} input
 */
export function buildCohortTeamMailNickname(input) {
    const L = labelsForArbeitTyp(input && input.arbeitTyp);
    const jahr = slugify(norm(input && input.maturajahrgang) || String(new Date().getFullYear())).replace(
        /\//g,
        ''
    );
    return (L.nickPrefix + '-jahrgang-' + jahr).replace(/[^a-z0-9-]/gi, '').toLowerCase().slice(0, 64);
}

/**
 * Privater Kanalname (Teams: max. 50 Zeichen, keine Sonderzeichen).
 * @param {{ topic?: string, student?: string, arbeitTyp?: ArbeitTyp }} input
 */
export function buildPrivateChannelDisplayName(input) {
    const L = labelsForArbeitTyp(input && input.arbeitTyp);
    const topic = norm(input && input.topic);
    const student = norm(input && input.student);
    let name = L.channelPrefix + ' – ' + (student || topic || 'Arbeit');
    if (student && topic) name = L.channelPrefix + ' – ' + student + ' / ' + topic;
    return name
        .replace(/[\\/:*?"<>|#%]/g, '-')
        .replace(/\s+/g, ' ')
        .trim()
        .slice(0, 50);
}

/**
 * @param {{ year?: string|number, topic?: string, student?: string, mentor?: string, arbeitTyp?: ArbeitTyp }} input
 */
export function buildDiplomDisplayName(input) {
    const L = labelsForArbeitTyp(input && input.arbeitTyp);
    const year = norm(input && input.year) || String(new Date().getFullYear());
    const topic = norm(input && input.topic) || norm(input && input.student) || 'ohne Thema';
    const mentor = norm(input && input.mentor);
    let name = L.singular + ' ' + year + ' – ' + topic;
    if (mentor) name += ' (' + mentor + ')';
    return name.slice(0, 120);
}

/**
 * @param {{ year?: string|number, topic?: string, student?: string, arbeitTyp?: ArbeitTyp }} input
 */
export function buildDiplomMailNickname(input) {
    const L = labelsForArbeitTyp(input && input.arbeitTyp);
    const year = norm(input && input.year) || String(new Date().getFullYear());
    const slug = slugify((input && (input.topic || input.student)) || 'arbeit') || 'arbeit';
    return (L.nickPrefix + '-' + year + '-' + slug).replace(/[^a-z0-9-]/gi, '').slice(0, 64).toLowerCase();
}

/**
 * Beschreibung mit Betreuer / Hinweis.
 */
export function buildDiplomDescription(input) {
    const L = labelsForArbeitTyp(input && input.arbeitTyp);
    const bits = [L.singular + ' (MS365-Schul-Tools)'];
    if (input && input.mentor) bits.push('Betreuung: ' + norm(input.mentor));
    if (input && input.student) bits.push('Schüler:in: ' + norm(input.student));
    bits.push('Mail/Freigaben nur über diese Gruppe – nicht über persönliche Postfächer.');
    return bits.join(' · ');
}

/**
 * Erkennt Diplom-/VWA-Gruppen an Nick oder DisplayName.
 * @param {{ displayName?: string, mailNickname?: string, mail?: string }} g
 */
export function isDiplomGroup(g) {
    const nick = norm(g && (g.mailNickname || g.mail)).toLowerCase();
    const dn = norm(g && g.displayName).toLowerCase();
    if (nick.indexOf(DIPL_PREFIX + '-') === 0) return true;
    if (nick.indexOf(VWA_PREFIX + '-') === 0) return true;
    if (/^diplomarbeit\b/.test(dn)) return true;
    if (/^vorwissenschaftliche\b/.test(dn) || /^vwa\b/.test(dn)) return true;
    if (/\bdipl[-_ ]?\d{4}\b/.test(nick) || /\bdipl[-_ ]?\d{4}\b/.test(dn)) return true;
    if (/\bvwa[-_ ]?\d{4}\b/.test(nick) || /\bvwa[-_ ]?\d{4}\b/.test(dn)) return true;
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
 * Plan für Vorschau / Tests (Legacy-Modus).
 */
export function buildDiplomPlan(input) {
    const year = norm(input && input.year) || String(new Date().getFullYear());
    const topic = norm(input && input.topic);
    const student = norm(input && input.student);
    const mentor = norm(input && input.mentor);
    const arbeitTyp = String((input && input.arbeitTyp) || 'diplom').toLowerCase() === 'vwa' ? 'vwa' : 'diplom';
    const issues = [];
    if (!topic && !student) issues.push('Thema oder Schülername fehlt');
    const displayName = buildDiplomDisplayName({ year, topic, student, mentor, arbeitTyp });
    const mailNickname = buildDiplomMailNickname({ year, topic, student, arbeitTyp });
    const description = buildDiplomDescription({ year, topic, student, mentor, arbeitTyp });
    return {
        ok: issues.length === 0,
        issues,
        year,
        topic,
        student,
        mentor,
        arbeitTyp,
        displayName,
        mailNickname,
        description,
        asTeam: !!(input && input.asTeam !== false)
    };
}

/**
 * @param {{ topic?: string, student?: string, mentorEmail?: string, studentEmails?: string[], arbeitTyp?: ArbeitTyp }} input
 */
export function validateArbeit(input) {
    const topic = norm(input && input.topic);
    const student = norm(input && input.student);
    const errors = [];
    if (!topic && !student) errors.push('Thema oder Schülername fehlt.');
    const mails = Array.isArray(input && input.studentEmails)
        ? input.studentEmails
        : String((input && input.studentEmails) || '')
              .split(/[\n;,]+/)
              .map((s) => s.trim())
              .filter(Boolean);
    return {
        ok: errors.length === 0,
        errors,
        channelName: buildPrivateChannelDisplayName({
            topic,
            student,
            arbeitTyp: input && input.arbeitTyp
        }),
        studentEmails: mails
            .map((s) => s.toLowerCase())
            .filter((s) => s.includes('@'))
    };
}

export default {
    DIPL_PREFIX,
    VWA_PREFIX,
    slugify,
    labelsForArbeitTyp,
    buildCohortTeamDisplayName,
    buildCohortTeamMailNickname,
    buildPrivateChannelDisplayName,
    buildDiplomDisplayName,
    buildDiplomMailNickname,
    buildDiplomDescription,
    isDiplomGroup,
    filterDiplomGroups,
    buildDiplomPlan,
    validateArbeit
};
