/**
 * Spielwiesen- / Demo-Team Naming + Lehrer-Bulk + Demo-Schüler-Pool.
 */

export const SPIEL_PREFIX = 'spiel';
export const DEMO_CLASS_CODE = 'DEMO';
export const MAX_DEMO_STUDENTS = 10;
export const DEMO_STUDENT_UPN_PREFIX = 'demo.schueler';

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
        'Schilf-/Spielwiesen-Kursteam (EDU_Class). Class Notebook und Aufgaben bewusst freigeben. ' +
        'Angelegt mit MS365-Schulverwaltung. Demo-Klasse: ' +
        DEMO_CLASS_CODE +
        '.';
    return {
        ok: issues.length === 0,
        issues,
        displayName,
        mailNickname,
        description,
        asTeam: true,
        educationClass: true,
        mode: 'single',
        notebookChecklist: defaultNotebookChecklist()
    };
}

function defaultNotebookChecklist() {
    return [
        'Team(s) anlegen (diese Seite) – ohne echte Live-Klassen.',
        'In Teams: Klassennotizbuch hinzufügen (Register „Klassennotizbuch“ / OneNote Class Notebook).',
        'Abschnitte: Inhaltsbibliothek, Zusammenarbeit, Lernermappe – Vorlagen erst nach Schilf.',
        'Aufgaben/Assignments erst nach Freigabe – sonst landen Probe-Aufgaben bei Live-Klassen.',
        'Teamcode oder Einladungslink nur an Fortbildungsgruppe / betroffene Lehrkräfte geben.',
        'Nach Schilf: Team archivieren oder in „DEMO/SPIEL“ belassen – nicht umbenennen auf echte Klasse.'
    ];
}

/**
 * Ein Lehrer-Spielwiesen-Team (1:1).
 * @param {{ code?: string, name?: string, email?: string, year?: string|number, asDemo?: boolean }} input
 */
export function buildTeacherSpielPlan(input) {
    const code = norm(input && input.code).toUpperCase();
    const name = norm(input && input.name);
    const email = norm(input && input.email).toLowerCase();
    const year = norm(input && input.year) || String(new Date().getFullYear());
    const asDemo = input && input.asDemo !== false;
    const issues = [];
    if (!code) issues.push('Lehrerkürzel fehlt');
    if (!email) issues.push('Lehrer-E-Mail fehlt');
    const label = code + (name ? ' ' + name : '');
    const displayName = (asDemo ? 'DEMO ' : 'Spielwiese ') + code + ' ' + year;
    const mailNickname = (SPIEL_PREFIX + '-' + year + '-' + slugify(code)).slice(0, 64);
    return {
        ok: issues.length === 0,
        issues,
        code,
        name,
        email,
        year,
        displayName,
        mailNickname,
        description:
            'Lehrer-Spielwiesen-Kursteam (EDU_Class) für ' +
            (name || code) +
            '. Gemeinsame Demo-Schüler der Klasse ' +
            DEMO_CLASS_CODE +
            '. MS365-Schulverwaltung.',
        ownerEmail: email,
        educationClass: true,
        mode: 'teacher'
    };
}

/**
 * @param {{ teachers?: Array<{ code?: string, name?: string, email?: string }>, year?: string|number, asDemo?: boolean, selectedCodes?: string[] }} input
 */
export function buildBulkTeacherPlans(input) {
    const year = norm(input && input.year) || String(new Date().getFullYear());
    const asDemo = input && input.asDemo !== false;
    const teachers = Array.isArray(input && input.teachers) ? input.teachers : [];
    const selected = Array.isArray(input && input.selectedCodes)
        ? input.selectedCodes.map(function (c) {
              return norm(c).toUpperCase();
          })
        : null;
    const selectedSet = selected && selected.length ? new Set(selected) : null;
    const plans = [];
    const issues = [];
    teachers.forEach(function (t) {
        const code = norm(t && t.code).toUpperCase();
        if (!code) return;
        if (selectedSet && !selectedSet.has(code)) return;
        const p = buildTeacherSpielPlan({
            code: code,
            name: t && t.name,
            email: t && t.email,
            year: year,
            asDemo: asDemo
        });
        plans.push(p);
        if (!p.ok) issues.push(code + ': ' + p.issues.join(', '));
    });
    if (!plans.length) issues.push('Keine Lehrkräfte ausgewählt');
    return {
        ok: issues.length === 0 && plans.every(function (p) {
            return p.ok;
        }),
        issues,
        year,
        plans,
        notebookChecklist: defaultNotebookChecklist()
    };
}

/**
 * Plan für Demo-Schüler Nr. 1…10.
 * @param {{ index: number, domain?: string }} input
 */
export function buildDemoStudentPlan(input) {
    const index = Math.max(1, Math.min(MAX_DEMO_STUDENTS, Number(input && input.index) || 1));
    const domain = norm(input && input.domain)
        .replace(/^@+/, '')
        .toLowerCase();
    const pad = index < 10 ? '0' + index : String(index);
    const mailNickname = 'demo-schueler' + pad;
    const upnLocal = DEMO_STUDENT_UPN_PREFIX + pad;
    const issues = [];
    if (!domain) issues.push('Schul-Domain fehlt');
    return {
        ok: issues.length === 0,
        issues,
        index: index,
        displayName: 'DEMO Schüler ' + pad,
        givenName: 'DEMO',
        surname: 'Schüler ' + pad,
        mailNickname: mailNickname,
        userPrincipalName: domain ? upnLocal + '@' + domain : '',
        department: DEMO_CLASS_CODE,
        jobTitle: 'Demo-Schüler',
        classCode: DEMO_CLASS_CODE
    };
}

/**
 * @param {{ displayName?: string, mailNickname?: string, userPrincipalName?: string, department?: string }} u
 */
export function isDemoStudentUser(u) {
    const dn = norm(u && u.displayName).toLowerCase();
    const nick = norm(u && u.mailNickname).toLowerCase();
    const upn = norm(u && u.userPrincipalName).toLowerCase();
    const dep = norm(u && u.department).toUpperCase();
    if (dep === DEMO_CLASS_CODE) return true;
    if (/^demo[\s._-]*sch(ü|ue)ler/.test(dn)) return true;
    if (nick.indexOf('demo-schueler') === 0 || nick.indexOf('demoschueler') === 0) return true;
    if (upn.indexOf(DEMO_STUDENT_UPN_PREFIX) === 0) return true;
    return false;
}

/**
 * @param {array} users
 */
export function filterDemoStudents(users) {
    return (Array.isArray(users) ? users : []).filter(isDemoStudentUser);
}

/**
 * @param {array} pool ausgewählte Demo-Schüler
 */
export function validateDemoPool(pool) {
    const list = Array.isArray(pool) ? pool : [];
    const issues = [];
    if (!list.length) issues.push('Keine Demo-Schüler im Pool');
    if (list.length > MAX_DEMO_STUDENTS) {
        issues.push('Maximal ' + MAX_DEMO_STUDENTS + ' Demo-Schüler erlaubt');
    }
    return { ok: issues.length === 0, issues, count: list.length, max: MAX_DEMO_STUDENTS };
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

/**
 * Zufälliges Startkennwort für neue Demo-Konten.
 */
export function generateDemoPassword() {
    const a = Math.random().toString(36).slice(2, 8);
    const b = Math.random().toString(36).slice(2, 6).toUpperCase();
    return 'Demo-' + a + '-' + b + '!1';
}

export default {
    SPIEL_PREFIX,
    DEMO_CLASS_CODE,
    MAX_DEMO_STUDENTS,
    DEMO_STUDENT_UPN_PREFIX,
    slugify,
    buildSpielwiesenPlan,
    buildTeacherSpielPlan,
    buildBulkTeacherPlans,
    buildDemoStudentPlan,
    isDemoStudentUser,
    filterDemoStudents,
    validateDemoPool,
    isSpielwiesenGroup,
    generateDemoPassword
};
