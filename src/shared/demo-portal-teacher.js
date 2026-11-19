/**
 * Demo-Portal-Accounts (@ms365.schule) in lokalen Stammdaten absichern,
 * damit die Portal-Persona automatisch „Lehrkraft“ bzw. „Schüler/in“ wird
 * und Direktion als Schulleitung für Planer-Genehmigung greift.
 */

/** @type {RegExp} */
export const MS365_DEMO_TEACHER_EMAIL_RE = /^lehrer(\d+)\.demo@ms365\.schule$/i;

/** @type {RegExp} */
export const MS365_DEMO_STUDENT_EMAIL_RE = /^schueler(\d+)\.demo@ms365\.schule$/i;

/** Demo-Direktor (Genehmigung Schulaktivitäten / Freistellung). */
export const MS365_DEMO_DIREKTOR_EMAIL = 'direktor@ms365.schule';

/**
 * @param {string|null|undefined} email
 * @returns {{ email: string, code: string, name: string }|null}
 */
export function parseMs365DemoTeacherAccount(email) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!em) return null;
    const m = MS365_DEMO_TEACHER_EMAIL_RE.exec(em);
    if (!m) return null;
    const n = String(m[1] || '').padStart(2, '0');
    return {
        email: em,
        code: 'L' + n,
        name: 'Lehrer' + n + ' DEMO'
    };
}

/**
 * @param {string|null|undefined} email
 * @returns {{ email: string, name: string, klasse: string }|null}
 */
export function parseMs365DemoStudentAccount(email) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!em) return null;
    const m = MS365_DEMO_STUDENT_EMAIL_RE.exec(em);
    if (!m) return null;
    const n = String(m[1] || '').padStart(2, '0');
    return {
        email: em,
        name: 'Schüler' + n + ' DEMO',
        // Persona braucht nur E-Mail-Match; Klasse optional (Tools nutzen Demo-Klasse)
        klasse: ''
    };
}

/**
 * @param {unknown} people
 * @param {string} email
 */
export function peopleListHasEmail(people, email) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!em) return false;
    const list = Array.isArray(people) ? people : [];
    return list.some((p) => {
        const a = String((p && (p.email || p.mail || p.upn)) || '')
            .trim()
            .toLowerCase();
        return a === em;
    });
}

/** @deprecated Alias – peopleListHasEmail */
export function teachersListHasEmail(teachers, email) {
    return peopleListHasEmail(teachers, email);
}

/**
 * @param {unknown} teachers
 * @param {{ email: string, code: string, name: string }} account
 * @returns {{ teachers: object[], changed: boolean }}
 */
export function upsertDemoTeacherInTeachersList(teachers, account) {
    const list = Array.isArray(teachers) ? teachers.slice() : [];
    if (!account || !account.email) return { teachers: list, changed: false };
    if (peopleListHasEmail(list, account.email)) {
        return { teachers: list, changed: false };
    }
    list.push({
        code: account.code,
        name: account.name,
        email: account.email
    });
    return { teachers: list, changed: true };
}

/**
 * @param {unknown} students
 * @param {{ email: string, name: string, klasse?: string }} account
 * @returns {{ students: object[], changed: boolean }}
 */
export function upsertDemoStudentInStudentsList(students, account) {
    const list = Array.isArray(students) ? students.slice() : [];
    if (!account || !account.email) return { students: list, changed: false };
    if (peopleListHasEmail(list, account.email)) {
        return { students: list, changed: false };
    }
    list.push({
        name: account.name,
        email: account.email,
        klasse: String(account.klasse || '').trim()
    });
    return { students: list, changed: true };
}

function readSignedInEmail() {
    try {
        if (typeof window !== 'undefined' && typeof window.ms365AuthGetUserPrincipalName === 'function') {
            return String(window.ms365AuthGetUserPrincipalName() || '')
                .trim()
                .toLowerCase();
        }
    } catch {
        /* ignore */
    }
    return '';
}

function loadTenantSettings() {
    if (typeof window === 'undefined' || typeof window.ms365TenantSettingsLoad !== 'function') {
        return { ok: false, reason: 'no-load' };
    }
    try {
        return { ok: true, settings: window.ms365TenantSettingsLoad() || {} };
    } catch {
        return { ok: false, reason: 'load-failed' };
    }
}

function saveTenantSettings(settings) {
    if (typeof window === 'undefined' || typeof window.ms365TenantSettingsSave !== 'function') {
        return { ok: false, reason: 'no-save' };
    }
    try {
        window.ms365TenantSettingsSave(settings);
        return { ok: true };
    } catch {
        return { ok: false, reason: 'save-failed' };
    }
}

/**
 * Wenn der angemeldete Account ein ms365.schule-Demo-Lehrer ist und lokal fehlt:
 * Eintrag in die Lehrerliste schreiben (ohne Event-Sturm – Aufrufer resolved Persona danach).
 *
 * @returns {{ ensured: boolean, changed: boolean, email?: string, reason?: string, role?: string }}
 */
export function ensureSignedInMs365DemoTeacherInLocalStammdaten() {
    const email = readSignedInEmail();
    const account = parseMs365DemoTeacherAccount(email);
    if (!account) {
        return { ensured: false, changed: false, reason: 'not-demo-teacher' };
    }

    const loaded = loadTenantSettings();
    if (!loaded.ok) {
        return { ensured: false, changed: false, email, reason: loaded.reason };
    }

    const settings = loaded.settings;
    const prev = Array.isArray(settings.teachers) ? settings.teachers : [];
    const { teachers, changed } = upsertDemoTeacherInTeachersList(prev, account);
    if (!changed) {
        return { ensured: true, changed: false, email, reason: 'already-present', role: 'lehrer' };
    }

    const saved = saveTenantSettings({ ...settings, teachers });
    if (!saved.ok) {
        return { ensured: false, changed: false, email, reason: saved.reason };
    }

    return { ensured: true, changed: true, email, role: 'lehrer' };
}

/**
 * Demo-Schüler (@ms365.schule) in die lokale Schülerliste schreiben.
 *
 * @returns {{ ensured: boolean, changed: boolean, email?: string, reason?: string, role?: string }}
 */
export function ensureSignedInMs365DemoStudentInLocalStammdaten() {
    const email = readSignedInEmail();
    const account = parseMs365DemoStudentAccount(email);
    if (!account) {
        return { ensured: false, changed: false, reason: 'not-demo-student' };
    }

    const loaded = loadTenantSettings();
    if (!loaded.ok) {
        return { ensured: false, changed: false, email, reason: loaded.reason };
    }

    const settings = loaded.settings;
    const prev = Array.isArray(settings.students) ? settings.students : [];
    const { students, changed } = upsertDemoStudentInStudentsList(prev, account);
    if (!changed) {
        return { ensured: true, changed: false, email, reason: 'already-present', role: 'schueler' };
    }

    const saved = saveTenantSettings({ ...settings, students });
    if (!saved.ok) {
        return { ensured: false, changed: false, email, reason: saved.reason };
    }

    return { ensured: true, changed: true, email, role: 'schueler' };
}

/**
 * @param {unknown} adminRows
 * @param {{ email: string, name: string, role: string }} account
 * @returns {{ admin: object[], changed: boolean }}
 */
export function upsertDemoDirektorInAdminList(adminRows, account) {
    const list = Array.isArray(adminRows) ? adminRows.slice() : [];
    if (!account || !account.email) return { admin: list, changed: false };
    if (peopleListHasEmail(list, account.email)) {
        return { admin: list, changed: false };
    }
    list.push({
        role: account.role || 'Direktion',
        name: account.name || 'Direktor',
        email: account.email,
        defaultKey: 'direktion',
        tier: 'schulleitung'
    });
    return { admin: list, changed: true };
}

/**
 * Demo-Direktor in Stammdaten → Verwaltung/Schulleitung (Genehmiger in Planern).
 *
 * @returns {{ ensured: boolean, changed: boolean, email?: string, reason?: string, role?: string }}
 */
export function ensureSignedInMs365DemoDirektorInLocalStammdaten() {
    const email = readSignedInEmail();
    if (email !== MS365_DEMO_DIREKTOR_EMAIL) {
        return { ensured: false, changed: false, reason: 'not-demo-direktor' };
    }

    const loaded = loadTenantSettings();
    if (!loaded.ok) {
        return { ensured: false, changed: false, email, reason: loaded.reason };
    }

    const settings = loaded.settings;
    const prev = Array.isArray(settings.admin) ? settings.admin : [];
    const { admin, changed } = upsertDemoDirektorInAdminList(prev, {
        email,
        name: 'Direktor',
        role: 'Direktion'
    });
    if (!changed) {
        return { ensured: true, changed: false, email, reason: 'already-present', role: 'direktion' };
    }

    const saved = saveTenantSettings({ ...settings, admin });
    if (!saved.ok) {
        return { ensured: false, changed: false, email, reason: saved.reason };
    }

    return { ensured: true, changed: true, email, role: 'direktion' };
}

/**
 * Lehrkraft, Schüler oder Direktor – je nach angemeldetem Demo-Account.
 *
 * @returns {{ ensured: boolean, changed: boolean, email?: string, reason?: string, role?: string }}
 */
export function ensureSignedInMs365DemoPortalAccountInLocalStammdaten() {
    const teacher = ensureSignedInMs365DemoTeacherInLocalStammdaten();
    if (teacher.ensured || teacher.reason !== 'not-demo-teacher') return teacher;
    const student = ensureSignedInMs365DemoStudentInLocalStammdaten();
    if (student.ensured || student.reason !== 'not-demo-student') return student;
    return ensureSignedInMs365DemoDirektorInLocalStammdaten();
}
