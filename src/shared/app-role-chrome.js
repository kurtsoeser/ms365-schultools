/**
 * Gemeinsame App-Rolle im Kopf jeder Schul-App.
 * Die Rolle kommt aus der Einrichtung dieser App, nicht aus der Portal-Persona.
 * Anzeige: „App · Rolle · Quelle“, sobald das Konto gelesen ist.
 */

/** @typedef {{ id: string, title: string, file: string, roles?: Record<string, string>, portalOnly?: boolean }} SchoolAppChrome */

/** @type {SchoolAppChrome[]} */
export const SCHOOL_APP_CHROME = [
    { id: 'schulaktivitaeten', title: 'Schulaktivitäten', file: 'schulaktivitaeten-planer.html', roles: { admin: 'Direktion', kv: 'Klassenvorstand', lehrer: 'Lehrkraft' } },
    { id: 'freistellung', title: 'Freistellungen Schüler', file: 'freistellung-planer.html', roles: { admin: 'Direktion', direktion: 'Direktion', kv: 'Klassenvorstand', schueler: 'Schüler/in', lehrer: 'Lehrkraft' } },
    { id: 'lehrer-freistellung', title: 'Freistellungen Lehrkräfte', file: 'lehrer-freistellung-planer.html', roles: { direktion: 'Direktion', lehrer: 'Lehrkraft' } },
    { id: 'schularbeiten', title: 'Schularbeiten', file: 'schularbeiten-planer.html', roles: { admin: 'Verwaltung', lehrer: 'Lehrkraft', schueler: 'Schüler/in' } },
    { id: 'projektwochen', title: 'Projektwochen', file: 'projektwochen.html', roles: { admin: 'Leitung', lehrer: 'Lehrkraft', schueler: 'Schüler/in' } },
    { id: 'srdp', title: 'sRDP-Anmeldung', file: 'srdp-anmeldung.html', roles: { admin: 'Direktion', lehrer: 'Lehrkraft', schueler: 'Schüler/in' } },
    { id: 'wahlfach', title: 'Wahlfach-Zuteilung', file: 'wahlfach-zuteilung.html', roles: { admin: 'Leitung', lehrer: 'Lehrkraft' } },
    { id: 'wahlfach-anmeldung', title: 'Wahlfach-Anmeldung', file: 'wahlfach-anmeldung.html', roles: { schueler: 'Schüler/in' }, portalOnly: true },
    { id: 'schulsprecherwahl', title: 'Schulsprecherwahl', file: 'schulsprecherwahl.html', portalOnly: true },
    { id: 'einsammler', title: 'Einsammler', file: 'einsammler.html', roles: { leitung: 'Leitung', lehrer: 'Lehrkraft' } },
    { id: 'elternbrief', title: 'Elternbrief', file: 'elternbrief.html', roles: { leitung: 'Leitung', lehrer: 'Lehrkraft' } },
    { id: 'krankmeldung', title: 'Krankmeldung', file: 'krankmeldung.html', roles: { leitung: 'Sekretariat', kv: 'Klassenvorstand', lehrer: 'Lehrkraft' } },
    { id: 'kollegium', title: 'Kollegiumsumfragen', file: 'kollegium-umfragen.html', roles: { leitung: 'Leitung', kv: 'Klassenvorstand', lehrer: 'Lehrkraft' } },
    { id: 'ausgaben', title: 'Ausgaben', file: 'ausgaben.html', roles: { leitung: 'Leitung', lehrer: 'Lehrkraft' } },
    { id: 'aufgaben', title: 'Aufgaben', file: 'aufgaben.html', roles: { leitung: 'Leitung', melder: 'Meldung' } },
    { id: 'technik', title: 'Technik-Meldungen', file: 'technik-meldungen.html', roles: { leitung: 'Leitung', melder: 'Meldung' } },
    { id: 'raumplan', title: 'Raumplan', file: 'raumplan.html', roles: { leitung: 'Leitung', lehrer: 'Lehrkraft' } },
    { id: 'termine', title: 'Kalender und Termine', file: 'termine-hub.html', roles: { it: 'Schul-IT', lehrer: 'Lehrkraft', schueler: 'Schüler/in', self: 'Eigenes Konto' } },
    { id: 'elternsprechtag', title: 'Elternsprechtag', file: 'elternsprechtag-bookings.html', portalOnly: true },
    { id: 'teams-antrag', title: 'Team-Antrag', file: 'teams-antrag.html', portalOnly: true },
    { id: 'diplomarbeiten', title: 'Diplomarbeiten', file: 'diplomarbeiten.html', portalOnly: true },
    { id: 'it-berechtigungen', title: 'Berechtigungen', file: 'it-berechtigungen.html', roles: { it: 'Schul-IT' } }
];

const DEFAULT_ROLE_LABELS = {
    direktion: 'Direktion',
    admin: 'Leitung',
    leitung: 'Leitung',
    kv: 'Klassenvorstand',
    lehrer: 'Lehrkraft',
    schueler: 'Schüler/in',
    melder: 'Meldung',
    it: 'Schul-IT'
};

/**
 * @param {string} source
 */
export function appRoleSourceLabel(source) {
    const s = String(source || '');
    if (!s) return '';
    if (s === 'entra') return 'Entra-Gruppe';
    if (s === 'global-admin') return 'Globaler Administrator';
    if (s === 'designated-it' || s === 'school-it') return 'Schul-IT';
    if (s === 'demo-direktor') return 'Demo-Direktor';
    if (s === 'setup-direktion') return 'Genehmiger (Setup)';
    if (s === 'stammdaten-direktion') return 'Schulleitung (Stammdaten)';
    if (s === 'stammdaten-kv') return 'Klassenvorstand (Stammdaten)';
    if (s === 'stammdaten') return 'Stammdaten';
    if (s === 'direktion-user') return 'Direktion (Person)';
    if (s === 'admin-user') return 'Verwaltung (Person)';
    if (s === 'lehrer-user') return 'Lehrkraft (Person)';
    if (s === 'schueler-user') return 'Schüler/in (Person)';
    if (s === 'legacy-admin-emails') return 'Hinterlegte Admin-Mail';
    if (s === 'operator') return 'Betreiber';
    if (s === 'demo') return 'IT-Vorschau';
    if (s === 'konto-offen') return 'Konto wird gelesen';
    if (s === 'default') return 'Standard, ohne Zuordnung';
    if (s === 'portal') return 'Portal-Persona';
    return s;
}

/**
 * @param {string} [pathname]
 * @returns {SchoolAppChrome|null}
 */
export function matchSchoolAppFromPath(pathname) {
    const path = String(pathname || '')
        .replace(/\\/g, '/')
        .toLowerCase();
    const apps = SCHOOL_APP_CHROME.slice().sort((a, b) => b.file.length - a.file.length);
    for (let i = 0; i < apps.length; i++) {
        const app = apps[i];
        const file = app.file.toLowerCase();
        if (path === file || path.endsWith('/' + file)) return app;
    }
    return null;
}

/**
 * @param {SchoolAppChrome|null} app
 * @param {{ role?: string, source?: string, email?: string, name?: string, authSettled?: boolean, loggedInWithoutEmail?: boolean, accessDenied?: boolean, demoMode?: boolean }} [input]
 */
export function appRoleChromeModel(app, input) {
    const info = input || {};
    const email = String(info.email || '')
        .trim()
        .toLowerCase();
    const role = String(info.role || '');
    const source = String(info.source || '');
    const title = app ? app.title : 'Schul-App';
    if (!email && !info.demoMode) {
        if (info.authSettled && !info.loggedInWithoutEmail) {
            return {
                appTitle: title,
                roleLabel: 'Nicht angemeldet',
                sourceLabel: '',
                email: '',
                pending: false
            };
        }
        return {
            appTitle: title,
            roleLabel: 'Konto wird gelesen',
            sourceLabel: '',
            email: '',
            pending: true
        };
    }
    if (email && info.accessDenied) {
        return {
            appTitle: title,
            roleLabel: 'Kein Zugriff',
            sourceLabel: '',
            email,
            name: String(info.name || '').trim(),
            pending: false
        };
    }
    if (!role && !(app && app.portalOnly)) {
        return {
            appTitle: title,
            roleLabel: 'Rolle wird bestimmt',
            sourceLabel: '',
            email,
            pending: true
        };
    }
    if (app && app.portalOnly && !role) {
        return {
            appTitle: title,
            roleLabel: 'Portal-Persona',
            sourceLabel: 'dieses Tool hat keine eigene Genehmiger-Rolle',
            email,
            pending: false
        };
    }
    const labels = (app && app.roles) || {};
    const roleLabel = labels[role] || DEFAULT_ROLE_LABELS[role] || (role ? role : 'Lehrkraft');
    const quiet = !source || source === 'default' || source === 'konto-offen';
    const sourceLabel = quiet ? '' : appRoleSourceLabel(source);
    return {
        appTitle: title,
        roleLabel,
        sourceLabel,
        email,
        name: String(info.name || '').trim(),
        pending: false
    };
}

/**
 * @param {ReturnType<typeof appRoleChromeModel>} model
 */
export function appRoleChromeText(model) {
    const m = model || { appTitle: 'Schul-App', roleLabel: 'Konto wird gelesen', sourceLabel: '' };
    const bits = [m.appTitle, m.roleLabel].filter(Boolean);
    if (m.sourceLabel) bits.push(m.sourceLabel);
    return bits.join(' · ');
}

export function appRoleChipLabel(model) {
    const m = model || { roleLabel: 'Konto wird gelesen' };
    return String(m.roleLabel || 'Konto wird gelesen');
}

/** @type {ReturnType<typeof appRoleChromeModel>|null} */
let currentModel = null;

/**
 * @param {string} appId
 * @param {{ role?: string, source?: string, email?: string, name?: string, planerRoleSources?: Record<string, string>, roleSource?: string, accountEmail?: string, accountName?: string, authRolePending?: boolean, authSettled?: boolean, kontoOffen?: boolean, planerAccessDenied?: boolean, demoRoleOverride?: boolean, localDemoOnly?: boolean, loading?: boolean }} state
 */
export function announcePlannerState(appId, state) {
    const s = state || {};
    const demoMode = !!(s.demoRoleOverride || s.localDemoOnly);
    const pendingRole = !!s.authRolePending && !demoMode;
    const role = pendingRole ? '' : String(s.role || '');
    const sources = s.planerRoleSources || {};
    const source = pendingRole ? '' : String(sources[role] || s.roleSource || s.source || '');
    announceAppRole({
        appId,
        role,
        source,
        email: s.accountEmail || s.email || '',
        name: s.accountName || s.name || '',
        authSettled: !!s.authSettled && !pendingRole,
        loggedInWithoutEmail: !!s.kontoOffen,
        accessDenied: !!s.planerAccessDenied && !pendingRole && !s.kontoOffen,
        demoMode
    });
}

/**
 * Liest den aktuellen MSAL-/Mock-Kontostand über alle unterstützten Quellen.
 * @returns {{ email: string, name: string, loggedIn: boolean, bootSettled: boolean }}
 */
export function readAuthAccountSnapshot() {
    let email = '';
    let name = '';
    let loggedIn = false;
    let bootSettled = false;
    try {
        if (typeof window !== 'undefined') {
            const info =
                typeof window.ms365AuthGetAccountInfo === 'function'
                    ? window.ms365AuthGetAccountInfo()
                    : null;
            const upn =
                typeof window.ms365AuthGetUserPrincipalName === 'function'
                    ? window.ms365AuthGetUserPrincipalName()
                    : '';
            const legacy =
                typeof window.ms365GetAccountInfo === 'function'
                    ? window.ms365GetAccountInfo()
                    : null;
            const msalAcc = window.msalAccount || null;
            email = String(
                (info && (info.email || info.username || info.mail || info.upn)) ||
                    (legacy && (legacy.email || legacy.username || legacy.mail || legacy.upn)) ||
                    (msalAcc && (msalAcc.username || msalAcc.mail)) ||
                    upn ||
                    ''
            )
                .trim()
                .toLowerCase();
            name = String(
                (info && info.name) ||
                    (legacy && legacy.name) ||
                    (msalAcc && msalAcc.name) ||
                    ''
            ).trim();
            loggedIn =
                !!email ||
                (typeof window.ms365AuthIsLoggedIn === 'function' && !!window.ms365AuthIsLoggedIn());
            bootSettled =
                typeof window.ms365AuthIsBootSettled === 'function'
                    ? !!window.ms365AuthIsBootSettled()
                    : false;
        }
    } catch {
        /* ignore */
    }
    return { email, name, loggedIn, bootSettled };
}

/**
 * Wartet deterministisch auf MSAL-Initialisierung + SSO-Boot, bevor ein Tool
 * über „Nicht angemeldet“ oder „Kein Zugriff“ urteilt (siehe rollen-vertrag.md).
 * @param {{ timeoutMs?: number }} [opts]
 * @returns {Promise<{ email: string, name: string, loggedIn: boolean, bootSettled: boolean, settled: boolean }>}
 */
export async function waitForAuthAccountRead(opts) {
    const timeoutMs = Math.max(200, Number((opts && opts.timeoutMs) || 3500));
    const immediate = readAuthAccountSnapshot();
    if (immediate.email) return Object.assign({}, immediate, { settled: true });
    if (typeof window === 'undefined') {
        return { email: '', name: '', loggedIn: false, bootSettled: true, settled: true };
    }
    const waitMs = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
    try {
        if (typeof window.ms365AuthEnsureInitialized === 'function') {
            await Promise.race([Promise.resolve(window.ms365AuthEnsureInitialized()), waitMs(timeoutMs)]);
        }
    } catch {
        /* ignore */
    }
    const afterPca = readAuthAccountSnapshot();
    if (afterPca.email) return Object.assign({}, afterPca, { settled: true });
    try {
        if (typeof window.ms365AuthEnsureBootSettled === 'function' && !afterPca.bootSettled) {
            await Promise.race([Promise.resolve(window.ms365AuthEnsureBootSettled()), waitMs(timeoutMs)]);
        }
    } catch {
        /* ignore */
    }
    let current = readAuthAccountSnapshot();
    if (current.loggedIn && !current.email) {
        const started = Date.now();
        while (Date.now() - started < Math.min(1000, timeoutMs)) {
            await waitMs(80);
            current = readAuthAccountSnapshot();
            if (current.email || !current.loggedIn) break;
        }
    }
    return Object.assign({}, current, { settled: true });
}

/**
 * Ablauf- und Betrieb-Apps: Leitung, sonst Lehrkraft oder Klassenvorstand.
 * @param {string} appId
 * @param {{ email?: string, isAdmin?: boolean, isKv?: boolean, isTeacher?: boolean, sources?: string[], name?: string }} access
 */
export function announceStaffAccess(appId, access) {
    const a = access || {};
    let role = String(a.role || '');
    if (!role) {
        if (a.isAdmin) role = 'leitung';
        else if (a.isKv) role = 'kv';
        else role = 'lehrer';
    }
    const sources = Array.isArray(a.sources) ? a.sources : [];
    announceAppRole({
        appId,
        role,
        source: sources[0] || (a.email ? 'stammdaten' : ''),
        email: a.email || '',
        name: a.name || ''
    });
}

/**
 * @param {{ appId: string, role?: string, source?: string, email?: string, name?: string }} detail
 */
export function announceAppRole(detail) {
    const d = detail || {};
    const app =
        SCHOOL_APP_CHROME.find((row) => row.id === d.appId) ||
        matchSchoolAppFromPath(typeof location !== 'undefined' ? location.pathname : '');
    currentModel = appRoleChromeModel(app, d);
    renderAppRoleChip();
    try {
        if (typeof window !== 'undefined') {
            window.__ms365AppRoleLast = Object.assign({ appId: d.appId || (app && app.id) || '' }, currentModel);
        }
    } catch {
        /* ignore */
    }
}

function signedInEmail() {
    try {
        if (typeof window !== 'undefined' && typeof window.ms365AuthGetAccountInfo === 'function') {
            const account = window.ms365AuthGetAccountInfo();
            return String((account && (account.username || account.upn)) || '')
                .trim()
                .toLowerCase();
        }
    } catch {
        /* ignore */
    }
    return '';
}

export function bootAppRoleChrome() {
    if (typeof document === 'undefined') return;
    const app = matchSchoolAppFromPath(window.location && window.location.pathname);
    if (!currentModel || currentModel.pending) {
        const snap = readAuthAccountSnapshot();
        const email = snap.email || signedInEmail();
        if (app && app.portalOnly) {
            currentModel = appRoleChromeModel(app, {
                email,
                authSettled: snap.bootSettled,
                loggedInWithoutEmail: snap.loggedIn && !email
            });
        } else if (!currentModel) {
            currentModel = appRoleChromeModel(app, {
                email,
                authSettled: snap.bootSettled,
                loggedInWithoutEmail: snap.loggedIn && !email
            });
        }
    }
    renderAppRoleChip();
    if (!bootAppRoleChrome._wired) {
        bootAppRoleChrome._wired = true;
        const refreshFromAuth = function () {
            const current = matchSchoolAppFromPath(window.location && window.location.pathname);
            const snap = readAuthAccountSnapshot();
            const email = snap.email || signedInEmail();
            if (current && current.portalOnly) {
                currentModel = appRoleChromeModel(current, {
                    email,
                    authSettled: snap.bootSettled,
                    loggedInWithoutEmail: snap.loggedIn && !email
                });
            } else if (currentModel && currentModel.pending && !email && snap.bootSettled) {
                currentModel = appRoleChromeModel(current, {
                    email: '',
                    authSettled: true,
                    loggedInWithoutEmail: snap.loggedIn
                });
            }
            renderAppRoleChip();
        };
        window.addEventListener('ms365-auth-state-changed', refreshFromAuth);
        window.addEventListener('ms365-auth-boot-settled', refreshFromAuth);
    }
}

function renderAppRoleChip() {
    if (typeof document === 'undefined') return;
    const actions = document.getElementById('adminAppTopActions');
    if (!actions) return;
    const app = matchSchoolAppFromPath(window.location && window.location.pathname);
    if (!app) {
        const old = document.getElementById('ms365AppRoleChip');
        if (old) old.hidden = true;
        return;
    }
    let chip = document.getElementById('ms365AppRoleChip');
    if (!chip) {
        chip = document.createElement('div');
        chip.id = 'ms365AppRoleChip';
        chip.className = 'ms365-app-role-chip';
        actions.insertBefore(chip, actions.firstChild);
    }
    chip.hidden = false;
    const model = currentModel || appRoleChromeModel(app, {});
    const label = appRoleChipLabel(model);
    chip.title = appRoleChromeText(model) + (model.email ? ' · ' + model.email : '');
    chip.setAttribute('aria-label', chip.title);
    chip.textContent = label;
    chip.style.cssText = [
        'display:inline-flex',
        'align-items:center',
        'flex:0 0 auto',
        'height:28px',
        'padding:0 12px',
        'margin-right:10px',
        'border-radius:999px',
        'font-size:12.5px',
        'font-weight:650',
        'line-height:1',
        'letter-spacing:0.01em',
        'white-space:nowrap',
        'background:color-mix(in srgb, var(--brand1, #5b4bdb) 16%, #fff)',
        'color:var(--brand1, #3a2f9a)',
        'border:1px solid color-mix(in srgb, var(--brand1, #5b4bdb) 38%, #fff)'
    ].join(';');
}
