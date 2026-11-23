/**
 * Login-Scopes je Seite / Persona (Plan a4 / R3, Trust-Seite c10).
 *
 * Regel: Der Anmelden-Button fordert nur User.Read plus die Scopes des
 * geöffneten Werkzeugs. Admin-Bundles aus graph-unified-groups.js gehören
 * nicht in den ersten Consent – IT-Werkzeuge holen sie bei Bedarf nach
 * (inkrementeller Consent über getGraphToken, Client: it-app-auth.js).
 */

export const USER_READ = 'https://graph.microsoft.com/User.Read';

/**
 * c5: Schreib- und andere hochprivilegierte Scopes laufen über die zweite
 * Entra-App (itClientId, Anzeigename ms365-schultools-it). Die Portal-App
 * behält ihre Schreibscopes vorerst; dieser Katalog entscheidet nur, welcher
 * Client sie anfordert. Zuweisung bleibt aus.
 */
export const SECOND_APP_REGISTRATION_LATER = 'c5';

/**
 * Kurzname ohne Ressource. RoleManagement.Read.Directory fehlt absichtlich
 * (Schularbeiten-Prüfung „Globaler Administrator“ bleibt auf der Portal-App).
 * Group.Read.All fehlt absichtlich (Termine-Hub bleibt auf der Portal-App).
 */
export const IT_APP_SCOPE_NAMES = Object.freeze([
    'user.readwrite.all',
    'group.readwrite.all',
    'directory.readwrite.all',
    'directory.read.all',
    'policy.readwrite.authorization',
    'rolemanagement.readwrite.directory',
    'team.create',
    'teamsettings.readwrite.all',
    'channel.create',
    'channelmember.readwrite.all',
    'channelsettings.readwrite.all',
    'sharepointtenantsettings.read.all',
    'sharepointtenantsettings.readwrite.all',
    'sites.create.all',
    'sites.manage.all',
    'eduroster.readwrite',
    'kursteams.create',
    'allsites.fullcontrol'
]);

const IT_APP_SCOPE_SET = new Set(IT_APP_SCOPE_NAMES);

const USER_READ_ALL = 'https://graph.microsoft.com/User.Read.All';
const USER_WRITE_ALL = 'https://graph.microsoft.com/User.ReadWrite.All';
const GROUP_READ_ALL = 'https://graph.microsoft.com/Group.Read.All';
const GROUP_MEMBER_READ_ALL = 'https://graph.microsoft.com/GroupMember.Read.All';
const GROUP_WRITE_ALL = 'https://graph.microsoft.com/Group.ReadWrite.All';
const TEAM_CREATE = 'https://graph.microsoft.com/Team.Create';
const TEAM_SETTINGS_WRITE = 'https://graph.microsoft.com/TeamSettings.ReadWrite.All';
const ORG_READ_ALL = 'https://graph.microsoft.com/Organization.Read.All';
const ROLE_READ_DIRECTORY = 'https://graph.microsoft.com/RoleManagement.Read.Directory';

/** Verzeichnis lesen — ohne Write.All. */
export const GRAPH_SCOPES_DIRECTORY_READ = Object.freeze([
    USER_READ,
    USER_READ_ALL,
    GROUP_READ_ALL,
    GROUP_MEMBER_READ_ALL
]);

/** Gruppen anlegen / Mitglieder ändern (Schul-IT, konkrete Schreibaktion). */
export const GRAPH_SCOPES_GROUP_WRITE = Object.freeze([
    USER_READ,
    USER_READ_ALL,
    GROUP_READ_ALL,
    GROUP_MEMBER_READ_ALL,
    GROUP_WRITE_ALL
]);

/** Benutzer anlegen oder Profil schreiben (Personenanlage). */
export const GRAPH_SCOPES_USER_WRITE = Object.freeze([USER_READ, USER_READ_ALL, USER_WRITE_ALL]);

/** Team anlegen bzw. Team-Einstellungen (nicht beim Portal-Login). */
export const GRAPH_SCOPES_TEAM_WRITE = Object.freeze([
    USER_READ,
    USER_READ_ALL,
    GROUP_READ_ALL,
    GROUP_WRITE_ALL,
    TEAM_CREATE,
    TEAM_SETTINGS_WRITE
]);

/** @typedef {{ id: string, match: RegExp, labelDe: string, persona: string, scopes: string[], warumDe: string }} LoginScopeRule */

/** Graph-Scope → kurze deutsche Erklärung (Consent / Trust-Seite). */
export const SCOPE_LABELS_DE = Object.freeze({
    [USER_READ]: 'Profil lesen (Name, E-Mail)',
    'https://graph.microsoft.com/Sites.Read.All': 'SharePoint-Seiten und Listen lesen',
    'https://graph.microsoft.com/Sites.ReadWrite.All': 'SharePoint-Seiten und Listen lesen und schreiben',
    'https://graph.microsoft.com/Calendars.ReadWrite': 'Kalender lesen und schreiben',
    'https://graph.microsoft.com/Calendars.ReadWrite.Shared': 'Freigegebene Kalender lesen und schreiben',
    'https://graph.microsoft.com/Group.Read.All': 'Gruppenmitgliedschaften lesen',
    'https://graph.microsoft.com/User.Read.All': 'Benutzerverzeichnis lesen (IT)',
    'https://graph.microsoft.com/User.ReadWrite.All': 'Benutzer anlegen und bearbeiten (IT)',
    'https://graph.microsoft.com/Group.ReadWrite.All': 'Gruppen und Teams verwalten (IT)',
    'https://graph.microsoft.com/GroupMember.Read.All': 'Gruppenmitglieder lesen (IT)',
    'https://graph.microsoft.com/Team.Create': 'Teams anlegen (IT)',
    'https://graph.microsoft.com/TeamSettings.ReadWrite.All': 'Team-Einstellungen ändern (IT)',
    'https://graph.microsoft.com/Organization.Read.All': 'Organisationsdaten lesen (IT)',
    'https://graph.microsoft.com/RoleManagement.Read.Directory': 'Verzeichnisrollen lesen (IT)'
});

/** Portal-Werkzeuge: schmale Scopes beim ersten Login. */
/** @type {LoginScopeRule[]} */
export const LOGIN_SCOPE_RULES = [
    {
        id: 'lehrer-freistellung',
        match: /lehrer-freistellung-planer\.html/i,
        labelDe: 'Lehrer-Freistellung',
        persona: 'Lehrkraft',
        scopes: [USER_READ, 'https://graph.microsoft.com/Sites.ReadWrite.All'],
        warumDe: 'Anträge und Genehmigungen in der SharePoint-Liste der Schule.'
    },
    {
        id: 'freistellung',
        match: /freistellung-planer\.html/i,
        labelDe: 'Schüler-Freistellung',
        persona: 'Lehrkraft / Schüler/in',
        scopes: [USER_READ, 'https://graph.microsoft.com/Sites.ReadWrite.All'],
        warumDe: 'Anträge in der SharePoint-Liste der Schule.'
    },
    {
        id: 'schulaktivitaeten',
        match: /schulaktivitaeten-planer\.html/i,
        labelDe: 'Schulaktivitäten',
        persona: 'Lehrkraft',
        scopes: [USER_READ, 'https://graph.microsoft.com/Sites.ReadWrite.All'],
        warumDe: 'Aktivitäten und Genehmigungen in SharePoint.'
    },
    {
        id: 'srdp',
        match: /srdp-anmeldung\.html/i,
        labelDe: 'sRDP-Anmeldung',
        persona: 'Lehrkraft / Schüler/in',
        scopes: [USER_READ, 'https://graph.microsoft.com/Sites.ReadWrite.All'],
        warumDe: 'Anmeldungen und Abschlussarbeiten in SharePoint.'
    },
    {
        id: 'schularbeiten',
        match: /schularbeiten-planer\.html/i,
        labelDe: 'Schularbeiten',
        persona: 'Lehrkraft',
        scopes: [USER_READ, 'https://graph.microsoft.com/Sites.ReadWrite.All'],
        warumDe: 'Schularbeiten-Termine in SharePoint.'
    },
    {
        id: 'termine-hub',
        match: /termine-hub\.html/i,
        labelDe: 'Termine-Hub',
        persona: 'Lehrkraft',
        scopes: [
            USER_READ,
            'https://graph.microsoft.com/Calendars.ReadWrite',
            'https://graph.microsoft.com/Calendars.ReadWrite.Shared',
            'https://graph.microsoft.com/Group.Read.All'
        ],
        warumDe: 'Schul- und Lehrerkalender in Outlook lesen und schreiben.'
    },
    {
        id: 'dashboard',
        match: /(?:^|\/)(?:index|dashboard[^/]*)\.html$/i,
        labelDe: 'Dashboard / Portal',
        persona: 'Lehrkraft / Schüler/in',
        scopes: [USER_READ],
        warumDe: 'Anmeldung und Anzeige des eigenen Kontos – Werkzeuge fordern später eigene Rechte.'
    },
    {
        id: 'tenant-it',
        match: /(?:^|\/)tenant\.html$/i,
        labelDe: 'Stammdaten / IT',
        persona: 'Schul-IT',
        scopes: [USER_READ],
        warumDe:
            'Erster Login nur Profil. Verzeichnis-, Gruppen- und Team-Rechte erscheinen erst, wenn ein IT-Werkzeug sie braucht.'
    }
];

/** Persona-Übersicht für die Trust-Seite. */
export const PERSONA_SCOPE_SUMMARY = Object.freeze([
    {
        persona: 'Lehrkraft / Schüler/in (Portal)',
        scopes: [USER_READ, 'https://graph.microsoft.com/Sites.ReadWrite.All'],
        textDe:
            'Beim Öffnen eines Planers: Profil plus SharePoint der Schule. Dashboard allein: nur Profil.'
    },
    {
        persona: 'Lehrkraft (Termine-Hub)',
        scopes: [
            USER_READ,
            'https://graph.microsoft.com/Calendars.ReadWrite',
            'https://graph.microsoft.com/Calendars.ReadWrite.Shared',
            'https://graph.microsoft.com/Group.Read.All'
        ],
        textDe: 'Kalender und freigegebene Kalender; Gruppen nur lesen für ARGE-/Lehrerkalender.'
    },
    {
        persona: 'Schul-IT',
        scopes: [USER_READ],
        textDe:
            'Erster Login nur Profil auf der Portal-App. Schreibzugriffe laufen über die IT-App ms365-schultools-it; die Schule stimmt einmalig zu. Die Portal-App hat die Schreibscopes noch.'
    },
    {
        persona: 'Betreiber (Lizenz-API)',
        scopes: [],
        textDe:
            'Kein Graph-Zugriff auf Schuldaten. Die Lizenz-API sieht Tenant-ID und Freischaltung, keine Klassenlisten.'
    }
]);

/**
 * @param {string} pathname
 * @returns {LoginScopeRule | null}
 */
export function matchLoginScopeRule(pathname) {
    const path = String(pathname || '');
    for (let i = 0; i < LOGIN_SCOPE_RULES.length; i++) {
        if (LOGIN_SCOPE_RULES[i].match.test(path)) return LOGIN_SCOPE_RULES[i];
    }
    return null;
}

/**
 * @param {string} [pathname]
 * @returns {string[]}
 */
export function resolveLoginScopesForPath(pathname) {
    const rule = matchLoginScopeRule(pathname);
    if (rule && rule.scopes.length) return rule.scopes.slice();
    return [USER_READ];
}

/**
 * @param {string} scope
 * @returns {string}
 */
export function labelScopeDe(scope) {
    const key = String(scope || '').trim();
    if (SCOPE_LABELS_DE[key]) return SCOPE_LABELS_DE[key];
    const short = key.replace(/^https:\/\/graph\.microsoft\.com\//i, '');
    return short || key;
}

/**
 * Schreibt die Seiten-Scopes nach window.MS365_AUTH_LOGIN_SCOPES, sofern die
 * Seite sie nicht schon selbst gesetzt hat.
 * @param {string} [pathname]
 * @returns {string[]}
 */
export function applyPreferredLoginScopes(pathname) {
    const path =
        pathname ||
        (typeof window !== 'undefined' && window.location ? window.location.pathname : '');
    try {
        if (typeof window !== 'undefined') {
            const existing = window.MS365_AUTH_LOGIN_SCOPES;
            if (Array.isArray(existing) && existing.length) {
                return existing.slice();
            }
        }
    } catch {
        /* ignore */
    }
    const scopes = resolveLoginScopesForPath(path);
    try {
        if (typeof window !== 'undefined') {
            window.MS365_AUTH_LOGIN_SCOPES = scopes.slice();
        }
    } catch {
        /* ignore */
    }
    return scopes;
}

/**
 * @param {string} scope
 * @returns {string}
 */
export function shortGraphScope(scope) {
    return String(scope || '')
        .replace(/^https:\/\/graph\.microsoft\.com\//i, '')
        .trim()
        .toLowerCase();
}

/**
 * Letztes Pfadsegment, kleingeschrieben. Gilt für Graph, SharePoint-Host
 * (`https://{host}/AllSites.Write`) und eigene APIs (`api://…/Kursteams.Create`).
 * @param {string} scope
 * @returns {string}
 */
export function scopeShortName(scope) {
    const raw = String(scope || '').trim();
    if (!raw) return '';
    const noQuery = raw.split('?')[0].split('#')[0];
    const slash = noQuery.lastIndexOf('/');
    const name = slash >= 0 ? noQuery.slice(slash + 1) : noQuery;
    return name.trim().toLowerCase();
}

/**
 * IT-Aufrufe, die noch `Sites.FullControl.All` anfordern, auf die erteilte
 * SharePoint-Berechtigung `AllSites.FullControl` umstellen.
 * `AllSites.Write` und `AllSites.Read` (Portal, Lehrkraft) bleiben.
 * @param {string} scope
 * @returns {string}
 */
export function rewriteItSharePointScope(scope) {
    const s = String(scope || '');
    if (/\/Sites\.FullControl\.All$/i.test(s.split('?')[0].split('#')[0])) {
        return s.replace(/\/Sites\.FullControl\.All/i, '/AllSites.FullControl');
    }
    return s;
}

/**
 * @param {string[] | null | undefined} scopes
 * @returns {string[]}
 */
export function prepareItAppScopes(scopes) {
    const list = Array.isArray(scopes) ? scopes : [];
    return list.map(function (scope) {
        return rewriteItSharePointScope(scope);
    });
}

/**
 * @param {string} scope
 * @returns {boolean}
 */
export function scopeRequiresItApp(scope) {
    const name = scopeShortName(rewriteItSharePointScope(scope));
    return IT_APP_SCOPE_SET.has(name);
}

/**
 * Ein IT-Scope in der Liste reicht: der ganze Token-Aufruf geht an die IT-App.
 * Reines `RoleManagement.Read.Directory` bleibt Portal.
 * @param {string[] | null | undefined} scopes
 * @returns {boolean}
 */
export function scopesRequireItApp(scopes) {
    const list = Array.isArray(scopes) ? scopes : [];
    for (let i = 0; i < list.length; i++) {
        if (scopeRequiresItApp(list[i])) return true;
    }
    return false;
}

/**
 * @param {string[] | null | undefined} scopes
 * @returns {'it' | 'portal'}
 */
export function clientKindForScopes(scopes) {
    return scopesRequireItApp(scopes) ? 'it' : 'portal';
}

/**
 * @param {string} accessToken
 * @returns {string[]}
 */
export function readAccessTokenScopeNames(accessToken) {
    const part = String(accessToken || '').split('.')[1];
    if (!part) return [];
    try {
        const b64 = part.replace(/-/g, '+').replace(/_/g, '/');
        const pad = b64.length % 4 === 0 ? b64 : b64 + '='.repeat(4 - (b64.length % 4));
        if (typeof atob !== 'function') return [];
        const binary = atob(pad);
        const json = decodeURIComponent(
            Array.prototype.map
                .call(binary, function (c) {
                    return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
                })
                .join('')
        );
        const payload = JSON.parse(json);
        const raw = payload && (payload.scp || payload.roles);
        const text = Array.isArray(raw) ? raw.join(' ') : String(raw || '');
        return text
            .split(/[\s,]+/)
            .map(function (s) {
                return s.trim().toLowerCase();
            })
            .filter(Boolean);
    } catch {
        return [];
    }
}

/**
 * Read-Scopes gelten als erfüllt, wenn der Token das Schreib-Äquivalent schon trägt.
 * Write.All gilt nicht als durch ein Read-Scope erfüllt.
 * @param {Set<string>} granted
 * @param {string} name
 */
function scopeGranted(granted, name) {
    const n = String(name || '').toLowerCase();
    if (!n) return true;
    if (granted.has(n)) return true;
    if (n === 'user.read' || n === 'user.read.all') return granted.has('user.readwrite.all');
    if (n === 'group.read.all' || n === 'groupmember.read.all') return granted.has('group.readwrite.all');
    return false;
}

/**
 * @param {string} accessToken
 * @param {string[]} scopes
 * @returns {boolean}
 */
export function accessTokenIncludesScopes(accessToken, scopes) {
    const needed = (Array.isArray(scopes) ? scopes : []).map(shortGraphScope).filter(Boolean);
    if (!needed.length) return true;
    const granted = new Set(readAccessTokenScopeNames(accessToken));
    if (!granted.size) return false;
    for (let i = 0; i < needed.length; i++) {
        if (!scopeGranted(granted, needed[i])) return false;
    }
    return true;
}

/**
 * @param {string} pathOrUrl
 * @returns {string}
 */
export function graphPathOnly(pathOrUrl) {
    let p = String(pathOrUrl || '').trim();
    p = p.replace(/^https:\/\/graph\.microsoft\.com\/(v1\.0|beta)/i, '');
    if (p && p.charAt(0) !== '/') p = '/' + p;
    return p.split('?')[0];
}

/**
 * Scopes für eine konkrete Graph-Operation.
 * `null` = Token des Aufrufers behalten (Portal: /me, SharePoint, unbekannte Pfade).
 * @param {string} method
 * @param {string} pathOrUrl
 * @returns {string[] | null}
 */
export function scopesForGraphCall(method, pathOrUrl) {
    const m = String(method || 'GET').toUpperCase();
    const path = graphPathOnly(pathOrUrl);
    if (!path || /^\/me(\/|$)/i.test(path)) return null;
    const write = m === 'POST' || m === 'PATCH' || m === 'PUT' || m === 'DELETE';

    if (/\/subscribedSkus(\/|$)/i.test(path) || /^\/organization(\/|$)/i.test(path)) {
        return [USER_READ, ORG_READ_ALL];
    }
    if (/\/directoryRoles(\/|$)/i.test(path) || /\/roleManagement\//i.test(path)) {
        return [USER_READ, ROLE_READ_DIRECTORY];
    }

    const teamResource = /\/teams(\/|$)/i.test(path) || /\/groups\/[^/]+\/team(\/|$)/i.test(path);
    if (teamResource) {
        if (!write) return [USER_READ, GROUP_READ_ALL];
        const scopes = [USER_READ, GROUP_READ_ALL, GROUP_WRITE_ALL];
        const provision =
            (m === 'POST' && /^\/teams\/?$/i.test(path)) ||
            ((m === 'PUT' || m === 'POST') && /\/groups\/[^/]+\/team\/?$/i.test(path));
        const settings = /\/(photo|archive|unarchive)(\/|$)/i.test(path);
        if (provision) scopes.push(TEAM_CREATE);
        if (settings) scopes.push(TEAM_SETTINGS_WRITE);
        if (!provision && !settings) scopes.push(TEAM_SETTINGS_WRITE);
        return scopes;
    }

    if (/\/groups(\/|$)/i.test(path)) {
        if (!write) {
            if (/\/(members|owners)(\/|$)/i.test(path)) {
                return [USER_READ, GROUP_READ_ALL, GROUP_MEMBER_READ_ALL];
            }
            return [USER_READ, GROUP_READ_ALL];
        }
        return GRAPH_SCOPES_GROUP_WRITE.slice();
    }

    if (/\/users(\/|$)/i.test(path)) {
        if (/\/assignLicense$/i.test(path)) {
            return write ? GRAPH_SCOPES_USER_WRITE.slice() : [USER_READ, USER_READ_ALL];
        }
        if (write) return GRAPH_SCOPES_USER_WRITE.slice();
        return [USER_READ, USER_READ_ALL];
    }

    return null;
}

/**
 * @param {string} kind portal | directory-read | group-write | user-write | team-write
 * @returns {string[]}
 */
export function scopesForGraphKind(kind) {
    const k = String(kind || '').toLowerCase();
    if (k === 'group-write' || k === 'groups-write') return GRAPH_SCOPES_GROUP_WRITE.slice();
    if (k === 'user-write' || k === 'users-write') return GRAPH_SCOPES_USER_WRITE.slice();
    if (k === 'team-write' || k === 'teams-write') return GRAPH_SCOPES_TEAM_WRITE.slice();
    if (k === 'directory-read' || k === 'read') return GRAPH_SCOPES_DIRECTORY_READ.slice();
    return [USER_READ];
}
