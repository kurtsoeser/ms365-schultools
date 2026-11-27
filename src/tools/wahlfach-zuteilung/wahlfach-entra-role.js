/**
 * Wahlfach-Zuteilung: App-Rolle, nicht Portal-Persona.
 * Direktion teilt zu. Klassenvorstand sieht nur die eigene Klasse.
 * Lehrkraft ohne KV-Klasse sieht die Zuteilung nicht.
 * Schüler nutzen die separate Anmelde-Seite.
 * IT-Vorschau: ?demoRole=1 (Direktion), ?demoRole=kv, ?demoRole=lehrer.
 */
import { userIsDesignatedSchoolIt } from '../../shared/designated-school-it.js';
import {
    fetchUserMemberGroupIds,
    userIsEntraGlobalAdministrator
} from '../schularbeiten-planer/schularbeiten-planer-entra-role.js';
import { loadEffectivePermissionsConfig as loadSaPerms } from '../schularbeiten-planer/schularbeiten-planer-permissions.js';
import { accountIsPlannerUserInList } from '../freistellung-planer/freistellung-planer-direktion-users.js';
import { announceAppRole } from '../../shared/app-role-chrome.js';
import { kvKlassenForEmail, readWzAccount } from './wahlfach-zuteilung-state.js';

function accountEmail() {
    return readWzAccount().email;
}

/**
 * @returns {'direktion'|'kv'|'lehrer'|''}
 */
export function readWahlfachDemoRole() {
    try {
        const p = new URLSearchParams(
            typeof window !== 'undefined' ? window.location.search || '' : ''
        );
        const v = String(p.get('demoRole') || '')
            .trim()
            .toLowerCase();
        if (v === '1' || v === 'direktion' || v === 'admin') return 'direktion';
        if (v === 'kv' || v === 'lehrer') return v;
    } catch {
        /* ignore */
    }
    return '';
}

/**
 * Klassen für die KV-Vorschau (?demoRole=kv). Standard: Demo-Klasse 3AK.
 * @returns {string[]}
 */
export function demoKvKlassenFromUrl() {
    try {
        const p = new URLSearchParams(
            typeof window !== 'undefined' ? window.location.search || '' : ''
        );
        const raw = String(p.get('klasse') || p.get('wzKlassen') || '3AK');
        const out = [];
        const seen = new Set();
        for (const part of raw.split(/[;,]/)) {
            const code = part.trim();
            if (!code) continue;
            const key = code.toLowerCase();
            if (seen.has(key)) continue;
            seen.add(key);
            out.push(code);
        }
        return out.length ? out : ['3AK'];
    } catch {
        return ['3AK'];
    }
}

function finishWahlfach(result) {
    const email = result && result.email ? result.email : accountEmail();
    const role = result && result.role;
    announceAppRole({
        appId: 'wahlfach',
        role: role === 'direktion' ? 'admin' : role === 'pending' ? '' : role || '',
        source: result && result.demo ? 'demo' : (result && result.source) || '',
        email,
        demoMode: !!(result && result.demo)
    });
    return result;
}

/**
 * @param {{ stammdaten?: { classes?: object[] }|null, direktion?: { ok?: boolean, reason?: string, source?: string }|null }} [opts]
 * @returns {Promise<{ role: 'direktion'|'kv'|'lehrer'|'pending', ok: boolean, kvKlassen: string[], reason: string, source: string, demo: boolean, email: string }>}
 */
export async function resolveWahlfachAppRole(opts) {
    const demoRole = readWahlfachDemoRole();
    if (demoRole === 'direktion') {
        return finishWahlfach({
            role: 'direktion',
            ok: true,
            kvKlassen: [],
            reason: 'IT-Vorschau Direktion',
            source: 'demo',
            demo: true,
            email: accountEmail()
        });
    }
    if (demoRole === 'lehrer') {
        return finishWahlfach({
            role: 'lehrer',
            ok: false,
            kvKlassen: [],
            reason: 'IT-Vorschau Lehrkraft',
            source: 'demo',
            demo: true,
            email: accountEmail()
        });
    }
    if (demoRole === 'kv') {
        const kvKlassen = demoKvKlassenFromUrl();
        return finishWahlfach({
            role: 'kv',
            ok: false,
            kvKlassen,
            reason: 'IT-Vorschau Klassenvorstand',
            source: 'demo',
            demo: true,
            email: accountEmail()
        });
    }

    const known = opts && opts.direktion;
    /** @type {{ ok: boolean, reason: string, source: string }} */
    let direktion = known && typeof known.ok === 'boolean' ? known : { ok: false, reason: '', source: '' };
    if (!(known && typeof known.ok === 'boolean')) {
        direktion = await detectWahlfachDirektion();
    }
    if (direktion.ok) {
        return finishWahlfach({
            role: 'direktion',
            ok: true,
            kvKlassen: [],
            reason: direktion.reason,
            source: direktion.source,
            demo: false,
            email: accountEmail()
        });
    }

    const email = accountEmail();
    if (!email) {
        return finishWahlfach({
            role: 'pending',
            ok: false,
            kvKlassen: [],
            reason: 'Konto wird gelesen',
            source: 'konto-offen',
            demo: false,
            email: ''
        });
    }

    const kvKlassen = kvKlassenForEmail(opts && opts.stammdaten, email);
    if (kvKlassen.length) {
        return finishWahlfach({
            role: 'kv',
            ok: false,
            kvKlassen,
            reason: 'Klassenvorstand ' + kvKlassen.join(', '),
            source: 'stammdaten-kv',
            demo: false,
            email
        });
    }

    return finishWahlfach({
        role: 'lehrer',
        ok: false,
        kvKlassen: [],
        reason: 'Lehrkraft ohne Klassenvorstand',
        source: '',
        demo: false,
        email
    });
}

/**
 * Direktion prüfen, ohne die App-Rolle im Kopf zu setzen.
 * @returns {Promise<{ ok: boolean, reason: string, source: string }>}
 */
export async function probeWahlfachDirektion() {
    return detectWahlfachDirektion();
}

/**
 * @returns {Promise<{ ok: boolean, reason: string, source: string }>}
 */
async function detectWahlfachDirektion() {
    try {
        if (await userIsEntraGlobalAdministrator()) {
            return { ok: true, reason: 'Globaler Administrator', source: 'global-admin' };
        }
    } catch {
        /* ignore */
    }
    try {
        if (await userIsDesignatedSchoolIt()) {
            return { ok: true, reason: 'Schul-IT', source: 'school-it' };
        }
    } catch {
        /* ignore */
    }

    let sa = { groupAdminId: '', adminUsers: [] };
    try {
        sa = loadSaPerms();
    } catch {
        /* ignore */
    }

    if (sa.groupAdminId) {
        try {
            const member = await fetchUserMemberGroupIds([sa.groupAdminId]);
            if (member.has(String(sa.groupAdminId).toLowerCase())) {
                return { ok: true, reason: 'Schulleitung / Admin-Gruppe', source: 'entra' };
            }
        } catch {
            /* ignore */
        }
    }

    const email = accountEmail();
    if (email && accountIsPlannerUserInList(email, sa.adminUsers || [])) {
        return { ok: true, reason: 'Direktion (Person)', source: 'admin-user' };
    }

    return { ok: false, reason: '', source: '' };
}

/**
 * @param {{ stammdaten?: { classes?: object[] }|null }} [opts]
 * @returns {Promise<{ ok: boolean, reason: string, demo: boolean }>}
 */
export async function resolveWahlfachAdminAccess(opts) {
    const role = await resolveWahlfachAppRole(opts);
    return { ok: role.role === 'direktion', reason: role.reason, demo: role.demo };
}
