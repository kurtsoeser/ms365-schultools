/**
 * Freigeber für Teams-Antrag: Sicherheitsgruppe (wie Group.Unified) + optionale Extra-Mails.
 */
import { fetchUserMemberGroupIds } from '../schularbeiten-planer/schularbeiten-planer-entra-role.js';
import { userIsDesignatedSchoolIt } from '../../shared/designated-school-it.js';
import { isFreigeber as isFreigeberMail } from './teams-antrag-logic.js';
import { loadTeamsAntragSettings } from './teams-antrag-settings.js';

const GP_STORAGE = 'ms365-gruppenerstellung-policy-v1';

function graphApi() {
    return typeof window !== 'undefined' ? window.ms365GraphUnifiedGroups : null;
}

/**
 * Zwischenstand aus dem Tool „Wer darf Teams anlegen?“.
 * @returns {{ groupId: string, groupDisplayName: string } | null}
 */
export function readGruppenerstellungLocal() {
    try {
        if (typeof localStorage === 'undefined') return null;
        const raw = localStorage.getItem(GP_STORAGE);
        if (!raw) return null;
        const s = JSON.parse(raw);
        const groupId = String(s.resolvedObjectId || s.groupObjectIdRaw || '').trim();
        if (!/^[0-9a-f-]{36}$/i.test(groupId)) return null;
        return {
            groupId,
            groupDisplayName: String(s.groupDisplayName || '').trim()
        };
    } catch {
        return null;
    }
}

/**
 * Liest GroupCreationAllowedGroupId aus der Directory-Einstellung Group.Unified.
 * @returns {Promise<{ groupId: string } | null>}
 */
export async function fetchGroupUnifiedAllowedGroupId() {
    const G = graphApi();
    if (!G || typeof G.getGraphToken !== 'function' || typeof G.graphJson !== 'function') {
        return null;
    }
    const scopes = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Directory.Read.All',
        'https://graph.microsoft.com/Group.Read.All'
    ];
    try {
        const token = await G.getGraphToken(scopes);
        const data = await G.graphJson('GET', '/groupSettings', token, undefined);
        const list = (data && data.value) || [];
        const unified = list.find((x) => x && String(x.displayName) === 'Group.Unified');
        if (!unified || !Array.isArray(unified.values)) return null;
        const row = unified.values.find(
            (v) => v && String(v.name) === 'GroupCreationAllowedGroupId'
        );
        const groupId = String((row && row.value) || '').trim();
        if (!/^[0-9a-f-]{36}$/i.test(groupId)) return null;
        return { groupId };
    } catch {
        return null;
    }
}

/**
 * @param {string} groupId
 * @returns {Promise<{ id: string, displayName: string, mail: string, securityEnabled: boolean, mailEnabled: boolean } | null>}
 */
export async function fetchGroupById(groupId) {
    const id = String(groupId || '').trim();
    if (!/^[0-9a-f-]{36}$/i.test(id)) return null;
    const G = graphApi();
    if (!G) return null;
    try {
        const token = await G.getGraphToken([
            'https://graph.microsoft.com/User.Read',
            'https://graph.microsoft.com/Group.Read.All'
        ]);
        const g = await G.graphJson(
            'GET',
            '/groups/' +
                encodeURIComponent(id) +
                '?$select=id,displayName,mail,securityEnabled,mailEnabled,groupTypes',
            token,
            undefined
        );
        if (!g || !g.id) return null;
        return {
            id: String(g.id),
            displayName: String(g.displayName || ''),
            mail: String(g.mail || ''),
            securityEnabled: !!g.securityEnabled,
            mailEnabled: !!g.mailEnabled
        };
    } catch {
        return null;
    }
}

/**
 * Einfache Gruppensuche für die Auswahl in den Einstellungen.
 * @param {string} query
 */
export async function searchSecurityGroups(query) {
    const q = String(query || '').trim();
    if (q.length < 2) return [];
    const G = graphApi();
    if (!G) return [];
    const token = await G.getGraphToken([
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Group.Read.All'
    ]);
    const safe = q.replace(/'/g, "''");
    const path =
        "/groups?$filter=startswith(displayName,'" +
        safe +
        "')&$select=id,displayName,mail,securityEnabled,mailEnabled,groupTypes&$top=25";
    const data = await G.graphJson('GET', path, token, undefined);
    return ((data && data.value) || []).map((g) => ({
        id: String(g.id || ''),
        displayName: String(g.displayName || ''),
        mail: String(g.mail || ''),
        securityEnabled: !!g.securityEnabled,
        mailEnabled: !!g.mailEnabled,
        isUnified: Array.isArray(g.groupTypes) && g.groupTypes.indexOf('Unified') !== -1
    }));
}

/**
 * Effektive Freigeber-Gruppe: Tool-Einstellung → sonst Group.Unified → sonst lokaler Gruppenerstellungs-Stand.
 * @param {ReturnType<typeof loadTeamsAntragSettings>} [settings]
 */
export async function resolveFreigeberGroup(settings) {
    const s = settings || loadTeamsAntragSettings();
    if (s.freigeberGroupId) {
        return {
            groupId: s.freigeberGroupId,
            groupDisplayName: s.freigeberGroupName || '',
            source: 'settings'
        };
    }
    const fromDir = await fetchGroupUnifiedAllowedGroupId();
    if (fromDir && fromDir.groupId) {
        const meta = await fetchGroupById(fromDir.groupId);
        return {
            groupId: fromDir.groupId,
            groupDisplayName: (meta && meta.displayName) || '',
            source: 'group-unified'
        };
    }
    const local = readGruppenerstellungLocal();
    if (local) {
        return {
            groupId: local.groupId,
            groupDisplayName: local.groupDisplayName,
            source: 'gruppenerstellung-local'
        };
    }
    return { groupId: '', groupDisplayName: '', source: '' };
}

/**
 * @param {{ email?: string, settings?: object }} [opts]
 * @returns {Promise<{ ok: boolean, reasons: string[], groupId: string, groupName: string, source: string, setupNeeded: boolean }>}
 */
export async function resolveFreigeberAccess(opts) {
    const settings = (opts && opts.settings) || loadTeamsAntragSettings();
    const email = String((opts && opts.email) || '')
        .trim()
        .toLowerCase();
    /** @type {string[]} */
    const reasons = [];
    let ok = false;

    const groupInfo = await resolveFreigeberGroup(settings);
    const groupId = groupInfo.groupId;
    const groupName = groupInfo.groupDisplayName;

    if (groupId && email) {
        try {
            const member = await fetchUserMemberGroupIds([groupId]);
            if (member.has(String(groupId).toLowerCase())) {
                ok = true;
                reasons.push('Mitglied der Freigeber-Gruppe');
            }
        } catch {
            /* ignore – Mail-Fallback */
        }
    }

    if (isFreigeberMail(email, settings.freigeberMails || [])) {
        ok = true;
        reasons.push('Zusätzliche Freigeber-Mail');
    }

    try {
        if (await userIsDesignatedSchoolIt()) {
            ok = true;
            reasons.push('Schul-IT (Lizenzkontakt)');
        }
    } catch {
        /* ignore */
    }

    const setupNeeded = !groupId && !(settings.freigeberMails || []).length;

    return {
        ok,
        reasons,
        groupId,
        groupName,
        source: groupInfo.source,
        setupNeeded
    };
}
