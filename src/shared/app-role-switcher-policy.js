/**
 * Einheitliche Policy für App-Rollen-Umschalter und IT-Demo-Vorschau.
 * Referenz: Schularbeiten-/Freistellungs-Planer (docs/rollen-vertrag.md).
 *
 * - Demo-UI nur bei ?demoRole=1 (bzw. explizitem Override)
 * - Produktiv: Umschalter nur bei echten Mehrfachrollen;
 *   optional nur wenn eine privilegierte Rolle dabei ist (z. B. Planer-Admin)
 */

/**
 * @param {boolean} [demoRoleOverride]
 * @returns {boolean}
 */
export function isDemoRoleUiEnabled(demoRoleOverride) {
    return !!demoRoleOverride;
}

/**
 * @param {boolean} [demoRoleOverride]
 * @returns {boolean}
 */
export function readDemoRoleOverrideFromUrl(demoRoleOverride) {
    if (demoRoleOverride) return true;
    try {
        const p = new URLSearchParams(
            typeof window !== 'undefined' ? window.location.search || '' : ''
        );
        return p.get('demoRole') === '1';
    } catch {
        return false;
    }
}

/**
 * @param {string[]|null|undefined} roles
 * @returns {string[]}
 */
export function normalizeAppRoles(roles) {
    const out = [];
    const seen = new Set();
    for (const raw of roles || []) {
        const r = String(raw || '')
            .trim()
            .toLowerCase();
        if (!r || seen.has(r)) continue;
        seen.add(r);
        out.push(r);
    }
    return out;
}

/**
 * @param {string[]|null|undefined} roles
 * @param {{
 *   demoRoleUi?: boolean,
 *   privilegedRoles?: string[]|null
 * }} [opts]
 * @returns {boolean}
 */
export function canUseAppRoleSwitcher(roles, opts) {
    const demoRoleUi = !!(opts && opts.demoRoleUi);
    if (demoRoleUi) return true;
    const list = normalizeAppRoles(roles);
    if (list.length <= 1) return false;
    const priv = opts && opts.privilegedRoles;
    if (!priv || !priv.length) return true;
    const need = new Set(normalizeAppRoles(priv));
    return list.some((r) => need.has(r));
}

/**
 * Toolbar/Setup nur in privilegierter aktiver Rolle.
 * @param {string[]|null|undefined} roles
 * @param {string|null|undefined} activeRole
 * @param {string} [privilegedRole]
 */
export function canShowPrivilegedAppToolbar(roles, activeRole, privilegedRole) {
    const need = String(privilegedRole || 'admin')
        .trim()
        .toLowerCase();
    if (!need) return false;
    const list = normalizeAppRoles(roles);
    return list.includes(need) && String(activeRole || '').trim().toLowerCase() === need;
}

/**
 * Anzeige der erkannten Portal-Persona (kein Umschalter).
 * @param {{
 *   loggedIn?: boolean,
 *   isIt?: boolean,
 *   isLehrer?: boolean,
 *   isSchueler?: boolean,
 *   unknownUser?: boolean
 * }|null|undefined} personas
 * @returns {string}
 */
export function primaryPersonaLabel(personas) {
    if (!personas || !personas.loggedIn) return '';
    if (personas.isIt) return 'Schul-IT';
    if (personas.isLehrer) return 'Lehrkraft';
    // Ohne lokale Gruppen/Stammdaten wirkt der Default wie „Schüler“ – klarer labeln
    if (personas.unknownUser) return 'Noch nicht zugeordnet';
    if (personas.isSchueler) return 'Schüler/in';
    return '';
}

/**
 * Präziser Hinweis bei unbekannter Portal-Persona.
 * Domain allein ≠ Lehrkraft – Checklist aus Diagnostik.
 *
 * @param {{
 *   loggedIn?: boolean,
 *   unknownUser?: boolean,
 *   personaSource?: string,
 *   personaDiagnostics?: {
 *     hasDomain?: boolean,
 *     teacherCount?: number,
 *     groupsConfigured?: boolean,
 *     inTeacherList?: boolean
 *   }
 * }|null|undefined} personas
 * @returns {string}
 */
export function unknownPersonaHint(personas) {
    if (!personas || !personas.unknownUser) return '';
    const d = personas.personaDiagnostics || {};
    const hasDomain = !!d.hasDomain;
    const teacherCount = Number(d.teacherCount) || 0;
    const groupsConfigured = !!d.groupsConfigured;
    const inTeacherList = !!d.inTeacherList;
    const source = String(personas.personaSource || '');

    if (!hasDomain) {
        return 'Schul-Domain fehlt noch in diesem Browser. IT: Lizenz-ToolSiteUrl setzen, IT-Sicherung ablegen; dann neu anmelden.';
    }
    if (teacherCount <= 0) {
        return 'Domain ist da, aber die Lehrerliste ist leer in diesem Browser. Seite neu laden – die App liest „Lehrerinnen“ von der Tool-Site. Sonst IT: IT-Sicherung ablegen oder unter Schulaktivitäten „Klassen neu laden“.';
    }
    if (groupsConfigured && source === 'dashboard-entra-none' && !inTeacherList) {
        return 'Nicht in der Lehrkräfte-Gruppe und E-Mail fehlt in der Lehrerliste. IT: Gruppenmitgliedschaft oder Lehrerlisten-Eintrag prüfen; dann Seite neu laden.';
    }
    if (!inTeacherList) {
        return 'E-Mail steht nicht in der Lehrerliste (Domain allein reicht nicht). IT: Eintrag in der Lehrerliste oder Lehrkräfte-Gruppe; dann Seite neu laden.';
    }
    return 'Lehrkraft noch nicht erkannt. IT: Lehrkräfte-Gruppe + Lehrerliste prüfen; dann Seite neu laden.';
}

/**
 * Kurzhinweis unter der Persona-Zeile im Auth-Menü.
 * @param {{
 *   loggedIn?: boolean,
 *   isIt?: boolean,
 *   isLehrer?: boolean,
 *   unknownUser?: boolean,
 *   personaSource?: string,
 *   personaDiagnostics?: object
 * }|null|undefined} personas
 * @returns {string}
 */
export function primaryPersonaHint(personas) {
    if (!personas || !personas.loggedIn) return '';
    if (personas.isIt) {
        return 'Bereich darunter: Portal / IT-Backend / Schüler-Vorschau wechseln.';
    }
    if (personas.unknownUser) {
        return unknownPersonaHint(personas);
    }
    return 'Fest aus Anmeldung und Gruppen – kein Umschalter.';
}
