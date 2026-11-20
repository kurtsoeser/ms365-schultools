/**
 * Kalender und Termine: Zielkalender (Schul-/Lehrer-/Schülerkalender-Postfächer + Mein Kalender).
 * UPNs/Quellen legt nur IT fest; Endnutzer sehen rollenabhängig die freigeschalteten Kalender.
 */

import { STORAGE_KEY as KALENDER_POSTFAECHER_KEY } from '../kalender-postfaecher/kalender-postfaecher-state.js';

export const SETTINGS_KEY = 'ms365-termine-hub-settings-v1';

export const SOURCE_IDS = {
    ME: 'me',
    SCHUL: 'schul',
    LEHRER: 'lehrer',
    SCHUELER: 'schueler'
};

export const SOURCE_LABEL = {
    me: 'Mein Kalender',
    schul: 'Schulkalender',
    lehrer: 'Lehrerkalender',
    schueler: 'Schülerkalender'
};

/** Welche Persona darf welchen Layer sehen (IT sieht alles). `self` = angemeldet ohne Portal-Rolle. */
export const SOURCE_AUDIENCE = {
    me: ['it', 'lehrer', 'schueler', 'self'],
    schul: ['it', 'lehrer', 'schueler'],
    lehrer: ['it', 'lehrer'],
    schueler: ['it', 'schueler']
};

/**
 * Welche Persona darf in welchen Layer schreiben.
 * Schüler: nur eigener Kalender; Schulkalender nur Lehrkraft/IT.
 * `self`: zumindest eigener Outlook-Kalender (auch ohne Lehrerliste/IT-Rolle).
 */
export const SOURCE_WRITE_AUDIENCE = {
    me: ['it', 'lehrer', 'schueler', 'self'],
    schul: ['it', 'lehrer'],
    lehrer: ['it', 'lehrer'],
    schueler: ['it', 'schueler']
};

/**
 * @param {{
 *   isIt?: boolean,
 *   isLehrer?: boolean,
 *   isSchueler?: boolean,
 *   loggedIn?: boolean,
 *   unknownUser?: boolean
 * }|null|undefined} personas
 * @returns {'it'|'lehrer'|'schueler'|'self'|''}
 */
export function audienceKeyFromPersonas(personas) {
    if (!personas || personas.loggedIn === false) return '';
    if (personas.isIt) return 'it';
    if (personas.isLehrer) return 'lehrer';
    // „Noch nicht zugeordnet“: Dashboard markiert oft unknownUser + Default-Schüler –
    // für Termine trotzdem nur Eigenkalender, keine Shared-Layer.
    if (personas.unknownUser) return 'self';
    if (personas.isSchueler) return 'schueler';
    // Angemeldet, Portal-Rolle noch nicht zugeordnet → eigener Kalender bleibt nutzbar
    if (personas.loggedIn) return 'self';
    return '';
}

/**
 * @param {string} sourceId
 * @param {'it'|'lehrer'|'schueler'|'self'|string} audience
 */
export function sourceAllowedForAudience(sourceId, audience) {
    const id = String(sourceId || '').trim();
    const aud = String(audience || '').trim().toLowerCase();
    const allowed = SOURCE_AUDIENCE[id];
    if (!allowed) return false;
    if (!aud) return false;
    return allowed.includes(aud);
}

/**
 * @param {string} sourceId
 * @param {'it'|'lehrer'|'schueler'|'self'|string} audience
 */
export function sourceWritableForAudience(sourceId, audience) {
    const id = String(sourceId || '').trim();
    const aud = String(audience || '').trim().toLowerCase();
    // M365-Gruppenkalender: Schreibrecht steuert Graph/Mitgliedschaft, nicht die Schul-Rolle.
    if (isGroupSourceId(id)) return !!aud;
    // Eigener Kalender: für jede bekannte Audience inkl. self immer schreibbar,
    // auch wenn IT den Layer „Mein Kalender“ global ausgeblendet hat.
    if (id === SOURCE_IDS.ME && aud) return true;
    const allowed = SOURCE_WRITE_AUDIENCE[id];
    if (!allowed) return false;
    if (!aud) return false;
    return allowed.includes(aud);
}

/** Prefix für M365-Gruppenkalender-Layer (`group:<guid>`). */
export const GROUP_SOURCE_PREFIX = 'group:';

/**
 * @param {string} groupId
 * @returns {string}
 */
export function groupSourceId(groupId) {
    const id = String(groupId || '').trim();
    return id ? GROUP_SOURCE_PREFIX + id : '';
}

/**
 * @param {string} sourceId
 * @returns {string}
 */
export function groupIdFromSourceId(sourceId) {
    const s = String(sourceId || '').trim();
    if (!s.startsWith(GROUP_SOURCE_PREFIX)) return '';
    return s.slice(GROUP_SOURCE_PREFIX.length).trim();
}

/**
 * @param {string} sourceId
 */
export function isGroupSourceId(sourceId) {
    return !!groupIdFromSourceId(sourceId);
}

/**
 * @returns {{
 *   schulkalenderUpn: string,
 *   lehrerkalenderUpn: string,
 *   schuelerkalenderUpn: string,
 *   enabledSources: string[],
 *   visibleSources: string[],
 *   defaultWriteTarget: string,
 *   groupCalendars: Array<{ id: string, displayName: string, mail: string, visible: boolean }>
 * }}
 */
export function emptySettings() {
    return {
        schulkalenderUpn: '',
        lehrerkalenderUpn: '',
        schuelerkalenderUpn: '',
        enabledSources: [
            SOURCE_IDS.ME,
            SOURCE_IDS.SCHUL,
            SOURCE_IDS.LEHRER,
            SOURCE_IDS.SCHUELER
        ],
        visibleSources: [
            SOURCE_IDS.ME,
            SOURCE_IDS.SCHUL,
            SOURCE_IDS.LEHRER,
            SOURCE_IDS.SCHUELER
        ],
        defaultWriteTarget: SOURCE_IDS.LEHRER,
        groupCalendars: []
    };
}

/**
 * @param {unknown} raw
 */
export function normalizeSettings(raw) {
    const base = emptySettings();
    const o = raw && typeof raw === 'object' ? raw : {};
    const enabled = Array.isArray(o.enabledSources)
        ? o.enabledSources.map((x) => String(x || '').trim()).filter(Boolean)
        : base.enabledSources.slice();
    const enabledFinal = enabled.length ? enabled : base.enabledSources.slice();
    const enabledSet = new Set(enabledFinal);
    const visibleRaw = Array.isArray(o.visibleSources)
        ? o.visibleSources.map((x) => String(x || '').trim()).filter(Boolean)
        : enabledFinal.slice();
    const visible = visibleRaw.filter((id) => enabledSet.has(id));
    const write = String(o.defaultWriteTarget || base.defaultWriteTarget)
        .trim()
        .toLowerCase();
    const groupCalendars = Array.isArray(o.groupCalendars)
        ? o.groupCalendars
              .map((g) => {
                  if (!g || typeof g !== 'object') return null;
                  const id = String(g.id || '').trim();
                  if (!id) return null;
                  return {
                      id,
                      displayName: String(g.displayName || g.mail || id).trim() || id,
                      mail: String(g.mail || '')
                          .trim()
                          .toLowerCase(),
                      visible: g.visible !== false
                  };
              })
              .filter(Boolean)
        : [];
    const groupIds = new Set(groupCalendars.map((g) => g.id));
    const dedupedGroups = [];
    groupIds.forEach((gid) => {
        const hit = groupCalendars.find((g) => g.id === gid);
        if (hit) dedupedGroups.push(hit);
    });
    return {
        schulkalenderUpn: String(o.schulkalenderUpn || '')
            .trim()
            .toLowerCase(),
        lehrerkalenderUpn: String(o.lehrerkalenderUpn || '')
            .trim()
            .toLowerCase(),
        schuelerkalenderUpn: String(o.schuelerkalenderUpn || '')
            .trim()
            .toLowerCase(),
        enabledSources: enabledFinal,
        visibleSources: visible.length ? visible : enabledFinal.slice(),
        defaultWriteTarget:
            Object.values(SOURCE_IDS).includes(write) || isGroupSourceId(write)
                ? write
                : SOURCE_IDS.LEHRER,
        groupCalendars: dedupedGroups
    };
}

/**
 * Sichtbare Layer umschalten (mindestens einer bleibt an).
 * @param {string} sourceId
 * @param {boolean} [force]
 * @param {{ audience?: string }} [opts]
 * @returns {ReturnType<typeof normalizeSettings>}
 */
export function toggleVisibleSource(sourceId, force, opts) {
    const s = loadSettings();
    const id = String(sourceId || '').trim();
    const audience = opts && opts.audience ? opts.audience : null;
    const gid = groupIdFromSourceId(id);
    if (gid) {
        const groups = (s.groupCalendars || []).map((g) => ({ ...g }));
        const hit = groups.find((g) => g.id === gid);
        if (!hit) return s;
        const nextOn = force == null ? !hit.visible : !!force;
        const otherVisible =
            s.visibleSources.filter(
                (x) => !audience || sourceAllowedForAudience(x, audience)
            ).length +
            groups.filter((g) => g.id !== gid && g.visible).length;
        if (!nextOn && otherVisible <= 0) return s;
        hit.visible = nextOn;
        return saveSettings({ groupCalendars: groups });
    }
    if (!s.enabledSources.includes(id)) return s;
    if (audience && !sourceAllowedForAudience(id, audience)) return s;
    const allowedVisible = s.visibleSources.filter(
        (x) => !audience || sourceAllowedForAudience(x, audience)
    );
    const groupVisibleCount = (s.groupCalendars || []).filter((g) => g.visible).length;
    const set = new Set(allowedVisible);
    const nextOn = force == null ? !set.has(id) : !!force;
    if (nextOn) set.add(id);
    else {
        if (set.size + groupVisibleCount <= 1 && set.has(id)) return s;
        set.delete(id);
    }
    const others = s.visibleSources.filter(
        (x) => audience && !sourceAllowedForAudience(x, audience)
    );
    return saveSettings({ visibleSources: [...new Set([...others, ...set])] });
}

/**
 * Gruppenkalender in die persönliche Layer-Auswahl übernehmen / entfernen.
 * @param {Array<{ id: string, displayName?: string, mail?: string }>} selected
 */
export function setSelectedGroupCalendars(selected) {
    const cur = loadSettings();
    const prev = new Map((cur.groupCalendars || []).map((g) => [g.id, g]));
    const next = (Array.isArray(selected) ? selected : [])
        .map((g) => {
            const id = String(g && g.id ? g.id : '').trim();
            if (!id) return null;
            const old = prev.get(id);
            return {
                id,
                displayName: String((g && g.displayName) || (old && old.displayName) || id).trim(),
                mail: String((g && g.mail) || (old && old.mail) || '')
                    .trim()
                    .toLowerCase(),
                visible: old ? old.visible !== false : true
            };
        })
        .filter(Boolean);
    return saveSettings({ groupCalendars: next });
}

export function loadSettings() {
    try {
        const raw = localStorage.getItem(SETTINGS_KEY);
        return normalizeSettings(raw ? JSON.parse(raw) : null);
    } catch {
        return emptySettings();
    }
}

/**
 * @param {Partial<ReturnType<typeof emptySettings>>} partial
 */
export function saveSettings(partial) {
    const next = normalizeSettings({ ...loadSettings(), ...(partial || {}) });
    try {
        localStorage.setItem(SETTINGS_KEY, JSON.stringify(next));
    } catch {
        /* ignore */
    }
    return next;
}

/**
 * Vorschläge aus dem Tool „Kalender-Postfächer“.
 * @returns {{ schulkalenderUpn: string, lehrerkalenderUpn: string, schuelerkalenderUpn: string }}
 */
export function suggestUpnsFromKalenderPostfaecher() {
    const out = {
        schulkalenderUpn: '',
        lehrerkalenderUpn: '',
        schuelerkalenderUpn: ''
    };
    try {
        const raw = localStorage.getItem(KALENDER_POSTFAECHER_KEY);
        if (!raw) return out;
        const parsed = JSON.parse(raw);
        const items = Array.isArray(parsed && parsed.items) ? parsed.items : [];
        for (const it of items) {
            const preset = String(it && it.presetId ? it.presetId : '')
                .trim()
                .toLowerCase();
            const smtp = String((it && (it.primarySmtp || it.upn || it.mail)) || '')
                .trim()
                .toLowerCase();
            if (!smtp) continue;
            if (preset === 'schulkalender' && !out.schulkalenderUpn) out.schulkalenderUpn = smtp;
            if (preset === 'lehrerkalender' && !out.lehrerkalenderUpn) out.lehrerkalenderUpn = smtp;
            if (
                (preset === 'schuelerkalender' || preset === 'schueler') &&
                !out.schuelerkalenderUpn
            ) {
                out.schuelerkalenderUpn = smtp;
            }
        }
    } catch {
        /* ignore */
    }
    return out;
}

/**
 * Settings mit leeren UPNs aus Kalender-Postfächer auffüllen (nicht überschreiben).
 */
export function ensureSettingsWithSuggestions() {
    const cur = loadSettings();
    const sug = suggestUpnsFromKalenderPostfaecher();
    const patch = {};
    if (!cur.schulkalenderUpn && sug.schulkalenderUpn) patch.schulkalenderUpn = sug.schulkalenderUpn;
    if (!cur.lehrerkalenderUpn && sug.lehrerkalenderUpn) {
        patch.lehrerkalenderUpn = sug.lehrerkalenderUpn;
    }
    if (!cur.schuelerkalenderUpn && sug.schuelerkalenderUpn) {
        patch.schuelerkalenderUpn = sug.schuelerkalenderUpn;
    }
    if (!Object.keys(patch).length) return cur;
    return saveSettings(patch);
}

/**
 * @param {ReturnType<typeof normalizeSettings>} settings
 * @param {{ audience?: string }} [opts]
 * @returns {Array<{ id: string, label: string, kind: 'me'|'user'|'group', upn: string, groupId?: string, visible: boolean }>}
 */
export function resolveEnabledSources(settings, opts) {
    const s = normalizeSettings(settings);
    const hasAudienceOpt = !!(opts && Object.prototype.hasOwnProperty.call(opts, 'audience'));
    const audience = hasAudienceOpt ? String(opts.audience || '').trim().toLowerCase() : '';
    // Explizit leere Audience (Persona noch unbekannt) → keine Layer
    if (hasAudienceOpt && !audience) return [];
    const enabled = new Set(s.enabledSources);
    const visible = new Set(s.visibleSources);
    /** @type {Array<{ id: string, label: string, kind: 'me'|'user'|'group', upn: string, groupId?: string, visible: boolean }>} */
    const out = [];

    function push(id, label, kind, upn, groupId, forceEnabled) {
        if (kind !== 'group' && !forceEnabled && !enabled.has(id)) return;
        if (kind !== 'group' && audience && !sourceAllowedForAudience(id, audience)) return;
        if (kind === 'user' && !upn) return;
        out.push({
            id,
            label,
            kind,
            upn: upn || '',
            groupId: groupId || '',
            visible:
                kind === 'group'
                    ? visible.has(id) || false
                    : forceEnabled && !enabled.has(id)
                      ? true
                      : visible.has(id)
        });
    }

    // Eigener Kalender immer für jede Audience – auch wenn IT „Mein Kalender“ deaktiviert hat
    push(SOURCE_IDS.ME, SOURCE_LABEL.me, 'me', '', undefined, !!audience);
    push(SOURCE_IDS.SCHUL, SOURCE_LABEL.schul, 'user', s.schulkalenderUpn);
    push(SOURCE_IDS.LEHRER, SOURCE_LABEL.lehrer, 'user', s.lehrerkalenderUpn);
    push(SOURCE_IDS.SCHUELER, SOURCE_LABEL.schueler, 'user', s.schuelerkalenderUpn);

    // Persönlich gewählte Gruppenkalender bleiben auch ohne Portal-Rolle sichtbar.
    (s.groupCalendars || []).forEach((g) => {
        const sid = groupSourceId(g.id);
        out.push({
            id: sid,
            label: g.displayName || g.mail || 'Gruppe',
            kind: 'group',
            upn: g.mail || '',
            groupId: g.id,
            visible: g.visible !== false
        });
    });

    if (out.length && !out.some((x) => x.visible)) {
        out.forEach((x) => {
            x.visible = true;
        });
    }
    return out;
}

/**
 * Termine auf aktuell sichtbare Layer filtern (rollenbewusst).
 * @param {object[]} events
 * @param {ReturnType<typeof normalizeSettings>|null} [settings]
 * @param {{ audience?: string }} [opts]
 */
export function filterEventsByVisibleSources(events, settings, opts) {
    const resolved = resolveEnabledSources(settings || loadSettings(), opts);
    const vis = new Set(resolved.filter((s) => s.visible).map((s) => s.id));
    const allowed = new Set(resolved.map((s) => s.id));
    return (events || []).filter((ev) => {
        const id = String((ev && ev.sourceId) || '').trim();
        if (!id) return true;
        if (allowed.size && !allowed.has(id)) return false;
        return vis.has(id);
    });
}
