/**
 * Schulweite Schnellzugriff-Links für das Lehrer-/Schüler-Portal.
 * Speicherung: setup.portalQuickLinks (App-Data / SharePoint-Sync) mit localStorage-Fallback.
 * Keine Betreiber-DB – Multi-School-konform.
 */

export const PORTAL_QUICK_LINKS_STORAGE_KEY = 'ms365-portal-quick-links-v1';

/** @typedef {{ id: string, label: string, url: string, icon: string, external?: boolean }} PortalQuickLink */

/** @type {PortalQuickLink[]} */
export const PORTAL_QUICK_LINK_DEFAULTS = [
    {
        id: 'outlook',
        label: 'Outlook',
        url: 'https://outlook.office.com/mail/',
        icon: 'bi-envelope',
        external: true
    },
    {
        id: 'teams',
        label: 'Teams',
        url: 'https://teams.microsoft.com/',
        icon: 'bi-chat-dots',
        external: true
    },
    {
        id: 'kalender',
        label: 'Kalender',
        url: 'https://outlook.office.com/calendar/',
        icon: 'bi-calendar3',
        external: true
    },
    {
        id: 'webuntis',
        label: 'WebUntis',
        url: '',
        icon: 'bi-clock-history',
        external: true
    },
    {
        id: 'sokrates',
        label: 'Sokrates',
        url: '',
        icon: 'bi-journal-bookmark',
        external: true
    },
    {
        id: 'intranet',
        label: 'Intranet',
        url: '',
        icon: 'bi-house-door',
        external: true
    }
];

/**
 * @param {unknown} raw
 * @returns {PortalQuickLink[]}
 */
export function normalizePortalQuickLinks(raw) {
    const byId = new Map(PORTAL_QUICK_LINK_DEFAULTS.map((d) => [d.id, { ...d }]));
    const list = Array.isArray(raw) ? raw : [];
    list.forEach((item) => {
        if (!item || typeof item !== 'object') return;
        const id = String(item.id || '')
            .trim()
            .toLowerCase();
        if (!id || !byId.has(id)) return;
        const cur = byId.get(id);
        const url = String(item.url || '').trim();
        const label = String(item.label || '').trim();
        byId.set(id, {
            id,
            label: label || cur.label,
            url,
            icon: String(item.icon || cur.icon || 'bi-link-45deg').trim() || cur.icon,
            external: item.external !== false
        });
    });
    return PORTAL_QUICK_LINK_DEFAULTS.map((d) => byId.get(d.id) || { ...d });
}

/**
 * @returns {string}
 */
function currentTenantId() {
    try {
        if (typeof window !== 'undefined' && typeof window.ms365AuthGetAccountInfo === 'function') {
            const info = window.ms365AuthGetAccountInfo();
            return info && info.tenantId ? String(info.tenantId).trim() : '';
        }
    } catch {
        /* ignore */
    }
    return '';
}

/**
 * @param {Storage} [storage]
 * @returns {PortalQuickLink[]|null}
 */
function loadFromLocalStorage(storage) {
    try {
        const store = storage || (typeof localStorage !== 'undefined' ? localStorage : null);
        if (!store) return null;
        const raw = store.getItem(PORTAL_QUICK_LINKS_STORAGE_KEY);
        if (!raw) return null;
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) return normalizePortalQuickLinks(parsed);
        if (parsed && typeof parsed === 'object') {
            const tid = currentTenantId();
            if (tid && Array.isArray(parsed[tid])) return normalizePortalQuickLinks(parsed[tid]);
            if (Array.isArray(parsed.links)) return normalizePortalQuickLinks(parsed.links);
        }
    } catch {
        /* ignore */
    }
    return null;
}

/**
 * @returns {PortalQuickLink[]|null}
 */
function loadFromAppData() {
    try {
        if (typeof window === 'undefined' || !window.ms365AppDataV2) return null;
        const setup =
            typeof window.ms365AppDataV2.getSetup === 'function'
                ? window.ms365AppDataV2.getSetup()
                : null;
        if (setup && Array.isArray(setup.portalQuickLinks)) {
            return normalizePortalQuickLinks(setup.portalQuickLinks);
        }
    } catch {
        /* ignore */
    }
    return null;
}

/**
 * Intranet-URL aus Stammdaten, falls Portal-Link noch leer.
 * @param {PortalQuickLink[]} links
 * @returns {PortalQuickLink[]}
 */
function enrichIntranetFromSetup(links) {
    try {
        if (typeof window === 'undefined' || !window.ms365AppDataV2) return links;
        const setup =
            typeof window.ms365AppDataV2.getSetup === 'function'
                ? window.ms365AppDataV2.getSetup()
                : null;
        const intranetUrl = String(
            (setup && (setup.schoolIntranetSiteUrl || setup.intranetSiteUrl)) || ''
        ).trim();
        if (!intranetUrl) return links;
        return links.map((l) => {
            if (l.id !== 'intranet' || String(l.url || '').trim()) return l;
            return { ...l, url: intranetUrl };
        });
    } catch {
        return links;
    }
}

/**
 * @param {Storage} [storage]
 * @returns {PortalQuickLink[]}
 */
export function loadPortalQuickLinks(storage) {
    const fromSetup = loadFromAppData();
    if (fromSetup) return enrichIntranetFromSetup(fromSetup);
    const fromLocal = loadFromLocalStorage(storage);
    if (fromLocal) return enrichIntranetFromSetup(fromLocal);
    return enrichIntranetFromSetup(normalizePortalQuickLinks([]));
}

/**
 * @param {PortalQuickLink[]} links
 * @param {Storage} [storage]
 */
export function savePortalQuickLinks(links, storage) {
    const normalized = normalizePortalQuickLinks(links);
    let savedToSetup = false;
    try {
        if (
            typeof window !== 'undefined' &&
            window.ms365AppDataV2 &&
            typeof window.ms365AppDataV2.patchSetup === 'function'
        ) {
            window.ms365AppDataV2.patchSetup({ portalQuickLinks: normalized });
            savedToSetup = true;
            try {
                window.dispatchEvent(
                    new CustomEvent('ms365-tenant-settings-changed', {
                        detail: { reason: 'autosave', source: 'portal-quick-links' }
                    })
                );
            } catch {
                /* ignore */
            }
        }
    } catch {
        /* ignore */
    }

    try {
        const store = storage || (typeof localStorage !== 'undefined' ? localStorage : null);
        if (!store) return normalized;
        const tid = currentTenantId();
        let payload;
        if (tid) {
            let map = {};
            try {
                const prev = JSON.parse(store.getItem(PORTAL_QUICK_LINKS_STORAGE_KEY) || '{}');
                if (prev && typeof prev === 'object' && !Array.isArray(prev)) map = prev;
            } catch {
                map = {};
            }
            map[tid] = normalized;
            payload = map;
        } else {
            payload = { links: normalized };
        }
        store.setItem(PORTAL_QUICK_LINKS_STORAGE_KEY, JSON.stringify(payload));
    } catch {
        /* ignore */
    }

    if (!savedToSetup) {
        /* local-only is fine for P0 */
    }
    return normalized;
}

/**
 * Nur Links mit URL (konfigurierte + Defaults mit fester Office-URL).
 * @param {PortalQuickLink[]} [links]
 * @returns {PortalQuickLink[]}
 */
export function activePortalQuickLinks(links) {
    return (links || loadPortalQuickLinks()).filter((l) => String(l.url || '').trim());
}

/**
 * Schulspezifische IDs, die oft leer starten.
 * @type {string[]}
 */
export const PORTAL_QUICK_LINK_SCHOOL_IDS = ['webuntis', 'sokrates', 'intranet'];
