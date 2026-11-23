/**
 * Schul-Stammdaten (Browser): Spiegel von
 * backend/license-api/src/lib/school-stammdaten-catalog.js
 */

/**
 * @typedef {{
 *   name: string,
 *   displayName: string,
 *   type: 'text'|'multiline',
 *   group: 'adresse'|'rollen'|'kontakt'|'meta'|'workshop',
 *   tableDefault?: boolean,
 *   excelKeys?: string[]
 * }} SchoolStammdatenField
 */

/** @type {SchoolStammdatenField[]} */
export const SCHOOL_STAMMDATEN_FIELDS = [
    {
        name: 'Standort',
        displayName: 'Standort',
        type: 'text',
        group: 'adresse',
        tableDefault: true,
        excelKeys: ['Standort']
    },
    {
        name: 'Strasse',
        displayName: 'Straße',
        type: 'text',
        group: 'adresse',
        tableDefault: false,
        excelKeys: ['Straße', 'Strasse']
    },
    {
        name: 'PLZ',
        displayName: 'PLZ',
        type: 'text',
        group: 'adresse',
        tableDefault: false,
        excelKeys: ['PLZ']
    },
    {
        name: 'Ort',
        displayName: 'Ort',
        type: 'text',
        group: 'adresse',
        tableDefault: true,
        excelKeys: ['Ort']
    },
    {
        name: 'Direktion',
        displayName: 'Direktion',
        type: 'text',
        group: 'rollen',
        pairKey: 'direktion',
        pairPart: 'name',
        tableDefault: true,
        excelKeys: ['Direktion']
    },
    {
        name: 'DirektionEmail',
        displayName: 'Direktion E-Mail',
        type: 'text',
        group: 'rollen',
        pairKey: 'direktion',
        pairPart: 'email',
        pairLabel: 'Direktion',
        tableDefault: false,
        excelKeys: ['Direktion E-Mail', 'Direktion Email']
    },
    {
        name: 'Schuladministration',
        displayName: 'Schuladministration',
        type: 'text',
        group: 'rollen',
        pairKey: 'schuladmin',
        pairPart: 'name',
        tableDefault: false,
        excelKeys: ['Schuladministration']
    },
    {
        name: 'SchuladministrationEmail',
        displayName: 'Schuladministration E-Mail',
        type: 'text',
        group: 'rollen',
        pairKey: 'schuladmin',
        pairPart: 'email',
        pairLabel: 'Schuladministration',
        tableDefault: false,
        excelKeys: ['Schuladministration E-Mail', 'Schuladministration Email']
    },
    {
        name: 'ItM365Admin',
        displayName: 'IT-/M365-Administration',
        type: 'text',
        group: 'rollen',
        pairKey: 'itadmin',
        pairPart: 'name',
        tableDefault: false,
        excelKeys: ['IT-/M365-Administration', 'IT/M365-Administration']
    },
    {
        name: 'ItM365AdminEmail',
        displayName: 'IT-/M365-Administration E-Mail',
        type: 'text',
        group: 'rollen',
        pairKey: 'itadmin',
        pairPart: 'email',
        pairLabel: 'IT-/M365-Administration',
        tableDefault: false,
        excelKeys: ['IT-/M365-Administration E-Mail', 'IT/M365-Administration E-Mail']
    },
    {
        name: 'WeitereKontakte',
        displayName: 'Weitere Kontakte / Rollen',
        type: 'multiline',
        group: 'rollen',
        tableDefault: false,
        excelKeys: ['Weitere Kontakte / Rollen']
    },
    {
        name: 'Telefon',
        displayName: 'Telefon',
        type: 'text',
        group: 'kontakt',
        tableDefault: true,
        excelKeys: ['Telefon']
    },
    {
        name: 'Schulnummer',
        displayName: 'Schulnummer',
        type: 'text',
        group: 'kontakt',
        tableDefault: true,
        excelKeys: ['Schulnummer(n)', 'Schulnummer']
    },
    {
        name: 'DomainHinweis',
        displayName: 'Domain-Hinweis',
        type: 'multiline',
        group: 'meta',
        tableDefault: false,
        excelKeys: ['Domain-Beleg / Hinweis']
    },
    {
        name: 'Datenstatus',
        displayName: 'Datenstatus',
        type: 'text',
        group: 'meta',
        tableDefault: false,
        excelKeys: ['Datenstatus']
    },
    {
        name: 'WorkshopTeilnehmende',
        displayName: 'Workshop-Teilnehmende',
        type: 'multiline',
        group: 'workshop',
        tableDefault: false,
        excelKeys: ['Workshop-Teilnehmende']
    },
    {
        name: 'WorkshopQuelle',
        displayName: 'Workshop-Quelle',
        type: 'multiline',
        group: 'workshop',
        tableDefault: false,
        excelKeys: ['Interne Workshopquelle']
    },
    {
        name: 'OeffentlicheQuellen',
        displayName: 'Öffentliche Quellen',
        type: 'multiline',
        group: 'workshop',
        tableDefault: false,
        excelKeys: ['Öffentliche Quelle(n)', 'Oeffentliche Quelle(n)']
    },
    {
        name: 'OffenePunkte',
        displayName: 'Offene Punkte',
        type: 'multiline',
        group: 'workshop',
        tableDefault: false,
        excelKeys: ['Offene Punkte / Prüfbedarf', 'Offene Punkte']
    }
];

/** @type {Record<string, string>} */
export const SCHOOL_STAMMDATEN_GROUPS = {
    adresse: 'Adresse',
    rollen: 'Rollen & Personen',
    kontakt: 'Kontakt & Kennzahlen',
    meta: 'Status & Domain',
    workshop: 'Workshop & Quellen'
};

export const TENANT_GUID_RE =
    /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
export const PENDING_TENANT_RE = /^PENDING(?:-[A-Za-z0-9_-]{1,48})?$/i;

/** @param {string} [raw] */
export function isRealTenantGuid(raw) {
    return TENANT_GUID_RE.test(String(raw || '').trim());
}

/** @param {string} [raw] */
export function isPendingTenantId(raw) {
    const s = String(raw || '').trim();
    if (!s) return true;
    if (/^offen\b/i.test(s)) return true;
    return PENDING_TENANT_RE.test(s);
}

/** @param {string} [schoolName] */
export function makePendingTenantId(schoolName) {
    let slug = String(schoolName || 'schule')
        .replace(/ä/gi, 'ae')
        .replace(/ö/gi, 'oe')
        .replace(/ü/gi, 'ue')
        .replace(/ß/g, 'ss')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/[^a-zA-Z0-9]+/g, '-')
        .replace(/^-+|-+$/g, '')
        .toLowerCase();
    if (!slug) slug = 'schule';
    // Lange Namen: Kurzkopf + letztes Token (meist Ort), damit Perg/Gmunden nicht kollidieren
    if (slug.length > 40) {
        const parts = slug.split('-').filter(Boolean);
        const last = parts[parts.length - 1] || 'schule';
        const head = (parts[0] || 'schule').slice(0, 22);
        slug = (head + '-' + last).slice(0, 40);
    }
    return 'PENDING-' + slug;
}

/**
 * @param {string} [raw]
 * @param {string} [schoolName]
 */
export function normalizeTenantIdInput(raw, schoolName) {
    const s = String(raw || '').trim();
    if (!s || /^offen\b/i.test(s) || /^pending$/i.test(s) || s.toUpperCase() === 'PENDING') {
        return makePendingTenantId(schoolName);
    }
    return s;
}

/**
 * @param {string} [raw]
 * @param {string} [schoolName]
 */
export function assertTenantIdAllowed(raw, schoolName) {
    const normalized = normalizeTenantIdInput(raw, schoolName);
    if (isRealTenantGuid(normalized) || PENDING_TENANT_RE.test(normalized)) {
        return normalized;
    }
    throw new Error(
        'Tenant-ID muss eine GUID sein oder mit PENDING- beginnen (Platzhalter bis die echte ID bekannt ist).'
    );
}

/** @param {string} [name] */
export function stammdatenFieldByName(name) {
    const want = String(name || '')
        .trim()
        .toLowerCase();
    return SCHOOL_STAMMDATEN_FIELDS.find((f) => String(f.name).toLowerCase() === want) || null;
}
