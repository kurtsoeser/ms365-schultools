/** Max. Größe Logo-Data-URL (ca. 400 KB Bild) – Schutz für localStorage/Backup */
export const SCHOOL_LOGO_MAX_BYTES = 400 * 1024;

/** Max. Länge für öffentliche Logo-URL */
export const SCHOOL_LOGO_URL_MAX_LEN = 2000;

export function emptySchoolProfile() {
    return {
        schoolCode: '',
        logoDataUrl: '',
        logoFileName: '',
        /** Öffentliche HTTPS-URL zum Schullogo (bevorzugt gegenüber Data-URL) */
        logoUrl: '',
        street: '',
        postalCode: '',
        city: '',
        country: 'Österreich',
        phone: '',
        phoneAlt: '',
        email: '',
        website: ''
    };
}

/**
 * Nur öffentliche https://-URLs (kein http, keine data:, keine relativen Pfade).
 * @param {unknown} raw
 * @returns {string}
 */
export function normalizeSchoolLogoUrl(raw) {
    const s = String(raw != null ? raw : '').trim();
    if (!s || s.length > SCHOOL_LOGO_URL_MAX_LEN) return '';
    if (!/^https:\/\//i.test(s)) return '';
    try {
        const u = new URL(s);
        if (u.protocol !== 'https:') return '';
        if (!u.hostname) return '';
        return u.toString();
    } catch {
        return '';
    }
}

/**
 * Anzeige-URL: öffentliche HTTPS-URL vor eingebettetem Upload.
 * @param {unknown} profile
 * @returns {string}
 */
export function resolveSchoolLogoSrc(profile) {
    const p = normalizeSchoolProfile(profile);
    if (p.logoUrl) return p.logoUrl;
    if (p.logoDataUrl) return p.logoDataUrl;
    return '';
}

/**
 * @param {unknown} profile
 * @returns {boolean}
 */
export function hasSchoolLogo(profile) {
    return !!resolveSchoolLogoSrc(profile);
}

/**
 * @param {unknown} raw
 * @returns {ReturnType<typeof emptySchoolProfile>}
 */
export function normalizeSchoolProfile(raw) {
    const x = raw && typeof raw === 'object' ? raw : {};
    const d = emptySchoolProfile();
    d.schoolCode = String(x.schoolCode != null ? x.schoolCode : '').trim();
    let logo = String(x.logoDataUrl != null ? x.logoDataUrl : '').trim();
    if (logo && !/^data:image\/(png|jpeg|jpg|gif|webp);base64,/i.test(logo)) {
        logo = '';
    }
    if (logo.length > SCHOOL_LOGO_MAX_BYTES * 1.4) {
        logo = '';
    }
    d.logoDataUrl = logo;
    d.logoFileName = String(x.logoFileName != null ? x.logoFileName : '').trim();
    d.logoUrl = normalizeSchoolLogoUrl(x.logoUrl);
    d.street = String(x.street != null ? x.street : '').trim();
    d.postalCode = String(x.postalCode != null ? x.postalCode : '').trim();
    d.city = String(x.city != null ? x.city : '').trim();
    d.country = String(x.country != null ? x.country : '').trim() || 'Österreich';
    d.phone = String(x.phone != null ? x.phone : '').trim();
    d.phoneAlt = String(x.phoneAlt != null ? x.phoneAlt : '').trim();
    d.email = String(x.email != null ? x.email : '').trim().toLowerCase();
    d.website = String(x.website != null ? x.website : '').trim();
    return d;
}

/**
 * @param {ReturnType<typeof emptySchoolProfile>} profile
 */
export function schoolProfileToPatch(profile) {
    return { schoolProfile: normalizeSchoolProfile(profile) };
}

/**
 * Einzeilige Adresse für Dashboard/Ausdrucke.
 * @param {ReturnType<typeof emptySchoolProfile>} profile
 */
export function formatSchoolAddressLine(profile) {
    const p = normalizeSchoolProfile(profile);
    const cityLine = [p.postalCode, p.city].filter(Boolean).join(' ').trim();
    return [p.street, cityLine].filter(Boolean).join(', ');
}

/**
 * Website-URL mit https:// falls nötig; sonst leer.
 * @param {string} raw
 */
export function normalizeWebsiteHref(raw) {
    const s = String(raw || '').trim();
    if (!s) return '';
    if (/^https?:\/\//i.test(s)) return s;
    if (/^[\w.-]+\.[\w.-]+(\/.*)?$/i.test(s)) return 'https://' + s;
    return '';
}

/**
 * Anzeige ohne Protokoll.
 * @param {string} raw
 */
export function formatWebsiteLabel(raw) {
    const href = normalizeWebsiteHref(raw) || String(raw || '').trim();
    return href.replace(/^https?:\/\//i, '').replace(/\/$/, '');
}

/**
 * Schulidentität für Portal-Dashboard (Name + Profil aus Stammdaten).
 * @param {{ schoolName?: string, domain?: string, schoolProfile?: unknown }|null} [input]
 */
export function buildDashboardSchoolCard(input) {
    const o = input && typeof input === 'object' ? input : {};
    const schoolName = String(o.schoolName || '').trim();
    const domain = String(o.domain || '').trim();
    const profile = normalizeSchoolProfile(o.schoolProfile);
    const addressLine = formatSchoolAddressLine(profile);
    const websiteHref = normalizeWebsiteHref(profile.website);
    const websiteLabel = formatWebsiteLabel(profile.website);
    const logoSrc = resolveSchoolLogoSrc(profile);
    return {
        schoolName: schoolName || (domain ? domain : ''),
        domain,
        profile,
        addressLine,
        websiteHref,
        websiteLabel,
        logoSrc,
        hasLogo: !!logoSrc,
        hasIdentity: !!(schoolName || domain || logoSrc || addressLine || websiteHref)
    };
}
