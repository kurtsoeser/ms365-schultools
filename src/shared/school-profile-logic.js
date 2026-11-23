/** Max. Dateigröße vor Kompression (Upload) */
export const SCHOOL_LOGO_MAX_BYTES = 2 * 1024 * 1024;
/** Zielgröße der gespeicherten Data-URL (Zeichen) – muss in localStorage/Backup passen */
export const SCHOOL_LOGO_DATA_URL_MAX_CHARS = 420000;

/** Max. Länge für öffentliche Logo-URL */
export const SCHOOL_LOGO_URL_MAX_LEN = 2000;

export function emptySchoolProfile() {
    return {
        schoolCode: '',
        logoDataUrl: '',
        logoFileName: '',
        /** Öffentliche HTTPS-URL zum Schullogo (bevorzugt gegenüber Data-URL) */
        logoUrl: '',
        /** Lizenz-Admin: Standort (kann vom Ort abweichen) */
        standort: '',
        street: '',
        postalCode: '',
        city: '',
        country: 'Österreich',
        phone: '',
        phoneAlt: '',
        email: '',
        website: '',
        /** Rollen / Personen (Name + E-Mail getrennt) */
        direktion: '',
        direktionEmail: '',
        schuladministration: '',
        schuladministrationEmail: '',
        itM365Admin: '',
        itM365AdminEmail: '',
        weitereKontakte: ''
    };
}

/**
 * Alte Einzelfelder („Name“, „mail@…“ oder „Name <mail>“) in Name + E-Mail trennen.
 * @param {unknown} nameRaw
 * @param {unknown} emailRaw
 * @returns {{ name: string, email: string }}
 */
export function splitRolePersonFields(nameRaw, emailRaw) {
    let name = String(nameRaw != null ? nameRaw : '').trim();
    let email = String(emailRaw != null ? emailRaw : '')
        .trim()
        .toLowerCase();
    if (!email && name) {
        const angled = name.match(/^(.+?)\s*<\s*([^>\s]+@[^>\s]+)\s*>$/);
        if (angled) {
            name = String(angled[1] || '').trim();
            email = String(angled[2] || '')
                .trim()
                .toLowerCase();
        } else if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(name)) {
            email = name.toLowerCase();
            name = '';
        }
    } else if (email) {
        email = email.toLowerCase();
    }
    return { name: name, email: email };
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
 * Lädt/komprimiert ein Bild fürs Speichern in localStorage + Browser-Backup / IT-Bibliothek.
 * @param {Blob|File} file
 * @param {{ maxEdge?: number, maxChars?: number }} [opts]
 * @returns {Promise<{ dataUrl: string, fileName: string }>}
 */
export function compressSchoolLogoToDataUrl(file, opts) {
    const maxEdge = opts && opts.maxEdge ? Number(opts.maxEdge) : 640;
    const maxChars =
        opts && opts.maxChars ? Number(opts.maxChars) : SCHOOL_LOGO_DATA_URL_MAX_CHARS;
    return new Promise(function (resolve, reject) {
        if (!file) {
            reject(new Error('no-file'));
            return;
        }
        const reader = new FileReader();
        reader.onerror = function () {
            reject(new Error('read'));
        };
        reader.onload = function () {
            const rawUrl = String(reader.result || '');
            const img = new Image();
            img.onerror = function () {
                reject(new Error('decode'));
            };
            img.onload = function () {
                try {
                    let w = img.naturalWidth || img.width || 1;
                    let h = img.naturalHeight || img.height || 1;
                    const scale = Math.min(1, maxEdge / Math.max(w, h, 1));
                    w = Math.max(1, Math.round(w * scale));
                    h = Math.max(1, Math.round(h * scale));
                    const canvas = document.createElement('canvas');
                    canvas.width = w;
                    canvas.height = h;
                    const ctx = canvas.getContext('2d');
                    if (!ctx) {
                        reject(new Error('canvas'));
                        return;
                    }
                    ctx.drawImage(img, 0, 0, w, h);
                    let quality = 0.88;
                    let dataUrl = canvas.toDataURL('image/jpeg', quality);
                    while (dataUrl.length > maxChars && quality > 0.42) {
                        quality -= 0.08;
                        dataUrl = canvas.toDataURL('image/jpeg', quality);
                    }
                    if (dataUrl.length > maxChars) {
                        reject(new Error('too-large'));
                        return;
                    }
                    const base = String(file.name || 'schullogo')
                        .replace(/\.[^.]+$/, '')
                        .trim() || 'schullogo';
                    resolve({ dataUrl: dataUrl, fileName: base + '.jpg' });
                } catch (err) {
                    reject(err || new Error('compress'));
                }
            };
            img.src = rawUrl;
        };
        reader.readAsDataURL(file);
    });
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
    if (logo.length > SCHOOL_LOGO_DATA_URL_MAX_CHARS) {
        logo = '';
    }
    d.logoDataUrl = logo;
    d.logoFileName = String(x.logoFileName != null ? x.logoFileName : '').trim();
    d.logoUrl = normalizeSchoolLogoUrl(x.logoUrl);
    d.standort = String(x.standort != null ? x.standort : '').trim();
    d.street = String(x.street != null ? x.street : '').trim();
    d.postalCode = String(x.postalCode != null ? x.postalCode : '').trim();
    d.city = String(x.city != null ? x.city : '').trim();
    d.country = String(x.country != null ? x.country : '').trim() || 'Österreich';
    d.phone = String(x.phone != null ? x.phone : '').trim();
    d.phoneAlt = String(x.phoneAlt != null ? x.phoneAlt : '').trim();
    d.email = String(x.email != null ? x.email : '').trim().toLowerCase();
    d.website = String(x.website != null ? x.website : '').trim();
    const dir = splitRolePersonFields(x.direktion, x.direktionEmail);
    d.direktion = dir.name;
    d.direktionEmail = dir.email;
    const adm = splitRolePersonFields(x.schuladministration, x.schuladministrationEmail);
    d.schuladministration = adm.name;
    d.schuladministrationEmail = adm.email;
    const it = splitRolePersonFields(x.itM365Admin, x.itM365AdminEmail);
    d.itM365Admin = it.name;
    d.itM365AdminEmail = it.email;
    d.weitereKontakte = String(x.weitereKontakte != null ? x.weitereKontakte : '').trim();
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
