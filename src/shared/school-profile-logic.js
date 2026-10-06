/** Max. Größe Logo-Data-URL (ca. 400 KB Bild) – Schutz für localStorage/Backup */
export const SCHOOL_LOGO_MAX_BYTES = 400 * 1024;

export function emptySchoolProfile() {
    return {
        schoolCode: '',
        logoDataUrl: '',
        logoFileName: '',
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
