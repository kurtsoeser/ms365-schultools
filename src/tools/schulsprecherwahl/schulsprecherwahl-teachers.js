/**
 * Lehrer:innen aus Stammdaten für SGA-Kandidaturen / Wahlberechtigte.
 */

/**
 * @returns {{ name: string, className: string, email: string, code: string }[]}
 */
export function loadTeacherPool() {
    try {
        /** @type {unknown[]} */
        let list = [];
        if (
            typeof window !== 'undefined' &&
            window.ms365AppDataV2 &&
            typeof window.ms365AppDataV2.getTeachers === 'function'
        ) {
            list = window.ms365AppDataV2.getTeachers() || [];
        } else if (
            typeof window !== 'undefined' &&
            window.ms365TenantSettingsLoad
        ) {
            const settings = window.ms365TenantSettingsLoad();
            list = settings && Array.isArray(settings.teachers) ? settings.teachers : [];
        }
        return list
            .map((t) => {
                const o = t && typeof t === 'object' ? t : {};
                return {
                    name: String(o.name || o.displayName || '').trim(),
                    className: String(o.code || o.kuerzel || o.shortName || '').trim(),
                    email: String(o.email || o.mail || o.upn || '')
                        .trim()
                        .toLowerCase(),
                    code: String(o.code || o.kuerzel || o.id || '').trim()
                };
            })
            .filter((t) => t.name || t.email)
            .sort((a, b) => a.name.localeCompare(b.name, 'de'));
    } catch {
        return [];
    }
}

/**
 * @param {string} email
 */
export function findTeacherByEmail(email) {
    const key = String(email || '')
        .trim()
        .toLowerCase();
    if (!key) return null;
    return loadTeacherPool().find((t) => t.email === key) || null;
}

/**
 * @param {string} query
 * @param {number} [limit]
 */
export function searchTeacherPool(query, limit = 40) {
    const q = String(query || '')
        .trim()
        .toLowerCase();
    const all = loadTeacherPool();
    if (!q) return all.slice(0, limit);
    return all
        .filter(
            (t) =>
                t.name.toLowerCase().includes(q) ||
                t.className.toLowerCase().includes(q) ||
                t.code.toLowerCase().includes(q) ||
                t.email.includes(q)
        )
        .slice(0, limit);
}
