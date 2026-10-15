/**
 * Schüler:innen aus Stammdaten für Kandidaturen / Wahlberechtigte.
 */

/**
 * @returns {{ name: string, className: string, email: string }[]}
 */
export function loadStudentPool() {
    try {
        const settings =
            typeof window !== 'undefined' &&
            window.ms365TenantSettingsLoad &&
            window.ms365TenantSettingsLoad();
        const list = settings && Array.isArray(settings.students) ? settings.students : [];
        return list
            .map((s) => ({
                name: String((s && (s.name || s.displayName)) || '').trim(),
                className: String((s && (s.klasse || s.className || s.class)) || '').trim(),
                email: String((s && (s.email || s.mail || s.upn)) || '')
                    .trim()
                    .toLowerCase()
            }))
            .filter((s) => s.name || s.email)
            .sort((a, b) => {
                const kc = a.className.localeCompare(b.className, 'de');
                if (kc) return kc;
                return a.name.localeCompare(b.name, 'de');
            });
    } catch {
        return [];
    }
}

/**
 * @param {string} email
 */
export function findStudentByEmail(email) {
    const key = String(email || '')
        .trim()
        .toLowerCase();
    if (!key) return null;
    return loadStudentPool().find((s) => s.email === key) || null;
}

/**
 * @param {string} query
 * @param {number} [limit]
 */
export function searchStudentPool(query, limit = 40) {
    const q = String(query || '')
        .trim()
        .toLowerCase();
    const all = loadStudentPool();
    if (!q) return all.slice(0, limit);
    return all
        .filter(
            (s) =>
                s.name.toLowerCase().includes(q) ||
                s.className.toLowerCase().includes(q) ||
                s.email.includes(q)
        )
        .slice(0, limit);
}
