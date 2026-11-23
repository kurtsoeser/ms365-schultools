/**
 * Reine Auswertung der Betreiber-Schulliste: Ablauf, Kontakte, Paket-Mix.
 */

export function parseValidUntil(value) {
    const s = String(value || '')
        .trim()
        .slice(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return null;
    const parts = s.split('-').map(Number);
    const dt = new Date(parts[0], parts[1] - 1, parts[2]);
    dt.setHours(0, 0, 0, 0);
    if (dt.getFullYear() !== parts[0] || dt.getMonth() !== parts[1] - 1 || dt.getDate() !== parts[2]) {
        return null;
    }
    return dt;
}

export function startOfDay(date) {
    const d = new Date(date);
    d.setHours(0, 0, 0, 0);
    return d;
}

export function addDays(date, days) {
    const d = startOfDay(date);
    d.setDate(d.getDate() + days);
    return d;
}

function extraOf(school) {
    return (school && school.extra && typeof school.extra === 'object' && school.extra) || {};
}

/**
 * @param {object} school
 * @returns {{ role: string, email: string }[]}
 */
export function collectContacts(school) {
    const rows = [];
    const seen = new Set();
    function add(role, raw) {
        const found = String(raw || '').match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi) || [];
        found.forEach((email) => {
            const key = role + '|' + email.toLowerCase();
            if (seen.has(key)) return;
            seen.add(key);
            rows.push({ role, email: email.toLowerCase() });
        });
    }
    const list = school && Array.isArray(school.contactEmails) ? school.contactEmails : [];
    if (list.length) list.forEach((entry) => add('Lizenz', entry));
    else add('Lizenz', school && school.contactEmail);
    const ex = extraOf(school);
    add('IT', ex.ItM365AdminEmail || ex.ItM365Admin);
    add('Direktion', ex.DirektionEmail || ex.Direktion);
    add('Schuladministration', ex.SchuladministrationEmail || ex.Schuladministration);
    add('Kontakt', ex.WeitereKontakte);
    return rows;
}

export function hasContact(school) {
    return collectContacts(school).length > 0;
}

export function isExpired(school, today) {
    const status = String((school && school.status) || '').toLowerCase();
    if (status === 'expired') return true;
    const until = parseValidUntil(school && school.validUntil);
    if (!until) return false;
    return until.getTime() < startOfDay(today).getTime();
}

export function isDueWithin(school, today, days) {
    const status = String((school && school.status) || '').toLowerCase();
    if (status === 'expired' || status === 'blocked') return false;
    const until = parseValidUntil(school && school.validUntil);
    if (!until) return false;
    const start = startOfDay(today).getTime();
    const end = addDays(today, days).getTime();
    const at = until.getTime();
    return at >= start && at <= end;
}

/**
 * @param {object[]} schools
 * @param {Date} today
 * @param {{ isPending?: (school: object) => boolean }} [opts]
 */
export function licenseCounts(schools, today, opts) {
    const isPending = opts && typeof opts.isPending === 'function' ? opts.isPending : () => false;
    const counts = {
        active: 0,
        trial: 0,
        due30: 0,
        expired: 0,
        pending: 0,
        nocontact: 0,
        total: (schools || []).length
    };
    (schools || []).forEach((school) => {
        const status = String(school.status || '').toLowerCase();
        if (status === 'active') counts.active += 1;
        if (status === 'trial') counts.trial += 1;
        if (isDueWithin(school, today, 30)) counts.due30 += 1;
        if (isExpired(school, today)) counts.expired += 1;
        if (isPending(school)) counts.pending += 1;
        if (!hasContact(school)) counts.nocontact += 1;
    });
    return counts;
}

export function packagesFromSchool(school) {
    const doc = school && (school.entitlementsDoc || school.entitlements);
    if (!doc || !Array.isArray(doc.packages)) return [];
    return doc.packages.map((id) => String(id || '').trim()).filter(Boolean);
}

/**
 * @param {object[]} schools
 * @param {{ id: string, label: string }[]} packages
 * @param {(ids: string[]) => { bundle: { id: string, label: string }|null, exact: boolean }} matchBundle
 */
export function adoptionCounts(schools, packages, matchBundle) {
    /** @type {Record<string, number>} */
    const byPackage = {};
    (packages || []).forEach((pkg) => {
        byPackage[pkg.id] = 0;
    });
    /** @type {Record<string, { id: string, label: string, count: number }>} */
    const byBundle = {};
    let withPackages = 0;
    let custom = 0;
    (schools || []).forEach((school) => {
        const ids = packagesFromSchool(school);
        if (!ids.length) return;
        withPackages += 1;
        ids.forEach((id) => {
            if (byPackage[id] == null) byPackage[id] = 0;
            byPackage[id] += 1;
        });
        const matched = typeof matchBundle === 'function' ? matchBundle(ids) : { bundle: null, exact: false };
        if (matched && matched.bundle) {
            const id = matched.bundle.id;
            if (!byBundle[id]) {
                byBundle[id] = { id, label: matched.bundle.label || id, count: 0 };
            }
            byBundle[id].count += 1;
        } else {
            custom += 1;
        }
    });
    return { byPackage, byBundle, withPackages, custom };
}

export function prepareMailHref(emails, subject, body) {
    const list = (emails || []).map((entry) => String(entry || '').trim()).filter(Boolean);
    if (!list.length) return '';
    return (
        'mailto:' +
        list.join(',') +
        '?subject=' +
        encodeURIComponent(subject || '') +
        '&body=' +
        encodeURIComponent(body || '')
    );
}
