/**
 * Stammdaten-Health: lokale Klasse vs. Entra department/officeLocation.
 * Kein DOM, kein Graph.
 */

function normStr(v) {
    return String(v == null ? '' : v).trim();
}

function normEmail(v) {
    return normStr(v).toLowerCase();
}

function normClass(v) {
    return normStr(v)
        .toLowerCase()
        .replace(/\s+/g, '');
}

function indexUsersByEmail(users) {
    const map = new Map();
    (Array.isArray(users) ? users : []).forEach(function (u) {
        if (!u || !u.id) return;
        [u.mail, u.userPrincipalName]
            .concat(Array.isArray(u.otherMails) ? u.otherMails : [])
            .forEach(function (raw) {
                const em = normEmail(raw);
                if (em && em.indexOf('@') !== -1 && !map.has(em)) map.set(em, u);
            });
    });
    return map;
}

function graphClassHint(user) {
    const dept = normStr(user && user.department);
    const office = normStr(user && user.officeLocation);
    // Prefer department; some schools put class in description — not always available
    return dept || office || '';
}

/**
 * @param {Array<{ name?: string, email?: string, klasse?: string }>} students
 * @param {object[]} graphUsers
 */
export function diffStudentAttributes(students, graphUsers) {
    const byEmail = indexUsersByEmail(graphUsers);
    const rows = [];
    let matched = 0;
    let mismatch = 0;
    let missingInGraph = 0;
    let missingClassLocal = 0;
    let ok = 0;

    (Array.isArray(students) ? students : []).forEach(function (s) {
        const em = normEmail(s && s.email);
        const localClass = normStr(s && s.klasse);
        if (!em || em.indexOf('@') === -1) return;
        const u = byEmail.get(em);
        if (!u) {
            missingInGraph++;
            rows.push({
                email: em,
                name: normStr(s && s.name),
                localClass,
                graphClass: '',
                status: 'missing_graph',
                message: 'Nicht im Tenant-Verzeichnis gefunden'
            });
            return;
        }
        matched++;
        const graphClass = graphClassHint(u);
        if (!localClass) {
            missingClassLocal++;
            rows.push({
                email: em,
                name: normStr(s && s.name) || normStr(u.displayName),
                localClass: '',
                graphClass,
                status: 'missing_local',
                message: 'Keine Klasse in Stammdaten'
            });
            return;
        }
        if (!graphClass) {
            rows.push({
                email: em,
                name: normStr(s && s.name) || normStr(u.displayName),
                localClass,
                graphClass: '',
                status: 'missing_ad_field',
                message: 'department/officeLocation leer (AD-Feld?)'
            });
            return;
        }
        if (normClass(localClass) !== normClass(graphClass)) {
            mismatch++;
            rows.push({
                email: em,
                name: normStr(s && s.name) || normStr(u.displayName),
                localClass,
                graphClass,
                status: 'mismatch',
                message: 'Stammdaten „' + localClass + '“ ≠ Entra „' + graphClass + '“'
            });
            return;
        }
        ok++;
    });

    return {
        rows,
        summary: {
            students: (Array.isArray(students) ? students : []).length,
            matched,
            ok,
            mismatch,
            missingInGraph,
            missingClassLocal,
            problems: rows.filter((r) => r.status !== 'ok').length
        }
    };
}

export default { diffStudentAttributes };
