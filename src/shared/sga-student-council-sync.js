/**
 * SGA-Schülervertretung (Board-Plätze) ↔ Schülervertretung (Stammdaten-Tab) koppeln.
 */

/**
 * @param {unknown} v
 */
function norm(v) {
    return String(v || '').trim();
}

/**
 * @param {Array<{ email?: string, name?: string, klasse?: string }>} allStudents
 * @param {{ name?: string, email?: string }} slot
 */
export function councilRowFromPerson(slot, allStudents) {
    const email = norm(slot && slot.email).toLowerCase();
    const name = norm(slot && slot.name);
    if (!email && !name) return null;
    const hit = (allStudents || []).find(
        (s) => email && norm(s && s.email).toLowerCase() === email
    );
    return {
        klasse: hit ? norm(hit.klasse) : '',
        name: (hit && norm(hit.name)) || name || '',
        email: email || (hit ? norm(hit.email).toLowerCase() : '')
    };
}

/**
 * SGA-Schüler-Slots → Schülervertretungs-Zeilen (erste Plätze).
 * Extra-Plätze der Schülervertretung (Index ≥ slotCount) bleiben erhalten.
 *
 * @param {Array<{ name?: string, email?: string }>} sgaStudentSlots
 * @param {Array<{ email?: string, name?: string, klasse?: string }>} allStudents
 * @param {Array<{ klasse?: string, name?: string, email?: string }>} [existingCouncil]
 * @param {number} [slotCount=3]
 */
export function studentCouncilFromSgaStudents(
    sgaStudentSlots,
    allStudents,
    existingCouncil,
    slotCount
) {
    const n = typeof slotCount === 'number' && slotCount > 0 ? slotCount : 3;
    const primary = [];
    const seen = new Set();
    (sgaStudentSlots || []).slice(0, n).forEach((slot) => {
        const row = councilRowFromPerson(slot, allStudents);
        if (!row || (!row.email && !row.name)) return;
        const key = row.email || row.name.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        primary.push(row);
    });

    const extras = [];
    (existingCouncil || []).slice(n).forEach((row) => {
        const email = norm(row && row.email).toLowerCase();
        const name = norm(row && row.name);
        if (!email && !name) return;
        const key = email || name.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        extras.push({
            klasse: norm(row && row.klasse),
            name,
            email
        });
    });

    return primary.concat(extras);
}

/**
 * Schülervertretung → SGA-Schüler-Slots (max. slotCount).
 *
 * @param {Array<{ name?: string, email?: string, klasse?: string }>} councilRows
 * @param {number} [slotCount=3]
 * @returns {Array<{ name: string, email: string }>}
 */
export function sgaStudentsFromStudentCouncil(councilRows, slotCount) {
    const n = typeof slotCount === 'number' && slotCount > 0 ? slotCount : 3;
    const out = [];
    for (const row of councilRows || []) {
        if (out.length >= n) break;
        const name = norm(row && row.name);
        const email = norm(row && row.email).toLowerCase();
        if (!name && !email) continue;
        out.push({ name, email });
    }
    while (out.length < n) out.push({ name: '', email: '' });
    return out;
}

/**
 * E-Mail-Sequenz der ersten n besetzten Plätze (Vergleich).
 * @param {Array<{ email?: string }>} rows
 * @param {number} [n=3]
 */
export function primaryEmails(rows, n) {
    const limit = typeof n === 'number' && n > 0 ? n : 3;
    return (rows || [])
        .map((r) => norm(r && r.email).toLowerCase())
        .filter((e) => e.includes('@'))
        .slice(0, limit);
}

/**
 * @param {string[]} a
 * @param {string[]} b
 */
export function sameEmailList(a, b) {
    const left = Array.isArray(a) ? a : [];
    const right = Array.isArray(b) ? b : [];
    if (left.length !== right.length) return false;
    return left.every((v, i) => v === right[i]);
}

/**
 * Beim Laden: welche Seite ist Quelle für die ersten Plätze?
 * @param {Array<{ email?: string }>} sgaStudents
 * @param {Array<{ email?: string }>} councilRows
 * @param {number} [slotCount=3]
 * @returns {'council'|'sga'|'none'}
 */
export function resolveStudentRepSyncSource(sgaStudents, councilRows, slotCount) {
    const n = typeof slotCount === 'number' && slotCount > 0 ? slotCount : 3;
    const sgaMails = primaryEmails(sgaStudents, n);
    const svMails = primaryEmails(councilRows, n);
    if (!sgaMails.length && !svMails.length) return 'none';
    if (sgaMails.length && !svMails.length) return 'sga';
    if (svMails.length && !sgaMails.length) return 'council';
    if (sameEmailList(sgaMails, svMails)) return 'none';
    // Beide befüllt und abweichend → Schülervertretung-Tab ist die kanonische Liste
    return 'council';
}
