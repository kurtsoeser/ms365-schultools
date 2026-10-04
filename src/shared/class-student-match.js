/**
 * Zuordnung Schüler:innen (Feld klasse) zu Klassenzeilen / verknüpften Gruppen.
 */
import { normStr, normCode } from './utils/strings.js';

/** Erstes Klassenkürzel aus Freitext (z. B. „1A Demo“ → 1A). */
export function classLabelHeadToken(label) {
    const s = normStr(label);
    if (!s) return '';
    const m = s.match(/^(\d+\s*[A-Za-zÄÖÜäöüß]+)/);
    if (m) return normCode(m[1].replace(/\s+/g, ''));
    const first = s.split(/[\s\-–]+/)[0];
    return first ? normCode(first) : '';
}

/**
 * @param {{ klasse?: string }} student
 * @param {{ code?: string, name?: string }} classRow
 * @param {{ classCode?: string, code?: string, displayName?: string }} [link]
 */
export function studentBelongsToClassRow(student, classRow, link) {
    const k = normStr(student && student.klasse);
    if (!k) return false;
    const nk = normCode(k);
    const kHead = classLabelHeadToken(k);
    const code = normCode(classRow && classRow.code);
    const name = normStr(classRow && classRow.name);
    const nameLc = name.toLowerCase();
    const kLc = k.toLowerCase();

    if (code && nk === code) return true;
    if (name && kLc === nameLc) return true;
    if (code && kHead && kHead === code) return true;

    const nameHead = classLabelHeadToken(name);
    if (nameHead && (nk === nameHead || (kHead && kHead === nameHead))) return true;

    const codeHead = classLabelHeadToken(code);
    if (codeHead && (nk === codeHead || (kHead && kHead === codeHead))) return true;

    const team = link && typeof link === 'object' ? link : null;
    if (team) {
        const tc = normCode(team.classCode || team.code);
        if (tc && (nk === tc || (kHead && kHead === tc))) return true;
        const dn = normStr(team.displayName);
        if (dn && kLc === dn.toLowerCase()) return true;
        const dnHead = classLabelHeadToken(dn);
        if (dnHead && (nk === dnHead || (kHead && kHead === dnHead))) return true;
    }
    return false;
}
