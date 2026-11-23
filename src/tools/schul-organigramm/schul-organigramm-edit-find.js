/**
 * Find helpers + teacher select wiring for Schul-Organigramm edit.
 */
import { normStr, normEmail, escapeHtml } from './schul-organigramm-edit-utils.js';
import { ensureSchulOrganigrammSgaStructure } from './schul-organigramm-logic.js';

export function sgaScopeLabel(scope) {
    const s = normStr(scope).toLowerCase();
    if (s === 'teacher') return 'Lehrer';
    if (s === 'student') return 'Schüler';
    if (s === 'external') return 'Extern';
    return s || '–';
}

export function normalizeSgaScope(raw) {
    const s = normStr(raw).toLowerCase();
    if (s === 'lehrer' || s === 'teacher') return 'teacher';
    if (s === 'schueler' || s === 'schüler' || s === 'student') return 'student';
    if (s === 'extern' || s === 'external' || s === 'eltern') return 'external';
    return '';
}

/**
 * @param {HTMLSelectElement} selectEl
 * @param {object[]} teachers
 */
export function fillTeacherSelect(selectEl, teachers) {
    (teachers || [])
        .filter(function (t) {
            return normEmail(t && t.email).indexOf('@') !== -1;
        })
        .sort(function (a, b) {
            return normStr(a.name).localeCompare(normStr(b.name), 'de');
        })
        .forEach(function (t) {
            const opt = document.createElement('option');
            opt.value = normEmail(t.email);
            opt.textContent =
                (normStr(t.name) || normEmail(t.email)) +
                (t.code ? ' (' + normStr(t.code) + ')' : '') +
                ' – ' +
                normEmail(t.email);
            selectEl.appendChild(opt);
        });
}

/**
 * Lehrerliste ↔ Name/E-Mail-Felder für Leitung/KV verdrahten.
 * @param {{ teacherEl: HTMLSelectElement, headNameEl: HTMLInputElement, headEmailEl: HTMLInputElement, hintEl: Element|null, teachers: object[], vacantText: string }} opts
 */
export function wireTeacherHeadFields(opts) {
    const teacherEl = opts.teacherEl;
    const headNameEl = opts.headNameEl;
    const headEmailEl = opts.headEmailEl;
    const hintEl = opts.hintEl;
    const teachers = opts.teachers || [];
    const vacantText = opts.vacantText || 'Keine Person gesetzt.';

    function syncTeacherSelect() {
        const em = normEmail(headEmailEl.value);
        teacherEl.value = em && teachers.some(function (t) {
            return normEmail(t.email) === em;
        })
            ? em
            : '';
    }

    function updateHeadHint() {
        if (!hintEl) return;
        const em = normEmail(headEmailEl.value);
        if (!em) {
            hintEl.textContent = vacantText;
            return;
        }
        const t = findTeacherByEmail({ teachers: teachers }, em);
        if (t) {
            hintEl.textContent =
                'Verknüpft mit Lehrerliste: ' +
                (t.name || t.code || em) +
                (t.code ? ' · ' + t.code : '');
        } else {
            hintEl.textContent =
                'E-Mail gesetzt, aber kein Treffer in der Lehrerliste (trotzdem gültig).';
        }
    }

    teacherEl.addEventListener('change', function () {
        const em = normEmail(teacherEl.value);
        if (!em) return;
        const t = findTeacherByEmail({ teachers: teachers }, em);
        headEmailEl.value = em;
        if (t && t.name) headNameEl.value = t.name;
        updateHeadHint();
    });
    headEmailEl.addEventListener('input', function () {
        syncTeacherSelect();
        updateHeadHint();
    });
    headNameEl.addEventListener('input', updateHeadHint);
    syncTeacherSelect();
    updateHeadHint();
    return { syncTeacherSelect: syncTeacherSelect, updateHeadHint: updateHeadHint };
}

/**
 * @param {object} settings
 * @param {string} refKey
 */
export function findArge(settings, refKey) {
    const key = normStr(refKey).toUpperCase();
    const keyL = normStr(refKey).toLowerCase();
    const arges = Array.isArray(settings.arges) ? settings.arges : [];
    return (
        arges.find(function (a) {
            return normStr(a.code).toUpperCase() === key || normStr(a.name).toLowerCase() === keyL;
        }) || null
    );
}

/**
 * @param {object} settings
 * @param {string} email
 */
export function findTeacherByEmail(settings, email) {
    const em = normEmail(email);
    if (!em || em.indexOf('@') === -1) return null;
    const teachers = Array.isArray(settings.teachers) ? settings.teachers : [];
    return (
        teachers.find(function (t) {
            return normEmail(t && t.email) === em;
        }) || null
    );
}

/**
 * @param {object} settings
 * @param {string} refKey – Klassen-Code oder Name
 */
export function findClass(settings, refKey) {
    const key = normStr(refKey).toUpperCase();
    const keyL = normStr(refKey).toLowerCase();
    const classes = Array.isArray(settings.classes) ? settings.classes : [];
    return (
        classes.find(function (c) {
            return normStr(c.code).toUpperCase() === key || normStr(c.name).toLowerCase() === keyL;
        }) || null
    );
}

/**
 * @param {object} settings
 * @param {string} refKey – E-Mail oder Name
 */
export function findStudentCouncil(settings, refKey) {
    const key = normStr(refKey);
    const keyL = key.toLowerCase();
    const list = Array.isArray(settings.studentCouncil) ? settings.studentCouncil : [];
    if (keyL.indexOf('@') !== -1) {
        const byMail = list.find(function (s) {
            return normEmail(s.email) === keyL;
        });
        if (byMail) return byMail;
    }
    return (
        list.find(function (s) {
            return normStr(s.name).toLowerCase() === keyL;
        }) || null
    );
}

/**
 * @param {object} settings
 * @param {string} refKey – E-Mail oder Name
 */
export function findSga(settings, refKey) {
    const key = normStr(refKey);
    const keyL = key.toLowerCase();
    const list = Array.isArray(settings.sga) ? settings.sga : [];
    if (keyL.indexOf('@') !== -1) {
        const byMail = list.find(function (s) {
            return normEmail(s.email) === keyL;
        });
        if (byMail) return byMail;
    }
    return (
        list.find(function (s) {
            return normStr(s.name).toLowerCase() === keyL;
        }) || null
    );
}
