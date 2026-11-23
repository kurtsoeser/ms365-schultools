/**
 * Bulk: bestehende Rollenfelder (Name/E-Mail vermischt) + IT-Kontakte
 * in getrennte Name- und E-Mail-Spalten überführen.
 */
import { splitRolePersonFields } from './school-profile-logic.js';

const EMAIL_RE = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i;

/**
 * @param {unknown} raw
 * @returns {string}
 */
function firstEmailIn(raw) {
    const m = String(raw == null ? '' : raw).match(EMAIL_RE);
    return m ? String(m[0]).trim().toLowerCase() : '';
}

/**
 * @param {unknown} nameRaw
 * @param {unknown} emailRaw
 * @returns {{ name: string, email: string }}
 */
export function extractRolePerson(nameRaw, emailRaw) {
    let { name, email } = splitRolePersonFields(nameRaw, emailRaw);
    if (!email && name) {
        const found = firstEmailIn(name);
        if (found) {
            email = found;
            name = name
                .replace(EMAIL_RE, '')
                .replace(/^(direktion|leitung|schuladmin(?:istration)?|it(?:-?\/?-?m365)?(?:-?admin)?)\s*[:–-]?\s*/i, '')
                .replace(/^[\s:;,.–\-–—|/]+|[\s:;,.–\-–—|/]+$/g, '')
                .trim();
        }
    }
    return { name: name, email: email };
}

/**
 * @param {object|null|undefined} school
 * @returns {string}
 */
export function pickItContactEmail(school) {
    const s = school && typeof school === 'object' ? school : {};
    const list = Array.isArray(s.contactEmails) ? s.contactEmails : [];
    for (let i = 0; i < list.length; i++) {
        const e = firstEmailIn(list[i]);
        if (e) return e;
    }
    const fromContact = firstEmailIn(s.contactEmail);
    if (fromContact) return fromContact;
    const ex = s.extra && typeof s.extra === 'object' ? s.extra : {};
    return firstEmailIn(ex.ItContacts) || '';
}

/**
 * Berechnet Extra-Patch für eine Schule (nur geänderte Felder).
 * Überschreibt bestehende E-Mails nicht; füllt Leeres und trennt vermischte Werte.
 * @param {object|null|undefined} school
 * @returns {{ changed: boolean, extra: Record<string, string>, notes: string[] }}
 */
export function planRolePersonMigration(school) {
    const s = school && typeof school === 'object' ? school : {};
    const ex = s.extra && typeof s.extra === 'object' ? s.extra : {};
    /** @type {Record<string, string>} */
    const extra = {};
    /** @type {string[]} */
    const notes = [];

    function cur(key) {
        return String(ex[key] != null ? ex[key] : '').trim();
    }

    function applyRole(label, nameKey, emailKey) {
        const beforeName = cur(nameKey);
        const beforeEmail = cur(emailKey).toLowerCase();
        const person = extractRolePerson(beforeName, beforeEmail);
        let nextName = person.name;
        let nextEmail = person.email || beforeEmail;

        // IT: leere E-Mail aus Lizenz-IT-Kontakten füllen (ContactEmail / ItContacts)
        if (emailKey === 'ItM365AdminEmail' && !nextEmail) {
            const fromLic = pickItContactEmail(s);
            if (fromLic) {
                nextEmail = fromLic;
                notes.push(label + ': E-Mail aus IT-Kontakt(en)');
            }
        }

        if (nextName !== beforeName) {
            extra[nameKey] = nextName;
            notes.push(label + ': Name „' + beforeName + '“ → „' + nextName + '“');
        }
        if (nextEmail && nextEmail !== beforeEmail) {
            extra[emailKey] = nextEmail;
            notes.push(label + ': E-Mail → ' + nextEmail);
        }
    }

    applyRole('Direktion', 'Direktion', 'DirektionEmail');
    applyRole('Schuladministration', 'Schuladministration', 'SchuladministrationEmail');
    applyRole('IT', 'ItM365Admin', 'ItM365AdminEmail');

    return {
        changed: Object.keys(extra).length > 0,
        extra: extra,
        notes: notes
    };
}

/**
 * @param {object[]} schools
 * @returns {{ candidates: number, plans: { id: string, schoolName: string, extra: Record<string, string>, notes: string[] }[] }}
 */
export function planRolePersonMigrationForSchools(schools) {
    const list = Array.isArray(schools) ? schools : [];
    /** @type {{ id: string, schoolName: string, extra: Record<string, string>, notes: string[] }[]} */
    const plans = [];
    list.forEach(function (school) {
        if (!school || !school.id) return;
        const plan = planRolePersonMigration(school);
        if (!plan.changed) return;
        plans.push({
            id: String(school.id),
            schoolName: String(school.schoolName || school.id),
            extra: plan.extra,
            notes: plan.notes
        });
    });
    return { candidates: plans.length, plans: plans };
}
