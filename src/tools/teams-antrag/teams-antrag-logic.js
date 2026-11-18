/**
 * Validierung, Filter, Naming-Preview für Teams-Anträge.
 */
import {
    ART,
    ART_LABEL,
    STATUS,
    STATUS_LABEL,
    normalizeAntrag,
    parseMailList
} from './teams-antrag-model.js';

function norm(s) {
    return String(s == null ? '' : s).trim();
}

export function slugify(raw) {
    return norm(raw)
        .replace(/[äÄ]/g, 'ae')
        .replace(/[öÖ]/g, 'oe')
        .replace(/[üÜ]/g, 'ue')
        .replace(/ß/g, 'ss')
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, '-')
        .replace(/-+/g, '-')
        .replace(/^-|-$/g, '')
        .slice(0, 40);
}

/**
 * @param {{ title?: string, namensvorschlag?: string, art?: string, schuljahr?: string }} input
 */
export function suggestMailNickname(input) {
    const base = slugify(input && (input.namensvorschlag || input.title)) || 'team';
    const art = norm(input && input.art) || 'sonstiges';
    const prefix =
        art === ART.AG
            ? 'ag'
            : art === ART.FACHGRUPPE
              ? 'fg'
              : art === ART.PROJEKT
                ? 'prj'
                : art === ART.KLASSE
                  ? 'kl'
                  : 'team';
    const yearBit = slugify((input && input.schuljahr) || '').replace(/\//g, '') || '';
    const nick = (prefix + (yearBit ? '-' + yearBit : '') + '-' + base)
        .replace(/[^a-z0-9-]/gi, '')
        .toLowerCase()
        .slice(0, 64);
    return nick;
}

/**
 * @param {{ title?: string, namensvorschlag?: string }} input
 */
export function suggestDisplayName(input) {
    return norm((input && (input.namensvorschlag || input.title)) || 'Neues Team').slice(0, 120);
}

/**
 * @param {object} draft
 */
export function validateAntrag(draft, opts) {
    const a = normalizeAntrag(draft);
    const errors = [];
    if (!a.title) errors.push('Titel fehlt.');
    if (!a.zweck || a.zweck.length < 8) errors.push('Zweck / Begründung bitte ausführlicher angeben.');
    if (a.art === ART.DIPLOM_VWA) {
        errors.push('Für Diplomarbeit/VWA bitte das Diplom-/VWA-Tool verwenden – kein generisches Team.');
    }
    const requireApplicant = !opts || opts.requireApplicant !== false;
    if (requireApplicant && !a.antragstellerEmail) {
        errors.push('Bitte mit dem schulischen Microsoft-Konto anmelden.');
    }
    return { ok: errors.length === 0, errors, antrag: a };
}

/**
 * @param {object[]} antraege
 * @param {{ q?: string, status?: string, mineEmail?: string, onlyMine?: boolean }} filters
 */
export function filterAntraege(antraege, filters) {
    const f = filters || {};
    const q = norm(f.q).toLowerCase();
    const status = norm(f.status).toLowerCase();
    const mine = norm(f.mineEmail).toLowerCase();
    const onlyMine = !!f.onlyMine;
    return (Array.isArray(antraege) ? antraege : [])
        .filter((a) => {
            if (status && a.status !== status) return false;
            if (onlyMine && mine && a.antragstellerEmail !== mine) return false;
            if (!q) return true;
            const blob = [
                a.title,
                a.zweck,
                a.namensvorschlag,
                a.finalDisplayName,
                a.antragstellerEmail,
                a.antragstellerName,
                ART_LABEL[a.art] || a.art,
                STATUS_LABEL[a.status] || a.status
            ]
                .join(' ')
                .toLowerCase();
            return blob.indexOf(q) !== -1;
        })
        .sort((x, y) => String(y.updatedAt || '').localeCompare(String(x.updatedAt || '')));
}

/**
 * Offene IT-Warteschlange.
 * @param {object[]} antraege
 */
export function openQueue(antraege) {
    const open = new Set([
        STATUS.BEANTRAGT,
        STATUS.IN_PRUEFUNG,
        STATUS.ZURUECKGESTELLT,
        STATUS.ANLAGE_FEHLGESCHLAGEN
    ]);
    return filterAntraege(antraege, {}).filter((a) => open.has(a.status));
}

/**
 * @param {string} email
 * @param {string[]} freigeberMails
 */
export function isFreigeber(email, freigeberMails) {
    const me = norm(email).toLowerCase();
    if (!me) return false;
    const list = (Array.isArray(freigeberMails) ? freigeberMails : [])
        .map((x) => norm(x).toLowerCase())
        .filter(Boolean);
    if (!list.length) return false;
    return list.indexOf(me) !== -1;
}

/**
 * @param {object} antrag
 */
export function buildProvisionPlan(antrag) {
    const a = normalizeAntrag(antrag);
    const displayName = a.finalDisplayName || suggestDisplayName(a);
    const mailNickname = a.finalMailNickname || suggestMailNickname(a);
    const owners = parseMailList(a.besitzerMails);
    if (a.antragstellerEmail && owners.indexOf(a.antragstellerEmail) === -1) {
        owners.unshift(a.antragstellerEmail);
    }
    const members = parseMailList(a.mitgliederMails);
    const issues = [];
    if (!displayName) issues.push('Anzeigename fehlt');
    if (!mailNickname || mailNickname.length < 3) issues.push('Alias ungültig');
    if (a.art === ART.DIPLOM_VWA) issues.push('Diplom/VWA nicht über Teams-Antrag anlegen');
    return {
        ok: issues.length === 0,
        issues,
        displayName,
        mailNickname,
        description:
            'Team aus MS365-Schul-Tools-Antrag ' +
            a.id +
            (a.zweck ? ' · ' + a.zweck.slice(0, 180) : ''),
        ownerEmails: owners,
        memberEmails: members.filter((m) => owners.indexOf(m) === -1)
    };
}

export { ART_LABEL, STATUS_LABEL, ART, STATUS };
