/**
 * Anzeige: Lehrer-/Schüler-Sammelgruppe kommt aus Stammdaten (nicht hier wählen).
 */
import { loadSchoolAudienceGroups } from './school-audience-groups.js';

export const SA_STAMMDATEN_GROUP_ROLES = new Set(['groupLehrer', 'groupSchueler']);
/** Freistellung: Schüler-Gruppe im Setup wählbar (zusätzlich zur Stammdaten-Sammelgruppe). */
export const FR_STAMMDATEN_GROUP_ROLES = new Set();

function escapeAttr(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/"/g, '&quot;')
        .replace(/</g, '&lt;');
}

function audienceForKind(kind) {
    const aud = loadSchoolAudienceGroups();
    if (kind === 'lehrer') {
        return { id: aud.groupLehrerId, label: aud.groupLehrerName };
    }
    return { id: aud.groupSchuelerId, label: aud.groupSchuelerName };
}

export function saRoleToAudienceKind(role) {
    if (role === 'groupLehrer') return 'lehrer';
    if (role === 'groupSchueler') return 'schueler';
    return '';
}

function tenantStammdatenHref() {
    try {
        const p = String(typeof location !== 'undefined' ? location.pathname : '')
            .replace(/\\/g, '/')
            .toLowerCase();
        if (/\/tools\/[^/]+\//.test(p)) return '../../tenant.html';
        if (/\/tools\//.test(p)) return '../tenant.html';
    } catch {
        /* ignore */
    }
    return 'tenant.html';
}

export function stammdatenAudienceHintLine() {
    return (
        '<p class="fr-setup-perm-sub fr-setup-perm-sub--stamm">' +
        'Standard-Sammelgruppe aus den <a href="' +
        escapeAttr(tenantStammdatenHref()) +
        '">Stammdaten</a> ' +
        '(Einrichtung / MS&nbsp;365-Gruppenverwaltung) – hier nicht änderbar.</p>'
    );
}

/**
 * @param {{ labelInputId: string, idInputId: string, kind: 'lehrer'|'schueler' }} opts
 */
export function htmlReadonlyStammdatenAudienceEgp(opts) {
    const { id, label } = audienceForKind(opts.kind);
    const placeholder = id ? '' : 'In Stammdaten noch nicht verknüpft …';
    return (
        stammdatenAudienceHintLine() +
        `<div class="fr-setup-egp fr-setup-egp--readonly-stamm">
            <input id="${escapeAttr(opts.labelInputId)}" type="text" readonly value="${escapeAttr(label || (id ? id : ''))}" placeholder="${escapeAttr(placeholder)}">
            <input id="${escapeAttr(opts.idInputId)}" type="hidden" value="${escapeAttr(id || '')}">
        </div>`
    );
}

/**
 * @param {{ labelInputId: string, idInputId: string, role: string }} fieldDef
 */
export function fillReadonlyStammdatenGroupField(fieldDef) {
    const kind = saRoleToAudienceKind(fieldDef.role);
    if (!kind) return;
    const { id, label } = audienceForKind(kind);
    const labelEl = document.getElementById(fieldDef.labelInputId);
    const idEl = document.getElementById(fieldDef.idInputId);
    if (labelEl) {
        labelEl.value = label || (id ? id : '');
        labelEl.placeholder = id ? '' : 'In Stammdaten noch nicht verknüpft …';
    }
    if (idEl) idEl.value = id || '';
}

/**
 * @param {Record<string, unknown>} patch
 * @param {Set<string>} stammdatenRoles
 */
export function stripStammdatenGroupFieldsFromPatch(patch, stammdatenRoles) {
    const out = Object.assign({}, patch);
    stammdatenRoles.forEach((role) => {
        delete out[role];
        delete out[role + 'Id'];
    });
    return out;
}

/**
 * @param {Array<{ role: string }>} fieldDefs
 * @param {Set<string>} stammdatenRoles
 */
export function filterEditableGroupFields(fieldDefs, stammdatenRoles) {
    return (fieldDefs || []).filter((f) => !stammdatenRoles.has(f.role));
}
