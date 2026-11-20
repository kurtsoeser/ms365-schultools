/**
 * UI: Entra-Gruppen + Einzelpersonen für Schulaktivitäten-Listenrechte.
 * Verwaltung (Genehmigung) und Einreicher sind frei wählbar (Suche in Microsoft 365);
 * Stammdaten-Sammelgruppe wird nur als Vorschlag vorausgefüllt, wenn leer.
 */
import {
    wireEntraGroupPickerFields,
    readGroupPickerField,
    fillGroupPickerField,
    pickEntraGroup
} from '../../shared/entra-group-picker.js';
import {
    normalizeAktPermissionsConfig,
    saveAktPermissionsConfig,
    loadAktPermissionsConfig,
    loadEffectiveAktPermissionsConfig,
    AKT_PERMS_STORAGE_KEY
} from './schulaktivitaeten-planer-permissions.js';
import { createPlannerExtraUsersController } from '../../shared/planner-extra-users-ui.js';
import { loadSchoolAudienceGroups } from '../../shared/school-audience-groups.js';

/** Keine Rolle gesperrt – Gruppen immer über Entra-Picker wählbar. */
export const AKT_STAMMDATEN_GROUP_ROLES = new Set();

export const AKT_SETUP_GROUP_FIELDS = [
    {
        role: 'groupAdmin',
        labelInputId: 'spaktGroupAdmin',
        idInputId: 'spaktGroupAdminId',
        pickBtnId: 'spaktGroupAdminPick',
        clearBtnId: 'spaktGroupAdminClear',
        dialogTitle: 'Genehmigung / Direktion'
    },
    {
        role: 'groupLehrer',
        labelInputId: 'spaktGroupLehrer',
        idInputId: 'spaktGroupLehrerId',
        pickBtnId: 'spaktGroupLehrerPick',
        clearBtnId: 'spaktGroupLehrerClear',
        dialogTitle: 'Einreicher (Schulaktivitäten)'
    }
];

const SETUP_EXTRA_USER_SPECS = [
    {
        configKey: 'adminUsers',
        listId: 'spaktAdminUsers',
        addBtnId: 'spaktAdminUserAdd',
        stammdatenBtnId: 'spaktAdminStammdaten',
        pickTitle: 'Person für Genehmigung / Direktion',
        pickHint:
            'Direktor:in oder weitere Genehmiger:innen – Vollzugriff auf Anträge, Regelwerk und Stammdaten-Listen (zusätzlich zur Entra-Gruppe).',
        emptyHint: 'Noch keine Einzelpersonen – Gruppe wählen oder Person hinzufügen.'
    },
    {
        configKey: 'lehrerUsers',
        listId: 'spaktLehrerUsers',
        addBtnId: 'spaktLehrerUserAdd',
        pickTitle: 'Einreicher:in einzeln',
        pickHint: 'Zusätzlich zur Einreicher-Gruppe – Anträge schreiben, Klassen/Lehrer lesen.',
        emptyHint: 'Noch keine Einzelpersonen.'
    }
];

const setupExtraCtrl = createPlannerExtraUsersController(SETUP_EXTRA_USER_SPECS);
let setupExtraWired = false;

function explicitSavedUserFlags() {
    try {
        const raw = JSON.parse(localStorage.getItem(AKT_PERMS_STORAGE_KEY) || '{}');
        return {
            adminUsers: Array.isArray(raw.adminUsers),
            lehrerUsers: Array.isArray(raw.lehrerUsers)
        };
    } catch {
        return {};
    }
}

function escapeAttr(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/"/g, '&quot;')
        .replace(/</g, '&lt;');
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

function stammdatenEinreicherHintHtml() {
    const aud = loadSchoolAudienceGroups();
    if (!aud.groupLehrerId) {
        return (
            '<p class="fr-setup-perm-sub fr-setup-perm-sub--stamm">' +
            'Keine Lehrkräfte-Sammelgruppe in den <a href="' +
            escapeAttr(tenantStammdatenHref()) +
            '">Stammdaten</a> – hier eine Entra-Gruppe für Einreicher suchen oder Personen setzen.</p>'
        );
    }
    const name = escapeAttr(aud.groupLehrerName || aud.groupLehrerId);
    return (
        '<p class="fr-setup-perm-sub fr-setup-perm-sub--stamm">' +
        'Vorschlag aus Stammdaten: <strong>' +
        name +
        '</strong> – hier als Einreicher-Gruppe änderbar.</p>'
    );
}

function htmlEditableGroupPicker(f, vals) {
    const label = escapeAttr(vals[f.role] || '');
    const gid = escapeAttr(vals[f.role + 'Id'] || '');
    return `<label class="sr-only" for="${f.labelInputId}">Entra-Gruppe ${escapeAttr(f.dialogTitle)}</label>
          <div class="fr-setup-egp">
            <input id="${f.labelInputId}" type="text" readonly placeholder="Klicken oder tippen zum Suchen …" value="${label}" title="Öffnet die Microsoft-365-Gruppensuche">
            <input id="${f.idInputId}" type="hidden" value="${gid}">
            <button type="button" class="btn btn-sm" id="${f.pickBtnId}"><i class="bi bi-search"></i>Gruppe</button>
            <button type="button" class="btn btn-sm alt" id="${f.clearBtnId}" title="Gruppe löschen"><i class="bi bi-x-lg"></i></button>
          </div>`;
}

/**
 * @param {'admin'|'einreicher'} which
 * @param {ReturnType<typeof normalizeAktPermissionsConfig>} vals
 * @param {{ embedded?: boolean }} [opts]
 */
function htmlPermCard(which, vals, opts) {
    const embedded = !!(opts && opts.embedded);
    const cardClass = embedded
        ? 'fr-setup-perm-card fr-setup-perm-card--embedded'
        : 'fr-setup-perm-card';
    if (which === 'admin') {
        const f = AKT_SETUP_GROUP_FIELDS[0];
        const spec = SETUP_EXTRA_USER_SPECS[0];
        const head = embedded
            ? `<h4 class="fr-setup-perm-card__title">Genehmigung / Direktion</h4>`
            : `<div class="fr-setup-perm-card__head">
            <span class="fr-setup-perm-card__icon" aria-hidden="true"><i class="bi bi-building"></i></span>
            <div>
              <h4>Genehmigung / Direktion</h4>
              <p>Genehmigt Anträge, pflegt Regelwerk und Stammdaten – Vollzugriff auf alle Paket-Listen.</p>
            </div>
          </div>`;
        return `<article class="${cardClass}">
          ${head}
          <p class="fr-setup-perm-sub">Microsoft-365-Gruppe</p>
          ${htmlEditableGroupPicker(f, vals)}
          <p class="fr-setup-perm-sub">Einzelpersonen (optional)</p>
          <ul id="${spec.listId}" class="fr-setup-user-list" aria-label="Genehmigung Einzelpersonen"></ul>
          <div class="fr-setup-egp fr-setup-egp--tight">
            <button type="button" class="btn btn-sm akt-btn-secondary" id="${spec.addBtnId}"><i class="bi bi-person-plus"></i>Person</button>
            <button type="button" class="btn btn-sm akt-btn-ghost" id="${spec.stammdatenBtnId}" title="Verwaltung aus Stammdaten"><i class="bi bi-journal-text"></i>Aus Stammdaten</button>
          </div>
        </article>`;
    }
    const f = AKT_SETUP_GROUP_FIELDS[1];
    const spec = SETUP_EXTRA_USER_SPECS[1];
    const head = embedded
        ? `<h4 class="fr-setup-perm-card__title">Einreicher (Lehrkräfte)</h4>
          <p class="fr-setup-perm-sub">Ohne diese Gruppe sehen Lehrer nach „Listenrechte anwenden“ die Liste nicht.</p>`
        : `<div class="fr-setup-perm-card__head">
        <span class="fr-setup-perm-card__icon" aria-hidden="true"><i class="bi bi-pencil-square"></i></span>
        <div>
          <h4>Einreicher</h4>
          <p>Dürfen Schulaktivitäten einreichen, den Schulkalender (alle Termine) lesen und Klassen/Lehrerinnen lesen – typischerweise die Lehrkräfte-Gruppe.</p>
        </div>
      </div>`;
    return `<article class="${cardClass}">
      ${head}
      <p class="fr-setup-perm-sub">Microsoft-365-Gruppe</p>
      ${stammdatenEinreicherHintHtml()}
      ${htmlEditableGroupPicker(f, vals)}
      <p class="fr-setup-perm-sub">Einzelpersonen (optional)</p>
      <ul id="${spec.listId}" class="fr-setup-user-list" aria-label="Einreicher Einzelpersonen"></ul>
      <div class="fr-setup-egp fr-setup-egp--tight">
        <button type="button" class="btn btn-sm akt-btn-secondary" id="${spec.addBtnId}"><i class="bi bi-person-plus"></i>Person</button>
      </div>
    </article>`;
}

/**
 * @param {ReturnType<typeof normalizeAktPermissionsConfig>} [p]
 * @param {'admin'|'einreicher'|'all'} [which]
 * @param {{ embedded?: boolean }} [opts]
 */
export function htmlAktSetupEntraPermGrid(p, which, opts) {
    const vals = normalizeAktPermissionsConfig(p || loadEffectiveAktPermissionsConfig());
    const mode = which || 'all';
    const cardOpts = { embedded: !!(opts && opts.embedded) };
    const gridClass = cardOpts.embedded
        ? 'fr-setup-perm-grid fr-setup-perm-grid--single fr-setup-perm-grid--embedded'
        : 'fr-setup-perm-grid fr-setup-perm-grid--single';
    if (mode === 'admin') {
        return `<div class="${gridClass}">${htmlPermCard('admin', vals, cardOpts)}</div>`;
    }
    if (mode === 'einreicher') {
        return `<div class="${gridClass}">${htmlPermCard('einreicher', vals, cardOpts)}</div>`;
    }
    return `<div class="fr-setup-perm-grid fr-setup-perm-grid--two${cardOpts.embedded ? ' fr-setup-perm-grid--embedded' : ''}">
      ${htmlPermCard('admin', vals, cardOpts)}
      ${htmlPermCard('einreicher', vals, cardOpts)}
    </div>`;
}

/**
 * Liest nur Felder, die im DOM wirklich gemountet sind.
 * Sonst würde z. B. die Systemadmin-Karte „nur Genehmiger“ die Einreicher-Gruppe
 * auf leer überschreiben – danach fehlen Lehrkräften die Listenrechte.
 * @param {'admin'|'einreicher'|'all'} [which]
 */
export function readAktPermissionsFromPickers(which) {
    const mode = which || 'all';
    const out = { skipPerms: false };
    const fields =
        mode === 'admin'
            ? AKT_SETUP_GROUP_FIELDS.filter((f) => f.role === 'groupAdmin')
            : mode === 'einreicher'
              ? AKT_SETUP_GROUP_FIELDS.filter((f) => f.role === 'groupLehrer')
              : AKT_SETUP_GROUP_FIELDS;
    fields.forEach((f) => {
        // Fehlende Inputs nicht als „Gruppe löschen“ werten
        if (!document.getElementById(f.labelInputId) && !document.getElementById(f.idInputId)) {
            return;
        }
        const r = readGroupPickerField(f);
        out[f.role] = r.label;
        out[f.role + 'Id'] = r.id;
    });
    const usersPatch = setupExtraCtrl.readPatch();
    const userSpecs =
        mode === 'admin'
            ? SETUP_EXTRA_USER_SPECS.filter((s) => s.configKey === 'adminUsers')
            : mode === 'einreicher'
              ? SETUP_EXTRA_USER_SPECS.filter((s) => s.configKey === 'lehrerUsers')
              : SETUP_EXTRA_USER_SPECS;
    userSpecs.forEach((spec) => {
        // Liste nicht im DOM → Draft nicht als leeres Array speichern
        if (!document.getElementById(spec.listId)) return;
        if (Object.prototype.hasOwnProperty.call(usersPatch, spec.configKey)) {
            out[spec.configKey] = usersPatch[spec.configKey];
        }
    });
    return out;
}

export function fillAktPermissionsPickers(cfg) {
    const c = normalizeAktPermissionsConfig(cfg || loadEffectiveAktPermissionsConfig());
    AKT_SETUP_GROUP_FIELDS.forEach((f) => {
        fillGroupPickerField(f, {
            id: c[f.role + 'Id'],
            label: c[f.role]
        });
    });
    setupExtraCtrl.loadFromConfig(c, explicitSavedUserFlags());
}

/**
 * @param {() => void} [onChange]
 */
export function wireAktPermissionGroupPickers(onChange) {
    wireEntraGroupPickerFields({
        fields: AKT_SETUP_GROUP_FIELDS.map((f) => ({
            labelInputId: f.labelInputId,
            idInputId: f.idInputId,
            pickBtnId: f.pickBtnId,
            clearBtnId: f.clearBtnId,
            dialogTitle: f.dialogTitle
        })),
        onChange
    });
}

/**
 * @param {boolean|null|undefined} skipPerms
 * @param {'admin'|'einreicher'|'all'} [which] welcher UI-Ausschnitt gemountet ist
 */
export function persistAktPickersToStorage(skipPerms, which) {
    const patch = readAktPermissionsFromPickers(which || detectMountedAktPermWhich());
    if (skipPerms != null) patch.skipPerms = !!skipPerms;
    return saveAktPermissionsConfig(patch);
}

/** @returns {'admin'|'einreicher'|'all'} */
export function detectMountedAktPermWhich() {
    const hasAdmin = !!(
        document.getElementById('spaktGroupAdmin') || document.getElementById('spaktAdminUsers')
    );
    const hasLehrer = !!(
        document.getElementById('spaktGroupLehrer') || document.getElementById('spaktLehrerUsers')
    );
    if (hasAdmin && hasLehrer) return 'all';
    if (hasAdmin) return 'admin';
    if (hasLehrer) return 'einreicher';
    return 'all';
}

/**
 * @param {() => void} [onChange]
 */
export function initAktSetupExtraUsers(onChange) {
    // wire() ist idempotent pro Button (dataset.pepWired) – nach DOM-Remount erneut aufrufen
    setupExtraCtrl.wire(() => {
        if (typeof onChange === 'function') onChange();
    });
    setupExtraWired = true;
}

/** @param {() => void} [onChange] */
export function remountAktSetupPermPickers(onChange) {
    wireAktPermissionGroupPickers(onChange);
    initAktSetupExtraUsers(onChange);
    fillAktPermissionsPickers(loadEffectiveAktPermissionsConfig());
}

export { pickEntraGroup };
