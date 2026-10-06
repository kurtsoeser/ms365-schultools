/**
 * UI-Hilfen: Entra-Gruppen-Picker für Schularbeiten-Berechtigungen (Setup + Planer).
 */
import {
    wireEntraGroupPickerFields,
    readGroupPickerField,
    fillGroupPickerField,
    pickEntraGroup
} from '../../shared/entra-group-picker.js';
import {
    normalizePermissionsConfig,
    savePermissionsConfig,
    loadPermissionsConfig,
    PERMS_STORAGE_KEY
} from './schularbeiten-planer-permissions.js';
import { createPlannerExtraUsersController } from '../../shared/planner-extra-users-ui.js';
import {
    SA_STAMMDATEN_GROUP_ROLES,
    htmlReadonlyStammdatenAudienceEgp,
    fillReadonlyStammdatenGroupField,
    stripStammdatenGroupFieldsFromPatch,
    filterEditableGroupFields,
    saRoleToAudienceKind
} from '../../shared/planner-stammdaten-audience-ui.js';

/** Setup-Tool sharepoint-liste-schularbeiten.html */
export const SETUP_GROUP_FIELDS = [
    {
        role: 'groupAdmin',
        labelInputId: 'spsaGroupAdmin',
        idInputId: 'spsaGroupAdminId',
        pickBtnId: 'spsaGroupAdminPick',
        clearBtnId: 'spsaGroupAdminClear',
        dialogTitle: 'Verwaltung / Admin'
    },
    {
        role: 'groupLehrer',
        labelInputId: 'spsaGroupLehrer',
        idInputId: 'spsaGroupLehrerId',
        pickBtnId: 'spsaGroupLehrerPick',
        clearBtnId: 'spsaGroupLehrerClear',
        dialogTitle: 'Lehrkräfte'
    },
    {
        role: 'groupSchueler',
        labelInputId: 'spsaGroupSchueler',
        idInputId: 'spsaGroupSchuelerId',
        pickBtnId: 'spsaGroupSchuelerPick',
        clearBtnId: 'spsaGroupSchuelerClear',
        dialogTitle: 'Schüler'
    }
];

/** Planer Administration */
export const PLANER_GROUP_FIELDS = [
    {
        role: 'groupAdmin',
        labelInputId: 'saPermAdmin',
        idInputId: 'saPermAdminId',
        pickBtnId: 'saPermAdminPick',
        clearBtnId: 'saPermAdminClear',
        dialogTitle: 'Verwaltung / Admin'
    },
    {
        role: 'groupLehrer',
        labelInputId: 'saPermLehrer',
        idInputId: 'saPermLehrerId',
        pickBtnId: 'saPermLehrerPick',
        clearBtnId: 'saPermLehrerClear',
        dialogTitle: 'Lehrkräfte'
    },
    {
        role: 'groupSchueler',
        labelInputId: 'saPermSchueler',
        idInputId: 'saPermSchuelerId',
        pickBtnId: 'saPermSchuelerPick',
        clearBtnId: 'saPermSchuelerClear',
        dialogTitle: 'Schüler'
    }
];

const SETUP_EXTRA_USER_SPECS = [
    {
        configKey: 'adminUsers',
        listId: 'spsaAdminUsers',
        addBtnId: 'spsaAdminUserAdd',
        stammdatenBtnId: 'spsaAdminStammdaten',
        pickTitle: 'Person für Verwaltung / Admin',
        pickHint: 'Sekretariat, Schulleitung oder IT – Vollzugriff auf alle Planer-Listen.',
        emptyHint: 'Noch keine Einzelpersonen – z. B. aus Stammdaten übernehmen.'
    },
    {
        configKey: 'lehrerUsers',
        listId: 'spsaLehrerUsers',
        addBtnId: 'spsaLehrerUserAdd',
        pickTitle: 'Lehrkraft einzeln',
        pickHint: 'Zusätzlich zur Lehrer-Entra-Gruppe – z. B. Vertretung ohne Gruppe.',
        emptyHint: 'Noch keine Einzelpersonen.'
    },
    {
        configKey: 'schuelerUsers',
        listId: 'spsaSchuelerUsers',
        addBtnId: 'spsaSchuelerUserAdd',
        pickTitle: 'Schüler/in einzeln',
        pickHint: 'Zusätzlich zur Schüler-Entra-Gruppe – Lesen auf der Schularbeiten-Liste.',
        emptyHint: 'Noch keine Einzelpersonen.'
    }
];

const PLANER_EXTRA_USER_SPECS = [
    {
        configKey: 'adminUsers',
        listId: 'saPermAdminUsers',
        addBtnId: 'saPermAdminUserAdd',
        stammdatenBtnId: 'saPermAdminStammdaten',
        pickTitle: 'Person für Verwaltung / Admin',
        pickHint: 'Sekretariat, Schulleitung oder IT – Vollzugriff auf alle Planer-Listen.',
        emptyHint: 'Noch keine Einzelpersonen – z. B. aus Stammdaten übernehmen.'
    },
    {
        configKey: 'lehrerUsers',
        listId: 'saPermLehrerUsers',
        addBtnId: 'saPermLehrerUserAdd',
        pickTitle: 'Lehrkraft einzeln',
        pickHint: 'Zusätzlich zur Lehrer-Entra-Gruppe.',
        emptyHint: 'Noch keine Einzelpersonen.'
    },
    {
        configKey: 'schuelerUsers',
        listId: 'saPermSchuelerUsers',
        addBtnId: 'saPermSchuelerUserAdd',
        pickTitle: 'Schüler/in einzeln',
        pickHint: 'Zusätzlich zur Schüler-Entra-Gruppe.',
        emptyHint: 'Noch keine Einzelpersonen.'
    }
];

const setupExtraCtrl = createPlannerExtraUsersController(SETUP_EXTRA_USER_SPECS);
const planerExtraCtrl = createPlannerExtraUsersController(PLANER_EXTRA_USER_SPECS);

let setupExtraWired = false;

function extraCtrlForFieldDefs(fieldDefs) {
    const first = fieldDefs && fieldDefs[0];
    if (!first) return null;
    if (first.labelInputId && String(first.labelInputId).startsWith('spsa')) return setupExtraCtrl;
    return planerExtraCtrl;
}

function explicitSavedUserFlags() {
    try {
        const raw = JSON.parse(localStorage.getItem(PERMS_STORAGE_KEY) || '{}');
        return {
            adminUsers: Array.isArray(raw.adminUsers),
            lehrerUsers: Array.isArray(raw.lehrerUsers),
            schuelerUsers: Array.isArray(raw.schuelerUsers)
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

/**
 * @param {ReturnType<typeof normalizePermissionsConfig>} [p]
 * @param {{ delegated?: boolean, mode?: 'setup'|'planer' }} [opts]
 */
export function htmlSchularbeitenEntraPermGrid(p, opts) {
    const vals = normalizePermissionsConfig(p || loadPermissionsConfig());
    const delegated = !!(opts && opts.delegated);
    const mode = opts && opts.mode === 'setup' ? 'setup' : delegated ? 'planer' : 'setup';
    const specs = mode === 'setup' ? SETUP_EXTRA_USER_SPECS : PLANER_EXTRA_USER_SPECS;
    const fields = mode === 'setup' ? SETUP_GROUP_FIELDS : PLANER_GROUP_FIELDS;
    const [adminEgp, lehrerEgp, schuelerEgp] = [0, 1, 2].map((i) => {
        const f = fields[i];
        const label = escapeAttr(vals[f.role] || '');
        const gid = escapeAttr(vals[f.role + 'Id'] || '');
        if (SA_STAMMDATEN_GROUP_ROLES.has(f.role)) {
            const kind = saRoleToAudienceKind(f.role);
            return htmlReadonlyStammdatenAudienceEgp({
                labelInputId: f.labelInputId,
                idInputId: f.idInputId,
                kind: kind === 'lehrer' ? 'lehrer' : 'schueler'
            });
        }
        if (delegated) {
            return `<div class="fr-setup-egp">
            <input id="${f.labelInputId}" type="text" readonly value="${label}" placeholder="Gruppe wählen …">
            <input id="${f.idInputId}" type="hidden" value="${gid}">
            <button type="button" class="btn btn-sm" data-egp-pick="${f.role}"><i class="bi bi-search"></i>Gruppe</button>
            <button type="button" class="btn btn-sm alt" data-egp-clear="${f.role}" title="Gruppe löschen"><i class="bi bi-x-lg"></i></button>
          </div>`;
        }
        return `<label class="sr-only" for="${f.labelInputId}">Entra-Gruppe ${escapeAttr(f.dialogTitle)}</label>
          <div class="fr-setup-egp">
            <input id="${f.labelInputId}" type="text" readonly placeholder="Gruppe wählen …">
            <input id="${f.idInputId}" type="hidden">
            <button type="button" class="btn btn-sm" id="${f.pickBtnId}"><i class="bi bi-search"></i>Gruppe</button>
            <button type="button" class="btn btn-sm alt" id="${f.clearBtnId}" title="Gruppe löschen"><i class="bi bi-x-lg"></i></button>
          </div>`;
    });

    const card = (icon, title, desc, egp, spec, showStammdaten) =>
        `<article class="fr-setup-perm-card">
          <div class="fr-setup-perm-card__head">
            <span class="fr-setup-perm-card__icon" aria-hidden="true"><i class="bi bi-${icon}"></i></span>
            <div><h4>${title}</h4><p>${desc}</p></div>
          </div>
          <p class="fr-setup-perm-sub">Entra-Gruppe (optional)</p>
          ${egp}
          <p class="fr-setup-perm-sub">Einzelpersonen</p>
          <ul id="${spec.listId}" class="fr-setup-user-list" aria-label="${escapeAttr(title)} Einzelpersonen"></ul>
          <div class="fr-setup-egp fr-setup-egp--tight">
            <button type="button" class="btn btn-sm" id="${spec.addBtnId}"><i class="bi bi-person-plus"></i>Person</button>
            ${
                showStammdaten && spec.stammdatenBtnId
                    ? `<button type="button" class="btn btn-sm alt" id="${spec.stammdatenBtnId}" title="Verwaltung aus Stammdaten"><i class="bi bi-journal-text"></i>Aus Stammdaten</button>`
                    : ''
            }
          </div>
        </article>`;

    return `<div class="fr-setup-perm-grid">
      ${card(
          'building',
          'Verwaltung / Admin',
          'Vollzugriff auf alle Planer-Listen – Entra-Gruppe und optional Einzelpersonen (Sekretariat, IT).',
          adminEgp,
          specs[0],
          true
      )}
      ${card(
          'person-badge',
          'Lehrkräfte',
          'Schularbeiten: Beitragen; Regelwerk &amp; Terminfenster: Lesen – Sammelgruppe aus Stammdaten; optional Einzelpersonen.',
          lehrerEgp,
          specs[1],
          false
      )}
      ${card(
          'mortarboard',
          'Schüler',
          'Nur Schularbeiten-Liste: Lesen – Schüler-Sammelgruppe aus Stammdaten; optional Einzelpersonen.',
          schuelerEgp,
          specs[2],
          false
      )}
    </div>`;
}

/**
 * @param {typeof SETUP_GROUP_FIELDS} fieldDefs
 */
export function readPermissionsFromPickers(fieldDefs) {
    const out = { skipPerms: false };
    filterEditableGroupFields(fieldDefs, SA_STAMMDATEN_GROUP_ROLES).forEach((f) => {
        const r = readGroupPickerField(f);
        out[f.role] = r.label;
        out[f.role + 'Id'] = r.id;
    });
    const ctrl = extraCtrlForFieldDefs(fieldDefs);
    if (ctrl) Object.assign(out, ctrl.readPatch());
    return out;
}

/**
 * @param {object} cfg
 * @param {typeof SETUP_GROUP_FIELDS} fieldDefs
 */
export function fillPermissionsPickers(cfg, fieldDefs) {
    const c = normalizePermissionsConfig(cfg);
    fieldDefs.forEach((f) => {
        if (SA_STAMMDATEN_GROUP_ROLES.has(f.role)) {
            fillReadonlyStammdatenGroupField(f);
            return;
        }
        fillGroupPickerField(f, {
            id: c[f.role + 'Id'],
            label: c[f.role]
        });
    });
    const ctrl = extraCtrlForFieldDefs(fieldDefs);
    if (ctrl) ctrl.loadFromConfig(c, explicitSavedUserFlags());
}

/**
 * @param {typeof SETUP_GROUP_FIELDS} fieldDefs
 * @param {() => void} [onChange]
 */
export function wirePermissionGroupPickers(fieldDefs, onChange) {
    wireEntraGroupPickerFields({
        fields: filterEditableGroupFields(fieldDefs, SA_STAMMDATEN_GROUP_ROLES).map((f) => ({
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
 * @param {typeof SETUP_GROUP_FIELDS} fieldDefs
 * @param {HTMLElement} root
 */
export function wirePermissionGroupPickersDelegated(root, fieldDefs, onChange) {
    if (!root || root.dataset.saEgpDelegated === '1') return;
    root.dataset.saEgpDelegated = '1';

    root.addEventListener('click', (ev) => {
        const pick = ev.target.closest('[data-egp-pick]');
        if (pick) {
            const id = pick.getAttribute('data-egp-pick');
            if (SA_STAMMDATEN_GROUP_ROLES.has(id)) return;
            const def = fieldDefs.find((f) => f.role === id);
            if (!def) return;
            pickEntraGroup({ title: def.dialogTitle })
                .then((sel) => {
                    if (!sel) return;
                    fillGroupPickerField(def, { id: sel.id, label: sel.label || sel.displayName });
                    if (typeof onChange === 'function') onChange();
                })
                .catch((e) => {
                    const msg = e && e.message ? e.message : String(e);
                    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
                    else window.alert(msg);
                });
            return;
        }
        const clear = ev.target.closest('[data-egp-clear]');
        if (clear) {
            const role = clear.getAttribute('data-egp-clear');
            if (SA_STAMMDATEN_GROUP_ROLES.has(role)) return;
            const def = fieldDefs.find((f) => f.role === role);
            if (!def) return;
            fillGroupPickerField(def, { id: '', label: '' });
            if (typeof onChange === 'function') onChange();
        }
    });
}

export function persistPickersToStorage(fieldDefs, skipPerms) {
    const patch = stripStammdatenGroupFieldsFromPatch(
        readPermissionsFromPickers(fieldDefs),
        SA_STAMMDATEN_GROUP_ROLES
    );
    if (skipPerms != null) patch.skipPerms = !!skipPerms;
    return savePermissionsConfig(patch);
}

/**
 * @param {() => void} [onChange]
 */
export function initSchularbeitenSetupExtraUsers(onChange) {
    if (!setupExtraWired) {
        setupExtraCtrl.wire(() => {
            if (typeof onChange === 'function') onChange();
        });
        setupExtraWired = true;
    }
    SETUP_GROUP_FIELDS.forEach((f) => {
        if (SA_STAMMDATEN_GROUP_ROLES.has(f.role)) fillReadonlyStammdatenGroupField(f);
    });
    setupExtraCtrl.loadFromConfig(loadPermissionsConfig(), explicitSavedUserFlags());
}

/**
 * Nach jedem Neuzeichnen der Planer-Administration.
 * @param {() => void} [onChange]
 */
export function refreshSchularbeitenPlanerExtraUsers(onChange) {
    if (!document.getElementById('saPermAdminUsers')) return;
    PLANER_GROUP_FIELDS.forEach((f) => {
        if (SA_STAMMDATEN_GROUP_ROLES.has(f.role)) fillReadonlyStammdatenGroupField(f);
    });
    planerExtraCtrl.wire(() => {
        if (typeof onChange === 'function') onChange();
    });
    planerExtraCtrl.loadFromConfig(loadPermissionsConfig(), explicitSavedUserFlags());
}
