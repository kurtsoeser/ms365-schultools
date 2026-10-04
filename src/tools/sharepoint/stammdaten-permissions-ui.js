/**
 * Entra-Gruppen-Picker für Stammdaten-Listen-Berechtigungen.
 */
import {
    wireEntraGroupPickerFields,
    readGroupPickerField,
    fillGroupPickerField
} from '../../shared/entra-group-picker.js';
import {
    loadPermissionsConfig,
    savePermissionsConfig,
    prefillPermissionsFromSchularbeitenIfEmpty
} from './stammdaten-liste-permissions.js';
import { normalizePermissionsConfig } from '../schularbeiten-planer/schularbeiten-planer-permissions.js';

/** @param {string} [prefix] z. B. spsm, tsdSpo, swSpo */
export function buildStammdatenGroupFields(prefix) {
    const p = String(prefix || 'spsm');
    return [
        {
            role: 'groupAdmin',
            labelInputId: p + 'GroupAdmin',
            idInputId: p + 'GroupAdminId',
            pickBtnId: p + 'GroupAdminPick',
            clearBtnId: p + 'GroupAdminClear',
            dialogTitle: 'Verwaltung / Admin'
        },
        {
            role: 'groupLehrer',
            labelInputId: p + 'GroupLehrer',
            idInputId: p + 'GroupLehrerId',
            pickBtnId: p + 'GroupLehrerPick',
            clearBtnId: p + 'GroupLehrerClear',
            dialogTitle: 'Lehrkräfte'
        },
        {
            role: 'groupSchueler',
            labelInputId: p + 'GroupSchueler',
            idInputId: p + 'GroupSchuelerId',
            pickBtnId: p + 'GroupSchuelerPick',
            clearBtnId: p + 'GroupSchuelerClear',
            dialogTitle: 'Schüler (Sammelgruppe)'
        }
    ];
}

export const STAMMDATEN_GROUP_FIELDS = buildStammdatenGroupFields('spsm');

/**
 * @param {ReturnType<typeof buildStammdatenGroupFields>} fieldDefs
 * @param {string} [skipPermsId]
 */
export function readPermissionsFromPickers(fieldDefs, skipPermsId) {
    const fields = fieldDefs || STAMMDATEN_GROUP_FIELDS;
    const out = { skipPerms: false };
    fields.forEach((f) => {
        const r = readGroupPickerField(f);
        out[f.role] = r.label;
        out[f.role + 'Id'] = r.id;
    });
    const skipEl = document.getElementById(skipPermsId || 'spsSkipPerms');
    if (skipEl && skipEl.checked) out.skipPerms = true;
    return out;
}

export function fillPermissionsPickers(cfg, fieldDefs) {
    const c = normalizePermissionsConfig(cfg);
    (fieldDefs || STAMMDATEN_GROUP_FIELDS).forEach((f) => {
        fillGroupPickerField(f, {
            id: c[f.role + 'Id'],
            label: c[f.role]
        });
    });
}

export function wireStammdatenPermissionPickers(fieldDefs, onChange) {
    const defs = fieldDefs || STAMMDATEN_GROUP_FIELDS;
    wireEntraGroupPickerFields({
        fields: defs.map((f) => ({
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
 * @param {ReturnType<typeof buildStammdatenGroupFields>} [fieldDefs]
 * @param {string} [skipPermsId]
 */
export function persistPickersToStorage(fieldDefs, skipPermsId) {
    const patch = readPermissionsFromPickers(fieldDefs, skipPermsId);
    savePermissionsConfig(patch);
    return patch;
}

export function initStammdatenPermissionsUi() {
    prefillPermissionsFromSchularbeitenIfEmpty();
    fillPermissionsPickers(loadPermissionsConfig());
    wireStammdatenPermissionPickers(STAMMDATEN_GROUP_FIELDS, () => persistPickersToStorage());
}

/**
 * @param {string} prefix
 * @param {string} skipPermsId
 */
export function initEmbeddedPermissionsUi(prefix, skipPermsId) {
    prefillPermissionsFromSchularbeitenIfEmpty();
    const defs = buildStammdatenGroupFields(prefix);
    fillPermissionsPickers(loadPermissionsConfig(), defs);
    wireStammdatenPermissionPickers(defs, () => persistPickersToStorage(defs, skipPermsId));
}
