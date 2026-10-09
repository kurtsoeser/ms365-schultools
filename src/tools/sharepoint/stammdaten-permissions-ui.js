/**
 * Berechtigungen für Stammdaten-Listen (Matrix + optional eingebettet).
 */
import {
    loadPermissionsConfig,
    savePermissionsConfig,
    prefillPermissionsFromSchularbeitenIfEmpty
} from './stammdaten-liste-permissions.js';
import { initStammdatenPermMatrixUi, readGrantRowsFromDom } from './stammdaten-perm-matrix-ui.js';

/** @deprecated Nur noch für eingebettete IDs – Matrix ersetzt die drei Picker. */
export function buildStammdatenGroupFields(prefix) {
    const p = String(prefix || 'spsm');
    return [
        {
            role: 'groupAdmin',
            labelInputId: p + 'GroupAdmin',
            idInputId: p + 'GroupAdminId',
            pickBtnId: p + 'GroupAdminPick',
            clearBtnId: p + 'GroupAdminClear',
            dialogTitle: 'Schulleitung (Admin-Gruppe)'
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
 * @param {string} [skipPermsId]
 * @param {string} [matrixBodyId]
 */
export function readPermissionsFromPickers(_fieldDefs, skipPermsId, matrixBodyId) {
    const grantRows = readGrantRowsFromDom(matrixBodyId || 'spsPermMatrixBody');
    const out = { grantRows, skipPerms: false };
    const skipEl = document.getElementById(skipPermsId || 'spsSkipPerms');
    if (skipEl && skipEl.checked) out.skipPerms = true;
    return out;
}

export function persistPickersToStorage(_fieldDefs, skipPermsId, matrixBodyId) {
    const patch = readPermissionsFromPickers(null, skipPermsId, matrixBodyId);
    savePermissionsConfig(patch);
    return patch;
}

export function initStammdatenPermissionsUi() {
    prefillPermissionsFromSchularbeitenIfEmpty();
    initStammdatenPermMatrixUi('spsPermMatrixBody');
}

/**
 * @param {string} prefix
 * @param {string} skipPermsId
 */
export function initEmbeddedPermissionsUi(prefix, skipPermsId) {
    prefillPermissionsFromSchularbeitenIfEmpty();
    const bodyId = String(prefix || 'tsdSpo') + 'PermMatrixBody';
    const tbody = document.getElementById(bodyId);
    if (tbody) {
        tbody.setAttribute('data-sps-add-btn', String(prefix || 'tsdSpo') + 'PermMatrixAddRow');
    }
    initStammdatenPermMatrixUi(bodyId);
    void skipPermsId;
}

export function htmlStammdatenPermMatrixBlock(prefix) {
    const p = String(prefix || 'sps');
    return (
        '<table class="tm-table sps-perm-matrix" style="width:100%;font-size:0.88em;margin:8px 0;" aria-label="Berechtigungen pro Entra-Gruppe und Liste">' +
        '<thead></thead>' +
        '<tbody id="' +
        p +
        'PermMatrixBody" data-sps-add-btn="' +
        p +
        'PermMatrixAddRow"></tbody>' +
        '</table>' +
        '<p class="hint" style="margin:0 0 8px;font-size:0.82em;">Pro Zeile eine Entra-Gruppe und die Rechte je Stammdaten-Liste (— = kein Zugriff). Gespeichert im Browser.</p>' +
        '<button type="button" class="btn btn-sm" id="' +
        p +
        'PermMatrixAddRow"><i class="bi bi-plus-lg"></i>Gruppe hinzufügen</button>'
    );
}
