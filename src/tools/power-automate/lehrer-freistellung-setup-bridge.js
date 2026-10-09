/**
 * ESM-Brücke: Schema + Setup-Status für lehrer-freistellung-setup.js (IIFE).
 */
import { LFR_COLUMNS } from '../lehrer-freistellung-planer/lfr-schema.js';
import { toGraphColumnBody } from '../freistellung-planer/freistellung-planer-schema.js';
import { computeLfrSetupGlance, effectiveLfrFlowAccount } from './lehrer-freistellung-setup-status.js';
import { applyLfrListPermissions } from '../lehrer-freistellung-planer/lfr-list-permissions.js';
import { loadLfrPermissions, normalizeLfrPermissions, saveLfrPermissions } from '../lehrer-freistellung-planer/lfr-permissions.js';
import { pickEntraUser } from '../../shared/entra-user-picker.js';
import { wireEntraGroupPickerFields, readGroupPickerField, fillGroupPickerField } from '../../shared/entra-group-picker.js';

window.ms365LfrSchema = {
    columns: LFR_COLUMNS,
    toGraphColumnBody
};

function permsStepConfigured() {
    const c = normalizeLfrPermissions(loadLfrPermissions());
    return !!(String(c.groupLehrerId || '').trim() || String(c.groupDirektionId || '').trim());
}

window.ms365LfrSetupStatus = {
    computeSetupGlance: computeLfrSetupGlance,
    effectiveLfrFlowAccount,
    permsStepConfigured
};

window.ms365LfrListPerms = {
    apply: applyLfrListPermissions
};

const FLOW_EMAIL_PICKERS = [
    { inputId: 'frEmailDirektion', title: 'Direktion (Genehmiger)' },
    { inputId: 'frEmailMailbox', title: 'Absender / Technik-Postfach' },
    { inputId: 'frFlowServiceAccount', title: 'Technik-Konto für Power Automate' },
    { inputId: 'lfrOutlookCalUser', title: 'Kalender-Besitzer (UPN)' }
];

function wireFlowEmailPickers() {
    FLOW_EMAIL_PICKERS.forEach(function (spec) {
        const input = document.getElementById(spec.inputId);
        const btn = document.getElementById(spec.inputId + 'Pick');
        if (!input || !btn || btn.dataset.lfrPickerWired === '1') return;
        btn.dataset.lfrPickerWired = '1';
        btn.addEventListener('click', async function () {
            const u = await pickEntraUser({ title: spec.title });
            if (!u || !u.mail) return;
            input.value = String(u.mail).trim().toLowerCase();
            input.dispatchEvent(new Event('change', { bubbles: true }));
        });
    });
}

const GROUP_FIELDS = [
    {
        labelInputId: 'lfrPermLehrer',
        idInputId: 'lfrPermLehrerId',
        pickBtnId: 'lfrPermLehrerPick',
        clearBtnId: 'lfrPermLehrerClear',
        dialogTitle: 'Lehrkräfte (Antrag stellen)'
    },
    {
        labelInputId: 'lfrPermDirektion',
        idInputId: 'lfrPermDirektionId',
        pickBtnId: 'lfrPermDirektionPick',
        clearBtnId: 'lfrPermDirektionClear',
        dialogTitle: 'Direktion (Freigabe)'
    }
];

function readPermsFromForm() {
    const lehrer = readGroupPickerField({ labelInputId: 'lfrPermLehrer', idInputId: 'lfrPermLehrerId' });
    const dir = readGroupPickerField({ labelInputId: 'lfrPermDirektion', idInputId: 'lfrPermDirektionId' });
    const cur = loadLfrPermissions();
    return normalizeLfrPermissions({
        ...cur,
        groupLehrerId: lehrer.id,
        groupLehrerName: lehrer.label,
        groupDirektionId: dir.id,
        groupDirektionName: dir.label,
        emailDirektion: String((document.getElementById('frEmailDirektion') || {}).value || cur.emailDirektion || '')
            .trim()
            .toLowerCase()
    });
}

function fillPermsForm(cfg) {
    fillGroupPickerField(
        { labelInputId: 'lfrPermLehrer', idInputId: 'lfrPermLehrerId' },
        { label: cfg.groupLehrerName, id: cfg.groupLehrerId }
    );
    fillGroupPickerField(
        { labelInputId: 'lfrPermDirektion', idInputId: 'lfrPermDirektionId' },
        { label: cfg.groupDirektionName, id: cfg.groupDirektionId }
    );
}

export function initLfrSetupPermissions() {
    wireEntraGroupPickerFields({ fields: GROUP_FIELDS });
    wireFlowEmailPickers();
    fillPermsForm(loadLfrPermissions());
    const saveBtn = document.getElementById('lfrBtnSavePerms');
    if (saveBtn && !saveBtn.dataset.lfrWired) {
        saveBtn.dataset.lfrWired = '1';
        saveBtn.addEventListener('click', function () {
            const next = saveLfrPermissions(readPermsFromForm());
            fillPermsForm(next);
            if (window.ms365LfrSetup && typeof window.ms365LfrSetup.refreshGlance === 'function') {
                window.ms365LfrSetup.refreshGlance();
            }
            if (typeof window.ms365ToastOrAlert === 'function') {
                window.ms365ToastOrAlert('Gruppen für Lehrer-Freistellungen gespeichert.');
            }
        });
    }
}

window.ms365LfrInitSetupPermissions = initLfrSetupPermissions;
