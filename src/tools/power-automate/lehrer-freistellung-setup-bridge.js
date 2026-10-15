/**
 * ESM-Brücke: Schema + Setup-Status + Berechtigungen für lehrer-freistellung-setup.js (IIFE).
 */
import { LFR_COLUMNS } from '../lehrer-freistellung-planer/lfr-schema.js';
import { toGraphColumnBody } from '../freistellung-planer/freistellung-planer-schema.js';
import { computeLfrSetupGlance, effectiveLfrFlowAccount } from './lehrer-freistellung-setup-status.js';
import { applyLfrListPermissions } from '../lehrer-freistellung-planer/lfr-list-permissions.js';
import {
    loadLfrPermissions,
    normalizeLfrPermissions,
    saveLfrPermissions,
    persistLfrApproverEmail,
    resolveLfrApproverEmail,
    lfrPermsStepConfigured
} from '../lehrer-freistellung-planer/lfr-permissions.js';
import { pickEntraUser } from '../../shared/entra-user-picker.js';
import { wireEntraGroupPickerFields, readGroupPickerField, fillGroupPickerField } from '../../shared/entra-group-picker.js';
import { createPlannerExtraUsersController } from '../../shared/planner-extra-users-ui.js';

window.ms365LfrSchema = {
    columns: LFR_COLUMNS,
    toGraphColumnBody
};

window.ms365LfrSetupStatus = {
    computeSetupGlance: computeLfrSetupGlance,
    effectiveLfrFlowAccount,
    permsStepConfigured: () => lfrPermsStepConfigured()
};

window.ms365LfrListPerms = {
    apply: applyLfrListPermissions
};

const FLOW_EMAIL_PICKERS = [
    { inputId: 'frEmailDirektion', title: 'Direktion (Genehmiger)' },
    { inputId: 'frEmailMailbox', title: 'Absender / Technik-Postfach' },
    { inputId: 'frFlowServiceAccount', title: 'Technik-Konto für Power Automate' },
    { inputId: 'lfrOutlookCalUser', title: 'Kalender-Besitzer (UPN)' },
    { inputId: 'lfrPermEmailDirektion', title: 'Direktion (Genehmiger)' }
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
            const mail = String(u.mail).trim().toLowerCase();
            input.value = mail;
            input.dispatchEvent(new Event('change', { bubbles: true }));
            if (spec.inputId === 'frEmailDirektion' || spec.inputId === 'lfrPermEmailDirektion') {
                syncApproverInputs(mail);
                persistLfrApproverEmail(mail);
            }
        });
    });
}

function syncApproverInputs(mail) {
    const m = String(mail || '')
        .trim()
        .toLowerCase();
    ['frEmailDirektion', 'lfrPermEmailDirektion'].forEach(function (id) {
        const el = document.getElementById(id);
        if (el && el.value !== m) el.value = m;
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
        dialogTitle: 'Direktion / Verwaltung (Freigabe)'
    }
];

const EXTRA_USER_SPECS = [
    {
        configKey: 'lehrerUsers',
        listId: 'lfrPermLehrerUsers',
        addBtnId: 'lfrPermLehrerUserAdd',
        pickTitle: 'Lehrkraft einzeln',
        pickHint: 'Zusätzlich zur Lehrer-Entra-Gruppe – z. B. Vertretung ohne Gruppe.',
        emptyHint: 'Noch keine Einzelpersonen.'
    },
    {
        configKey: 'direktionUsers',
        listId: 'lfrPermDirektionUsers',
        addBtnId: 'lfrPermDirektionUserAdd',
        stammdatenBtnId: 'lfrPermDirektionStammdaten',
        pickTitle: 'Person für Direktion / Verwaltung',
        pickHint: 'Zusätzlich zur Direktion-Entra-Gruppe – Zugriff auf Freigabe und Übersicht.',
        emptyHint: 'Noch keine Einzelpersonen – z. B. aus Stammdaten übernehmen.'
    }
];

const extraUsersCtrl = createPlannerExtraUsersController(EXTRA_USER_SPECS);

function explicitSavedUserFlags() {
    try {
        const raw = JSON.parse(localStorage.getItem('ms365-lfr-perms-v1') || '{}');
        return {
            lehrerUsers: Array.isArray(raw.lehrerUsers),
            direktionUsers: Array.isArray(raw.direktionUsers)
        };
    } catch {
        return {};
    }
}

function readPermsFromForm() {
    const lehrer = readGroupPickerField({ labelInputId: 'lfrPermLehrer', idInputId: 'lfrPermLehrerId' });
    const dir = readGroupPickerField({ labelInputId: 'lfrPermDirektion', idInputId: 'lfrPermDirektionId' });
    const cur = loadLfrPermissions();
    const approverEl = document.getElementById('lfrPermEmailDirektion') || document.getElementById('frEmailDirektion');
    const emailDirektion = String((approverEl && approverEl.value) || cur.emailDirektion || '')
        .trim()
        .toLowerCase();
    return normalizeLfrPermissions({
        ...cur,
        groupLehrerId: lehrer.id,
        groupLehrerName: lehrer.label,
        groupDirektionId: dir.id,
        groupDirektionName: dir.label,
        ...extraUsersCtrl.readPatch(),
        emailDirektion
    });
}

function fillPermsForm(cfg) {
    const c = normalizeLfrPermissions(cfg);
    fillGroupPickerField(
        { labelInputId: 'lfrPermLehrer', idInputId: 'lfrPermLehrerId' },
        { label: c.groupLehrerName, id: c.groupLehrerId }
    );
    fillGroupPickerField(
        { labelInputId: 'lfrPermDirektion', idInputId: 'lfrPermDirektionId' },
        { label: c.groupDirektionName, id: c.groupDirektionId }
    );
    extraUsersCtrl.loadFromConfig(c, explicitSavedUserFlags());
    syncApproverInputs(resolveLfrApproverEmail(c));
}

function wireApproverMirror() {
    ['frEmailDirektion', 'lfrPermEmailDirektion'].forEach(function (id) {
        const el = document.getElementById(id);
        if (!el || el.dataset.lfrApproverMirror === '1') return;
        el.dataset.lfrApproverMirror = '1';
        el.addEventListener('change', function () {
            const mail = String(el.value || '')
                .trim()
                .toLowerCase();
            syncApproverInputs(mail);
        });
    });
}

export function initLfrSetupPermissions() {
    wireEntraGroupPickerFields({ fields: GROUP_FIELDS });
    wireFlowEmailPickers();
    wireApproverMirror();
    fillPermsForm(loadLfrPermissions());
    extraUsersCtrl.wire();
    const saveBtn = document.getElementById('lfrBtnSavePerms');
    if (saveBtn && !saveBtn.dataset.lfrWired) {
        saveBtn.dataset.lfrWired = '1';
        saveBtn.addEventListener('click', function () {
            const next = saveLfrPermissions(readPermsFromForm());
            persistLfrApproverEmail(next.emailDirektion);
            fillPermsForm(next);
            if (window.ms365LfrSetup && typeof window.ms365LfrSetup.refreshGlance === 'function') {
                window.ms365LfrSetup.refreshGlance();
            }
            if (typeof window.ms365ToastOrAlert === 'function') {
                window.ms365ToastOrAlert('Zugriff und Genehmiger für Lehrer-Freistellungen gespeichert.');
            }
        });
    }
}

window.ms365LfrInitSetupPermissions = initLfrSetupPermissions;
