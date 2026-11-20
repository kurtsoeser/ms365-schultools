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
import { wireFlowSetupEmailPickers } from '../../shared/flow-setup-email-pickers.js';
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
    {
        inputId: 'frEmailDirektion',
        title: 'Direktion (Genehmiger)',
        onPicked: function (mail) {
            syncApproverInputs(mail);
            persistLfrApproverEmail(mail);
        }
    },
    { inputId: 'frEmailMailbox', title: 'Absender / Technik-Postfach' },
    { inputId: 'frFlowServiceAccount', title: 'Technik-Konto für Power Automate' },
    { inputId: 'lfrOutlookCalUser', title: 'Kalender-Besitzer (UPN)' },
    {
        inputId: 'lfrPermEmailDirektion',
        title: 'Direktion (Genehmiger)',
        onPicked: function (mail) {
            syncApproverInputs(mail);
            persistLfrApproverEmail(mail);
        }
    }
];

function wireFlowEmailPickers() {
    wireFlowSetupEmailPickers(FLOW_EMAIL_PICKERS, { wiredFlag: 'lfrPickerWired' });
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

function permsLog(msg) {
    const el = document.getElementById('frPermsLog') || document.getElementById('frLog');
    if (!el) return;
    const line = document.createElement('div');
    line.textContent = String(msg || '');
    el.appendChild(line);
    const panel = el.closest('details');
    if (panel) panel.open = true;
}

async function applySavedPermsToSharePoint() {
    const siteEl = document.getElementById('frSiteUrl');
    const nameEl = document.getElementById('frListName');
    const idEl = document.getElementById('frListId');
    const siteUrl = String((siteEl && siteEl.value) || '').trim();
    const listName = String((nameEl && nameEl.value) || '').trim() || 'Lehrer-Freistellungen';
    const listId = String((idEl && idEl.value) || '').trim();
    if (!siteUrl) {
        throw new Error('SharePoint-Website fehlt (Schritt 3).');
    }
    if (!window.ms365LfrListPerms || typeof window.ms365LfrListPerms.apply !== 'function') {
        throw new Error('Listenrechte-Modul nicht geladen.');
    }
    let flowAccount = '';
    try {
        if (window.ms365LfrSetupStatus && typeof window.ms365LfrSetupStatus.effectiveLfrFlowAccount === 'function') {
            const cfg = {
                flowServiceAccount: String(
                    (document.getElementById('frFlowServiceAccount') &&
                        document.getElementById('frFlowServiceAccount').value) ||
                        ''
                ),
                emailMailbox: String(
                    (document.getElementById('frEmailMailbox') && document.getElementById('frEmailMailbox').value) ||
                        ''
                )
            };
            flowAccount = window.ms365LfrSetupStatus.effectiveLfrFlowAccount(cfg);
        }
    } catch {
        /* ignore */
    }
    permsLog('Listenrechte für „' + listName + '“ …');
    await window.ms365LfrListPerms.apply(
        siteUrl,
        listName,
        { listId: listId, flowServiceAccount: flowAccount },
        permsLog
    );
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
        saveBtn.addEventListener('click', async function () {
            const next = saveLfrPermissions(readPermsFromForm());
            persistLfrApproverEmail(next.emailDirektion);
            fillPermsForm(next);
            if (window.ms365LfrSetup && typeof window.ms365LfrSetup.refreshGlance === 'function') {
                window.ms365LfrSetup.refreshGlance();
            }
            saveBtn.disabled = true;
            try {
                await applySavedPermsToSharePoint();
                if (typeof window.ms365ToastOrAlert === 'function') {
                    window.ms365ToastOrAlert(
                        'Rollen gespeichert und Listenrechte auf SharePoint gesetzt.'
                    );
                }
            } catch (e) {
                const msg = e && e.message ? e.message : String(e);
                permsLog('Fehler: ' + msg);
                if (typeof window.ms365ToastOrAlert === 'function') {
                    window.ms365ToastOrAlert(
                        'Rollen lokal gespeichert, Listenrechte fehlgeschlagen: ' + msg
                    );
                }
            } finally {
                saveBtn.disabled = false;
            }
        });
    }
}

window.ms365LfrInitSetupPermissions = initLfrSetupPermissions;
