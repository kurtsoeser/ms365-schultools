/**
 * ESM-Brücke: Microsoft-Suche (Konten) + Entra-Rollen (Schritt 4) für Schulaktivitäten-Flow-Setup.
 */
import { wireFlowSetupEmailPickers } from '../../shared/flow-setup-email-pickers.js';
import {
    htmlAktSetupEntraPermGrid,
    remountAktSetupPermPickers,
    persistAktPickersToStorage
} from '../schulaktivitaeten-planer/schulaktivitaeten-permissions-ui.js';
import {
    loadEffectiveAktPermissionsConfig,
    aktEntraGroupsConfigured
} from '../schulaktivitaeten-planer/schulaktivitaeten-planer-permissions.js';

const FLOW_EMAIL_PICKERS = [
    {
        inputId: 'aktFlowServiceAccount',
        title: 'Technik-Konto für Power Automate',
        hint: 'Dienstkonto, mit dem der Flow importiert und die Connections verbunden werden.'
    },
    {
        inputId: 'aktEmailDirektion',
        title: 'Direktion (Genehmiger für Approvals)',
        hint: 'Person oder Postfach als einziger Approvals-Genehmiger im Flow.'
    },
    {
        inputId: 'aktEmailMailbox',
        title: 'Absender für Status-Mails',
        hint: 'Technik-Konto oder freigegebenes Postfach als Absender der Status-Mails.'
    }
];

function mountRolesStep() {
    const cfg = loadEffectiveAktPermissionsConfig();
    const admin = document.getElementById('aktFlowEntraPermAdmin');
    const ein = document.getElementById('aktFlowEntraPermEinreicher');
    if (admin && admin.dataset.aktPermMounted !== '1') {
        admin.innerHTML = htmlAktSetupEntraPermGrid(cfg, 'admin');
        admin.dataset.aktPermMounted = '1';
    }
    if (ein && ein.dataset.aktPermMounted !== '1') {
        ein.innerHTML = htmlAktSetupEntraPermGrid(cfg, 'einreicher');
        ein.dataset.aktPermMounted = '1';
    }
    remountAktSetupPermPickers(() => {
        if (window.ms365AktSetup && typeof window.ms365AktSetup.refreshGlance === 'function') {
            window.ms365AktSetup.refreshGlance();
        }
    });
}

function wireSavePerms() {
    const btn = document.getElementById('aktBtnSavePerms');
    if (!btn || btn.dataset.aktPermWired === '1') return;
    btn.dataset.aktPermWired = '1';
    btn.addEventListener('click', () => {
        const next = persistAktPickersToStorage(false);
        if (window.ms365AktSetup && typeof window.ms365AktSetup.refreshGlance === 'function') {
            window.ms365AktSetup.refreshGlance();
        }
        const msg = aktEntraGroupsConfigured(next)
            ? 'Rollen für Schulaktivitäten gespeichert (wie im Listen-Tool).'
            : 'Gespeichert – bitte noch Genehmigung und/oder Einreicher setzen.';
        if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
    });
}

export function initAktSetupPermissions() {
    wireFlowSetupEmailPickers(FLOW_EMAIL_PICKERS, { wiredFlag: 'aktEmailPickerWired' });
    mountRolesStep();
    wireSavePerms();
}

window.ms365AktInitSetupPermissions = initAktSetupPermissions;
window.ms365AktRolesConfigured = () => aktEntraGroupsConfigured(loadEffectiveAktPermissionsConfig());

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => initAktSetupPermissions());
} else {
    initAktSetupPermissions();
}
