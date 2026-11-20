/**
 * ESM-Brücke: Schema + Setup-Status für freistellung-setup.js (IIFE).
 */
import { FREISTELLUNG_COLUMNS, toGraphColumnBody } from '../freistellung-planer/freistellung-planer-schema.js';
import { computeSetupGlance, effectiveFreistellungFlowAccount } from './freistellung-setup-status.js';
import {
    applyFreistellungListPermissions,
    grantFreistellungFlowServiceAccountOnList
} from '../freistellung-planer/freistellung-planer-list-permissions.js';
import {
    publishPlannerPermissionsToSite,
    remoteConfigPathHint,
    classCatalogFromSchoolStammdaten
} from '../freistellung-planer/freistellung-planer-remote-config.js';
import {
    loadPermissionsConfig,
    normalizePermissionsConfig
} from '../freistellung-planer/freistellung-planer-permissions.js';
import {
    saveAndPublishFreistellungPlannerGroups,
    formatFreistellungPlannerPublishToast,
    initFreistellungSetupPermissions
} from '../freistellung-planer/freistellung-permissions-ui.js';
import { wireFlowSetupEmailPickers } from '../../shared/flow-setup-email-pickers.js';

window.ms365FreistellungSchema = {
    columns: FREISTELLUNG_COLUMNS,
    toGraphColumnBody
};
function klassenStepConfigured() {
    if (classCatalogFromSchoolStammdaten().length) return true;
    const c = normalizePermissionsConfig(loadPermissionsConfig());
    return (c.classCatalog && c.classCatalog.length) > 0;
}

function einstellungenStepConfigured() {
    try {
        const setup = JSON.parse(localStorage.getItem('ms365-freistellung-setup-v1') || '{}');
        return !!String(setup.listId || '').trim();
    } catch {
        return false;
    }
}

function permsStepConfigured() {
    const c = normalizePermissionsConfig(loadPermissionsConfig());
    return !!(
        String(c.groupKvId || '').trim() ||
        String(c.groupDirektionId || '').trim() ||
        String(c.groupSchuelerId || '').trim() ||
        (c.direktionGroups && c.direktionGroups.length) ||
        (c.kvGroups && c.kvGroups.length) ||
        (c.schuelerGroups && c.schuelerGroups.length) ||
        (c.direktionUsers && c.direktionUsers.length) ||
        (c.kvUsers && c.kvUsers.length) ||
        (c.schuelerUsers && c.schuelerUsers.length)
    );
}

window.ms365FreistellungSetupStatus = {
    computeSetupGlance,
    effectiveFreistellungFlowAccount,
    klassenStepConfigured,
    einstellungenStepConfigured,
    permsStepConfigured
};
window.ms365FreistellungListPerms = {
    apply: applyFreistellungListPermissions,
    grantFlowServiceAccount: grantFreistellungFlowServiceAccountOnList,
    publishConfig: publishPlannerPermissionsToSite,
    configPathHint: remoteConfigPathHint
};
window.ms365FreistellungPlannerSave = {
    saveAndPublish: saveAndPublishFreistellungPlannerGroups,
    formatToast: formatFreistellungPlannerPublishToast
};
window.ms365FreistellungInitSetupPermissions = initFreistellungSetupPermissions;

/** MS365-Personenpicker für Flow-/Mail-Konten im Setup. */
const FLOW_EMAIL_PICKERS = [
    {
        inputId: 'frFlowServiceAccount',
        btnId: 'frFlowServiceAccountPick',
        title: 'Technik-Konto für Power Automate',
        hint: 'Dienstkonto, mit dem der Flow importiert und die Connections verbunden werden.'
    },
    {
        inputId: 'frEmailDirektion',
        btnId: 'frEmailDirektionPick',
        title: 'Direktion (2. Genehmiger)',
        hint: 'Person oder Postfach für den zweiten Genehmigungsschritt und den Fall KV = Direktion.'
    },
    {
        inputId: 'frEmailMailbox',
        btnId: 'frEmailMailboxPick',
        title: 'Absender für Status-Mails',
        hint: 'Technik-Konto oder freigegebenes Postfach als Absender der Status-Mails.'
    }
];

function wireFreistellungSetupEmailPickers() {
    wireFlowSetupEmailPickers(FLOW_EMAIL_PICKERS, { wiredFlag: 'frEmailPickerWired' });
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', wireFreistellungSetupEmailPickers);
} else {
    wireFreistellungSetupEmailPickers();
}
