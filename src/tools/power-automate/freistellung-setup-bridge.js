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
    remoteConfigPathHint
} from '../freistellung-planer/freistellung-planer-remote-config.js';
import {
    saveAndPublishFreistellungPlannerGroups,
    formatFreistellungPlannerPublishToast,
    initFreistellungSetupPermissions
} from '../freistellung-planer/freistellung-permissions-ui.js';

window.ms365FreistellungSchema = {
    columns: FREISTELLUNG_COLUMNS,
    toGraphColumnBody
};
window.ms365FreistellungSetupStatus = {
    computeSetupGlance,
    effectiveFreistellungFlowAccount
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
