/**
 * ESM-Brücke: Schema + Setup-Status für freistellung-setup.js (IIFE).
 */
import { FREISTELLUNG_COLUMNS, toGraphColumnBody } from '../freistellung-planer/freistellung-planer-schema.js';
import { computeSetupGlance } from './freistellung-setup-status.js';
import { applyFreistellungListPermissions } from '../freistellung-planer/freistellung-planer-list-permissions.js';
import {
    publishPlannerPermissionsToSite,
    remoteConfigPathHint
} from '../freistellung-planer/freistellung-planer-remote-config.js';
import {
    saveAndPublishFreistellungPlannerGroups,
    formatFreistellungPlannerPublishToast
} from '../freistellung-planer/freistellung-permissions-ui.js';

window.ms365FreistellungSchema = {
    columns: FREISTELLUNG_COLUMNS,
    toGraphColumnBody
};
window.ms365FreistellungSetupStatus = {
    computeSetupGlance
};
window.ms365FreistellungListPerms = {
    apply: applyFreistellungListPermissions,
    publishConfig: publishPlannerPermissionsToSite,
    configPathHint: remoteConfigPathHint
};
window.ms365FreistellungPlannerSave = {
    saveAndPublish: saveAndPublishFreistellungPlannerGroups,
    formatToast: formatFreistellungPlannerPublishToast
};
