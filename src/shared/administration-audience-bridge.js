import {
    ADMIN_TIER_SCHULLEITUNG,
    ADMIN_TIER_VERWALTUNG,
    collectAdminEmailsForTier,
    inferAdminTierForRole,
    splitAdminEmailsByAudienceTier
} from './administration-audience-logic.js';
import {
    BUILTIN_AUDIENCE_SCHULLEITUNG,
    BUILTIN_AUDIENCE_VERWALTUNG,
    addCustomVerwaltungAudienceGroup,
    audienceGroupIdsForEmail,
    collectEmailsForAudienceGroup,
    ensureAdminAudienceOnSettings,
    normalizeAdminAudienceMemberships,
    normalizeVerwaltungAudienceGroups,
    removeCustomVerwaltungAudienceGroup,
    resolveAudienceGroupGraphId,
    setAudienceGroupsForEmail,
    splitAdminEmailsByAudienceGroups
} from './administration-audience-groups.js';
import {
    ADMIN_ROLE_M365_GROUP,
    ADMIN_ROLE_M365_NONE,
    ADMIN_ROLE_M365_SHARED_MAILBOX,
    ADMIN_ROLE_SEAT_MULTI,
    ADMIN_ROLE_SEAT_SINGLE,
    adminRoleM365KindShortLabel,
    adminRolePolicyFromRecord,
    adminRoleSeatModeLabel,
    canAddPersonToAdminRole,
    inferAdminRoleSeatMode,
    normalizeAdminRoleM365Kind,
    normalizeAdminRoleM365Resource,
    normalizeAdminRoleSeatMode
} from './administration-role-policy.js';

window.ms365AdministrationAudience = {
    ADMIN_TIER_SCHULLEITUNG,
    ADMIN_TIER_VERWALTUNG,
    BUILTIN_AUDIENCE_SCHULLEITUNG,
    BUILTIN_AUDIENCE_VERWALTUNG,
    collectAdminEmailsForTier,
    inferAdminTierForRole,
    splitAdminEmailsByAudienceTier,
    splitAdminEmailsByAudienceGroups,
    ensureAdminAudienceOnSettings,
    normalizeVerwaltungAudienceGroups,
    normalizeAdminAudienceMemberships,
    collectEmailsForAudienceGroup,
    audienceGroupIdsForEmail,
    setAudienceGroupsForEmail,
    addCustomVerwaltungAudienceGroup,
    removeCustomVerwaltungAudienceGroup,
    resolveAudienceGroupGraphId,
    ADMIN_ROLE_SEAT_SINGLE,
    ADMIN_ROLE_SEAT_MULTI,
    ADMIN_ROLE_M365_NONE,
    ADMIN_ROLE_M365_GROUP,
    ADMIN_ROLE_M365_SHARED_MAILBOX,
    normalizeAdminRoleSeatMode,
    inferAdminRoleSeatMode,
    normalizeAdminRoleM365Kind,
    normalizeAdminRoleM365Resource,
    adminRolePolicyFromRecord,
    canAddPersonToAdminRole,
    adminRoleSeatModeLabel,
    adminRoleM365KindShortLabel
};
