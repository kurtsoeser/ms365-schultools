/**
 * Shared mutable state – Personen-Verwaltung (Analyse 02 Phase B).
 */
export const pv = {
    USER_LIST_SELECT:
        'id,displayName,givenName,surname,mail,mailNickname,userPrincipalName,jobTitle,department,' +
        'officeLocation,mobilePhone,businessPhones,companyName,preferredLanguage,accountEnabled,' +
        'streetAddress,city,postalCode,country,createdDateTime,userType,assignedLicenses,usageLocation,' +
        'onPremisesSyncEnabled,onPremisesLastSyncDateTime,onPremisesSamAccountName,onPremisesDomainName,onPremisesSecurityIdentifier',
    get USER_REFRESH_SELECT() {
        return this.USER_LIST_SELECT;
    },
    GROUP_MEMBEROF_SELECT: 'id,displayName,mail,mailNickname,groupTypes,securityEnabled,mailEnabled',
    AD_FLAGS_KEY: 'ms365-pv-ad-flags-v1',
    SESSION_CACHE_KEY: 'ms365-pv-users-cache-v1',
    SESSION_CACHE_MAX_AGE_MS: 30 * 60 * 1000,
    loadedUsers: [],
    selectedUserId: null,
    pendingTabAfterSelect: '',
    activeTab: 'profil',
    cachedGroupsForSelection: null,
    profileEditMode: false,
    subscribedSkus: [],
    subscribedSkusOk: false,
    licenseBusy: false,
    groupBusy: false
};
