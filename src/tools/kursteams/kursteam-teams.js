(function () {
    'use strict';

    const ns = (window.ms365Kursteam = window.ms365Kursteam || {});

    if (typeof window.ms365AssertModules !== 'function') {
        throw new Error(
            'ms365-module-guard.js muss vor kursteam-teams.js geladen werden (tools/kursteams.html).'
        );
    }

    const KUI = window.ms365KursteamTeamNameBuilderUI;
    const KSub = window.ms365KursteamSubjectFilterUI;
    const KActions = window.ms365KursteamTeamsActions;
    window.ms365AssertModules(
        {
            teamNames: window.ms365KursteamTeamNames,
            KUI,
            KSub,
            KActions
        },
        'kursteam-teams.js'
    );

    KUI.mount(ns);
    KSub.mount(ns);
    KActions.mount(ns);

    window.startKursteamFromWebuntis = ns.startKursteamFromWebuntis;
    window.startKursteamManual = ns.startKursteamManual;
    window.addManualDataRow = ns.addManualDataRow;
    window.addManualDataRowInline = ns.addManualDataRowInline;
    window.applyFilters = ns.applyFilters;
    window.resetFilters = ns.resetFilters;
    window.generateTeamNames = ns.generateTeamNames;
    window.addManualKursteamTeam = ns.addManualKursteamTeam;

    if (typeof ns.refreshSubjectFilterUI === 'function') ns.refreshSubjectFilterUI();
    if (typeof ns.renderTeamNameBuilder === 'function') ns.renderTeamNameBuilder();
})();
