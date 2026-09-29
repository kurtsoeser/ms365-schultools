/** UI-State OneNote-Verteilung (Analyse 02 Phase B). */
export const ui = {
    /** @type {1|2|3|4} */
    step: 1,
    /** @type {'central'|'me'|'team'|'site'} */
    srcMode: 'central',
    /** @type {'catalog-api'|'catalog-snapshot'|'site-graph'|''} */
    srcVia: '',
    /** @type {Set<string>} IDs im Schul-Snapshot */
    publishedNotebookIds: new Set(),
    /** @type {Map<string, string>} notebookId → publishedAt (ISO) */
    publishedAtById: new Map(),
    /** @type {Set<string>} Mehrfachauswahl zum Veröffentlichen */
    publishPickIds: new Set(),
    /** @type {{ id: string, displayName: string, webUrl: string }|null} */
    srcSite: null,
    /** @type {{ id: string, displayName: string, mailNickname?: string }|null} */
    srcGroup: null,
    /** @type {Array<{ id: string, displayName: string }>} */
    onSrcNotebooks: [],
    /** @type {{ sections: Array, groups: Array }|null} */
    onSrcTree: null,
    /** @type {Set<string>} */
    onSrcExpanded: new Set(),
    /** @type {Set<string>} */
    onSrcChecked: new Set(),
    /** @type {string} zuletzt geladene Quell-Notizbuch-ID (Struktur) */
    loadedSrcNotebookId: '',
    /** @type {number} */
    srcTreeLoadGen: 0,
    /** @type {string} */
    previewSectionId: '',
    /** @type {string} */
    previewPageId: '',
    /** @type {Array<{ id: string, title: string, webUrl?: string }>} */
    previewPages: [],
    /**
     * @type {Array<{
     *   teamId: string,
     *   teamName: string,
     *   mailNickname: string,
     *   notebookId: string,
     *   notebookName: string,
     *   loading?: boolean,
     *   warn?: string
     * }>}
     */
    selectedTeams: [],
    /** @type {'contentLibrary'|'teacherOnly'|'collaboration'} */
    destKind: 'contentLibrary',
    /**
     * @type {Array<{
     *   key: string,
     *   teamId: string,
     *   teamName: string,
     *   notebookId: string,
     *   notebookName: string,
     *   sectionGroupId: string,
     *   sectionGroupName: string,
     *   status?: string,
     *   statusKind?: string
     * }>}
     */
    onTargets: [],
    buildingTargets: false
};
