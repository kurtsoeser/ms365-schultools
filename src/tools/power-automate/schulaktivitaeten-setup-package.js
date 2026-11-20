/**
 * Reine Parametrisierung der Schulaktivitäten-Flow-Vorlage (Multitenant).
 * Vorlage in assets enthält feste Platzhalter – beim ZIP-Bau werden Site/Liste/Mails
 * durch die Werte aus dem Setup-Formular ersetzt.
 */

export const AKT_FLOW_ASSET_ID = 'a7f3c91e-2b4d-4e8a-9c1f-6d5e8b0a4f27';

/** Muss 1:1 zu scripts/build-akt-pa-template.mjs SOURCE passen. */
export const AKT_FLOW_SOURCE = {
    siteUrl: 'https://kurtrocks.sharepoint.com/sites/MS365-Schultools',
    /**
     * SharePoint-Connector „table“ = Listen-GUID (wie Freistellung).
     * Anzeigenamen führen beim Legacy-Import oft zu GetTable 404 „List not found“.
     */
    listId: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee',
    listTitle: 'Schulaktivitaeten',
    emailDirektion: 'direktor@ms365.schule',
    emailMailbox: 'automate@ms365.schule',
    connectionOwner: 'kurt@kurtsoeser.at'
};

/** @deprecated Alias für ältere Tests/Skripte */
export const AKT_FLOW_LEGACY_LIST_GUID = AKT_FLOW_SOURCE.listId;

const GUID_RE =
    /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * @param {string} raw
 * @returns {string}
 */
export function normalizeSharePointListGuid(raw) {
    const s = String(raw || '')
        .trim()
        .replace(/^\{|\}$/g, '');
    if (!GUID_RE.test(s)) return '';
    return s.toLowerCase();
}

/**
 * @param {string} haystack
 * @param {string} needle
 * @param {string} replacement
 */
export function replaceAllLiteral(haystack, needle, replacement) {
    if (!needle) return haystack;
    return String(haystack).split(needle).join(replacement == null ? '' : String(replacement));
}

/**
 * @param {string|object} rawDef
 * @param {{
 *   siteUrl: string,
 *   listId: string,
 *   listName?: string,
 *   emailDirektion?: string,
 *   emailMailbox?: string,
 *   flowServiceAccount?: string,
 *   flowDisplayName?: string,
 *   mailAsTechnikUser?: boolean,
 *   connectionOwner?: string
 * }} cfg
 * @param {{
 *   flowAccount?: string,
 *   convertSharedMailboxToUserSend?: (definitionRoot: object, useUserMailboxSend: boolean) => void,
 *   newFlowId?: string
 * }} [opts]
 */
export function applyAktFlowParamsToDefinition(rawDef, cfg, opts) {
    const options = opts && typeof opts === 'object' ? opts : {};
    const siteUrl = String((cfg && cfg.siteUrl) || '')
        .trim()
        .replace(/\/+$/, '');
    if (!siteUrl) throw new Error('SharePoint-Website fehlt für das Flow-Paket.');

    const listId = normalizeSharePointListGuid((cfg && cfg.listId) || '');
    if (!listId) {
        throw new Error(
            'Listen-ID (GUID) fehlt für das Flow-Paket – zuerst Liste anlegen/prüfen, dann ZIP erzeugen.'
        );
    }
    if (listId === normalizeSharePointListGuid(AKT_FLOW_SOURCE.listId)) {
        throw new Error('Listen-ID ist noch der Vorlagen-Platzhalter – echte GUID der Zielliste nötig.');
    }

    const emailDirektion = String((cfg && cfg.emailDirektion) || '').trim().toLowerCase();
    const emailMailbox = String((cfg && cfg.emailMailbox) || '').trim().toLowerCase();
    const flowAccount = String(
        options.flowAccount ||
            (cfg && cfg.connectionOwner) ||
            (cfg && cfg.flowServiceAccount) ||
            ''
    )
        .trim()
        .toLowerCase();
    const flowDisplayName =
        String((cfg && cfg.flowDisplayName) || '').trim() ||
        'Schulaktivitäten - Genehmigung Direktion';

    const sourceSite = String(AKT_FLOW_SOURCE.siteUrl || '')
        .trim()
        .replace(/\/+$/, '');
    const siteSameAsPlaceholder = sourceSite.toLowerCase() === siteUrl.toLowerCase();

    let s = typeof rawDef === 'string' ? rawDef : JSON.stringify(rawDef);
    // Auch Varianten mit trailing slash ersetzen
    s = replaceAllLiteral(s, AKT_FLOW_SOURCE.siteUrl + '/', siteUrl + '/');
    s = replaceAllLiteral(s, AKT_FLOW_SOURCE.siteUrl, siteUrl);
    s = replaceAllLiteral(s, sourceSite + '/', siteUrl + '/');
    s = replaceAllLiteral(s, sourceSite, siteUrl);
    s = replaceAllLiteral(s, AKT_FLOW_SOURCE.listId, listId);
    // Alte Vorlagen mit Anzeigename als table
    if (AKT_FLOW_SOURCE.listTitle) {
        s = replaceAllLiteral(s, '"table":"' + AKT_FLOW_SOURCE.listTitle + '"', '"table":"' + listId + '"');
        s = replaceAllLiteral(
            s,
            '"table": "' + AKT_FLOW_SOURCE.listTitle + '"',
            '"table": "' + listId + '"'
        );
    }
    if (emailDirektion) s = replaceAllLiteral(s, AKT_FLOW_SOURCE.emailDirektion, emailDirektion);
    if (emailMailbox) s = replaceAllLiteral(s, AKT_FLOW_SOURCE.emailMailbox, emailMailbox);
    if (flowAccount) s = replaceAllLiteral(s, AKT_FLOW_SOURCE.connectionOwner, flowAccount);

    // Platzhalter-Site nur bemängeln, wenn Ziel-Site anders ist (sonst ist Gleichheit korrekt)
    if (!siteSameAsPlaceholder) {
        if (s.indexOf(AKT_FLOW_SOURCE.siteUrl) !== -1 || s.indexOf(sourceSite) !== -1) {
            throw new Error(
                'Flow-Paket: Platzhalter-Site wurde nicht ersetzt – Vorlage und Setup-SOURCE prüfen.'
            );
        }
    }
    if (s.toLowerCase().indexOf(AKT_FLOW_SOURCE.listId.toLowerCase()) !== -1) {
        throw new Error('Flow-Paket: Platzhalter-Listen-GUID wurde nicht ersetzt.');
    }

    const obj = JSON.parse(s);
    assertAktFlowDefinitionMatchesSchool(obj, { siteUrl, listId });
    if (obj.properties) {
        obj.properties.displayName = flowDisplayName;
        const newId =
            options.newFlowId ||
            (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function'
                ? crypto.randomUUID()
                : 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
                      const r = (Math.random() * 16) | 0;
                      const v = c === 'x' ? r : (r & 0x3) | 0x8;
                      return v.toString(16);
                  }));
        obj.name = newId;
        obj.id = '/providers/Microsoft.Flow/flows/' + newId;
        if (
            obj.properties.definition &&
            typeof options.convertSharedMailboxToUserSend === 'function'
        ) {
            options.convertSharedMailboxToUserSend(
                obj.properties.definition,
                (cfg && cfg.mailAsTechnikUser) !== false
            );
        }
    }
    return obj;
}

/**
 * Prüft, ob dataset/table in der Definition der Ziel-Schule entsprechen.
 * @param {object} definitionWrapper
 * @param {{ siteUrl: string, listId: string }} cfg
 */
export function assertAktFlowDefinitionMatchesSchool(definitionWrapper, cfg) {
    const def =
        (definitionWrapper && definitionWrapper.properties && definitionWrapper.properties.definition) ||
        definitionWrapper;
    const params =
        def &&
        def.triggers &&
        def.triggers.Item_Is_Created &&
        def.triggers.Item_Is_Created.inputs &&
        def.triggers.Item_Is_Created.inputs.parameters;
    if (!params) throw new Error('Flow-Definition ohne SharePoint-Trigger-Parameter.');
    const wantSite = String(cfg.siteUrl || '')
        .trim()
        .replace(/\/+$/, '');
    const wantTable = normalizeSharePointListGuid(cfg.listId);
    const gotSite = String(params.dataset || '')
        .trim()
        .replace(/\/+$/, '');
    const gotTable = normalizeSharePointListGuid(params.table);
    if (gotSite.toLowerCase() !== wantSite.toLowerCase()) {
        throw new Error('Paket-Site „' + gotSite + '“ ≠ Setup „' + wantSite + '“.');
    }
    if (!wantTable || gotTable !== wantTable) {
        throw new Error(
            'Paket-Listen-GUID „' +
                String(params.table || '') +
                '“ ≠ Setup „' +
                String(cfg.listId || '') +
                '“.'
        );
    }
    return true;
}
