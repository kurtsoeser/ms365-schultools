/**
 * SharePoint-Listen-Schema: Wahlfach-Zuteilung (Linz-Auhof).
 * System of Record auf der Schul-Intranet-Site; lokales Access-Export später aus Ergebnislisten.
 */

/** @typedef {{ name: string, displayName: string, [k: string]: unknown }} GraphColumnDef */

export const LIST_TITLES = {
    runden: 'Wahlfach-Runden',
    angebote: 'Wahlfach-Angebote',
    anmeldungen: 'Wahlfach-Anmeldungen',
    zuteilungen: 'Wahlfach-Zuteilungen'
};

export const RUNDE_STATUS_CHOICES = [
    'entwurf',
    'offen',
    'geschlossen',
    'zuteilung',
    'veroeffentlicht',
    'archiviert'
];

export const ANMELDUNG_STATUS_CHOICES = [
    'eingereicht',
    'zugeordnet',
    'ersatz',
    'offen',
    'storniert'
];

/** @type {GraphColumnDef[]} */
export const RUNDEN_COLUMNS = [
    {
        name: 'RundeId',
        displayName: 'Runden-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'Schuljahr',
        displayName: 'Schuljahr',
        text: { allowMultipleLines: false, maxLength: 20 }
    },
    {
        name: 'AnmeldeVon',
        displayName: 'Anmelde von',
        dateTime: { displayAs: 'default', format: 'dateOnly' }
    },
    {
        name: 'AnmeldeBis',
        displayName: 'Anmelde bis',
        dateTime: { displayAs: 'default', format: 'dateOnly' }
    },
    {
        name: 'Status',
        displayName: 'Status',
        choice: { allowTextEntry: false, choices: [...RUNDE_STATUS_CHOICES] }
    }
];

/** @type {GraphColumnDef[]} */
export const ANGEBOTE_COLUMNS = [
    {
        name: 'AngebotId',
        displayName: 'Angebots-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'RundeId',
        displayName: 'Runden-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'KapazitaetMin',
        displayName: 'Kapazität min.',
        number: {}
    },
    {
        name: 'KapazitaetMax',
        displayName: 'Kapazität max.',
        number: {}
    },
    {
        name: 'ZielJahrgaenge',
        displayName: 'Zieljahrgänge',
        text: { allowMultipleLines: false, maxLength: 80 }
    },
    {
        name: 'Aktiv',
        displayName: 'Aktiv',
        boolean: {}
    },
    {
        name: 'Schuljahr',
        displayName: 'Schuljahr',
        text: { allowMultipleLines: false, maxLength: 20 }
    }
];

/** @type {GraphColumnDef[]} */
export const ANMELDUNGEN_COLUMNS = [
    {
        name: 'AnmeldungId',
        displayName: 'Anmeldungs-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'RundeId',
        displayName: 'Runden-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'SchuelerName',
        displayName: 'Schüler:in',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    {
        name: 'Klasse',
        displayName: 'Klasse',
        text: { allowMultipleLines: false, maxLength: 40 }
    },
    {
        name: 'SchuelerEmail',
        displayName: 'Schüler-E-Mail',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'Wunsch1Id',
        displayName: '1. Wunsch-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'Wunsch2Id',
        displayName: '2. Wunsch-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'Wunsch3Id',
        displayName: '3. Wunsch-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'Wunsch1Text',
        displayName: '1. Wunsch',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    {
        name: 'Wunsch2Text',
        displayName: '2. Wunsch',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    {
        name: 'Wunsch3Text',
        displayName: '3. Wunsch',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    {
        name: 'Status',
        displayName: 'Status',
        choice: { allowTextEntry: false, choices: [...ANMELDUNG_STATUS_CHOICES] }
    },
    {
        name: 'EingereichtAm',
        displayName: 'Eingereicht am',
        dateTime: { displayAs: 'default', format: 'dateTime' }
    },
    {
        name: 'Schuljahr',
        displayName: 'Schuljahr',
        text: { allowMultipleLines: false, maxLength: 20 }
    }
];

/** @type {GraphColumnDef[]} */
export const ZUTEILUNGEN_COLUMNS = [
    {
        name: 'ZuteilungId',
        displayName: 'Zuteilungs-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'RundeId',
        displayName: 'Runden-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'AnmeldungId',
        displayName: 'Anmeldungs-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'AngebotId',
        displayName: 'Angebots-ID',
        text: { allowMultipleLines: false, maxLength: 60 }
    },
    {
        name: 'AngebotText',
        displayName: 'Zugeteiltes Angebot',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    {
        name: 'SchuelerName',
        displayName: 'Schüler:in',
        text: { allowMultipleLines: false, maxLength: 120 }
    },
    {
        name: 'Klasse',
        displayName: 'Klasse',
        text: { allowMultipleLines: false, maxLength: 40 }
    },
    {
        name: 'Begruendung',
        displayName: 'Begründung',
        text: { allowMultipleLines: true, maxLength: 2000 }
    },
    {
        name: 'ZugeteiltVon',
        displayName: 'Zugeteilt von',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'ZugeteiltAm',
        displayName: 'Zugeteilt am',
        dateTime: { displayAs: 'default', format: 'dateTime' }
    },
    {
        name: 'Schuljahr',
        displayName: 'Schuljahr',
        text: { allowMultipleLines: false, maxLength: 20 }
    },
    {
        name: 'AccessExportKey',
        displayName: 'Access-Export-Schlüssel',
        text: { allowMultipleLines: false, maxLength: 80 }
    }
];

export const REQUIRED_COLUMNS = {
    [LIST_TITLES.runden]: RUNDEN_COLUMNS.map((c) => c.name),
    [LIST_TITLES.angebote]: ANGEBOTE_COLUMNS.map((c) => c.name),
    [LIST_TITLES.anmeldungen]: ANMELDUNGEN_COLUMNS.map((c) => c.name),
    [LIST_TITLES.zuteilungen]: ZUTEILUNGEN_COLUMNS.map((c) => c.name)
};

/**
 * @param {GraphColumnDef} def
 */
export function toGraphColumnBody(def) {
    const body = {
        name: def.name,
        displayName: def.displayName || def.name
    };
    if (def.text) body.text = def.text;
    if (def.number) body.number = def.number;
    if (def.boolean) body.boolean = def.boolean;
    if (def.dateTime) body.dateTime = def.dateTime;
    if (def.choice) body.choice = def.choice;
    return body;
}
