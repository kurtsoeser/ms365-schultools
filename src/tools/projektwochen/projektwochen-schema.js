/**
 * SharePoint-Listen-Schema für Projektwochen (Phase 1).
 * Stammdaten (Klassen, Lehrer) aus tenant-settings – keine Lookup-Listen.
 * Spec: docs/projektwochen.md
 *
 * Graph-Hinweis: Single-Line-Text maxLength ≤ 255; Mehrzeiler ohne maxLength + textType plain;
 * number braucht decimalPlaces.
 */

/** @typedef {{ name: string, displayName: string, [k: string]: unknown }} GraphColumnDef */

export const LIST_TITLES = {
    aktionen: 'PW-Aktionen',
    angebote: 'PW-Angebote'
};

/** @type {GraphColumnDef[]} */
export const AKTIONEN_COLUMNS = [
    {
        name: 'AktionId',
        displayName: 'Aktion-ID',
        text: { allowMultipleLines: false, maxLength: 40 }
    },
    {
        name: 'Startdatum',
        displayName: 'Startdatum',
        dateTime: { displayAs: 'default', format: 'dateOnly' }
    },
    {
        name: 'Enddatum',
        displayName: 'Enddatum',
        dateTime: { displayAs: 'default', format: 'dateOnly' }
    },
    {
        name: 'BuchungAbDefault',
        displayName: 'Standard-Buchungsstart',
        dateTime: { displayAs: 'default', format: 'dateTime' }
    },
    {
        name: 'BookingsBusinessId',
        displayName: 'Bookings-Business-ID',
        text: { allowMultipleLines: false, maxLength: 80 }
    },
    {
        name: 'BookingsBusinessName',
        displayName: 'Bookings-Anzeigename',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'Status',
        displayName: 'Status',
        choice: {
            allowTextEntry: false,
            choices: ['entwurf', 'offen', 'geschlossen']
        }
    },
    {
        name: 'Beschreibung',
        displayName: 'Beschreibung',
        text: { allowMultipleLines: true, textType: 'plain' }
    }
];

/** @type {GraphColumnDef[]} */
export const ANGEBOTE_COLUMNS = [
    {
        name: 'AngebotId',
        displayName: 'Angebot-ID',
        text: { allowMultipleLines: false, maxLength: 40 }
    },
    {
        name: 'AktionId',
        displayName: 'Aktion-ID',
        text: { allowMultipleLines: false, maxLength: 40 }
    },
    {
        name: 'Beschreibung',
        displayName: 'Beschreibung',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'HinweisEltern',
        displayName: 'Hinweis Eltern',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'Ort',
        displayName: 'Ort',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'Treffpunkt',
        displayName: 'Treffpunkt',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'Wochentag',
        displayName: 'Wochentag',
        choice: {
            allowTextEntry: false,
            choices: ['Mo', 'Di', 'Mi', 'Do', 'Fr']
        }
    },
    {
        name: 'Datum',
        displayName: 'Datum',
        dateTime: { displayAs: 'default', format: 'dateOnly' }
    },
    {
        name: 'Slot',
        displayName: 'Slot',
        choice: {
            allowTextEntry: false,
            choices: ['ganztags', 'vormittag', 'nachmittag', 'abend']
        }
    },
    {
        name: 'Startzeit',
        displayName: 'Startzeit',
        text: { allowMultipleLines: false, maxLength: 10 }
    },
    {
        name: 'Endzeit',
        displayName: 'Endzeit',
        text: { allowMultipleLines: false, maxLength: 10 }
    },
    {
        name: 'Kapazitaet',
        displayName: 'Max. Teilnehmer',
        number: { decimalPlaces: 'none' }
    },
    {
        name: 'PreisEuro',
        displayName: 'Preis EUR',
        number: { decimalPlaces: 'two' }
    },
    {
        name: 'KostenHinweis',
        displayName: 'Kostenhinweis',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'Zielklassen',
        displayName: 'Zielklassen',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'LehrerCode',
        displayName: 'Lehrer-Kuerzel',
        text: { allowMultipleLines: false, maxLength: 40 }
    },
    {
        name: 'LehrerEmail',
        displayName: 'Lehrer-E-Mail',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'Begleitung',
        displayName: 'Begleitung',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'Kategorie',
        displayName: 'Kategorie',
        choice: {
            allowTextEntry: false,
            choices: ['exkursion', 'workshop', 'kultur', 'sport', 'sonstiges']
        }
    },
    {
        name: 'Status',
        displayName: 'Status',
        choice: {
            allowTextEntry: false,
            choices: ['entwurf', 'beantragt', 'freigegeben', 'abgelehnt', 'abgesagt']
        }
    },
    {
        name: 'BuchungAb',
        displayName: 'Buchung moeglich ab',
        dateTime: { displayAs: 'default', format: 'dateTime' }
    },
    {
        name: 'AblehnungsGrund',
        displayName: 'Ablehnungsgrund',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'BookingsServiceId',
        displayName: 'Bookings-Service-ID',
        text: { allowMultipleLines: false, maxLength: 80 }
    },
    {
        name: 'BookingsBookingUrl',
        displayName: 'Buchungslink',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'SyncStatus',
        displayName: 'Sync-Status',
        text: { allowMultipleLines: false, maxLength: 40 }
    },
    {
        name: 'SyncFehler',
        displayName: 'Sync-Fehler',
        text: { allowMultipleLines: true, textType: 'plain' }
    },
    {
        name: 'SyncAm',
        displayName: 'Sync am',
        dateTime: { displayAs: 'default', format: 'dateTime' }
    },
    {
        name: 'BeantragtVon',
        displayName: 'Beantragt von',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'FreigegebenVon',
        displayName: 'Freigegeben von',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'FreigegebenAm',
        displayName: 'Freigegeben am',
        dateTime: { displayAs: 'default', format: 'dateTime' }
    },
    {
        name: 'NotizIntern',
        displayName: 'Interne Notiz',
        text: { allowMultipleLines: true, textType: 'plain' }
    }
];

export const REQUIRED_COLUMNS = {
    'PW-Aktionen': AKTIONEN_COLUMNS.map((c) => c.name),
    'PW-Angebote': ANGEBOTE_COLUMNS.map((c) => c.name)
};

/**
 * Demo-/Seed-Aktion (nur wenn PW-Aktionen leer).
 * Datum: nächste Kalenderwoche Mo–Fr relativ zu „heute“ – Setup setzt konkrete ISO beim Seed.
 */
export const DEFAULT_AKTION_TEMPLATE = {
    Title: 'Projektwoche (Demo)',
    AktionId: 'pw-demo',
    Status: 'offen',
    Beschreibung: 'Seed-Eintrag aus dem Projektwochen-Setup. Zeitraum und Buchungsstart bei Bedarf anpassen.'
};

/**
 * Graph-Spaltendefinition bereinigen (SharePoint-Limits).
 * @param {GraphColumnDef} def
 */
export function toGraphColumnBody(def) {
    const body = {
        name: def.name,
        displayName: def.displayName || def.name
    };
    if (def.text) {
        const multi = !!def.text.allowMultipleLines;
        body.text = {
            allowMultipleLines: multi,
            textType: def.text.textType || 'plain'
        };
        if (!multi) {
            const max = Number(def.text.maxLength);
            body.text.maxLength = Number.isFinite(max) ? Math.min(255, Math.max(1, max)) : 255;
        }
    } else if (def.number) {
        body.number = {
            decimalPlaces: def.number.decimalPlaces || 'automatic'
        };
    } else if (def.boolean) {
        body.boolean = {};
    } else if (def.dateTime) {
        body.dateTime = {
            displayAs: def.dateTime.displayAs || 'default',
            format: def.dateTime.format || 'dateOnly'
        };
    } else if (def.choice) {
        body.choice = {
            allowTextEntry: !!def.choice.allowTextEntry,
            choices: Array.isArray(def.choice.choices) ? def.choice.choices.slice() : []
        };
    }
    return body;
}

/**
 * @param {string} prefix
 */
export function newEntityId(prefix) {
    const p = String(prefix || 'id').replace(/[^a-z0-9-]/gi, '').slice(0, 12) || 'id';
    const rand =
        typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function'
            ? crypto.randomUUID().replace(/-/g, '').slice(0, 8)
            : String(Date.now().toString(36) + Math.random().toString(36).slice(2, 8)).slice(0, 8);
    return p + '-' + rand;
}

/**
 * Nächsten Montag und Freitag als ISO dateOnly (Lokalzeit).
 * @param {Date} [from]
 * @returns {{ startIso: string, endIso: string, buchungAbIso: string }}
 */
export function nextProjectWeekRange(from) {
    const base = from instanceof Date && !Number.isNaN(from.getTime()) ? new Date(from.getTime()) : new Date();
    const day = base.getDay(); // 0 So … 6 Sa
    const daysUntilMon = day === 1 ? 7 : (8 - day) % 7 || 7;
    const mon = new Date(base.getFullYear(), base.getMonth(), base.getDate() + daysUntilMon);
    const fri = new Date(mon.getFullYear(), mon.getMonth(), mon.getDate() + 4);
    const pad = (n) => String(n).padStart(2, '0');
    const toDate = (d) => d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
    const startIso = toDate(mon);
    const endIso = toDate(fri);
    // Buchungsstart: 14 Tage vor Montag, 08:00 (SharePoint dateTime, ohne Offset-Suffix)
    const open = new Date(mon.getFullYear(), mon.getMonth(), mon.getDate() - 14, 8, 0, 0);
    const buchungAbIso =
        toDate(open) +
        'T' +
        pad(open.getHours()) +
        ':' +
        pad(open.getMinutes()) +
        ':' +
        pad(open.getSeconds());
    return { startIso, endIso, buchungAbIso };
}
