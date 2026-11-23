/**
 * Ready-Checks für Tool- und App-Playbooks (Wave 2/3).
 */
import { tenantCountsFromSettings } from './playbook-schuljahresstart-gates.js';

/**
 * @param {Record<string, unknown>|null|undefined} settings
 */
function planerSiteFromSettings(settings) {
    const s = settings && typeof settings === 'object' ? settings : {};
    const url = String(s.planerDefaultSiteUrl || s.tenantPlanerDefaultSiteUrl || '').trim();
    const intranet = String(s.intranetSiteUrl || s.tenantIntranetSiteUrl || '').trim();
    if (/^https?:\/\//i.test(url)) return url;
    if (/^https?:\/\//i.test(intranet)) return intranet;
    return '';
}

function siteReady(settings, id) {
    const site = planerSiteFromSettings(settings);
    if (id === 'site' || id === 'listen' || id === 'liste' || id === 'hub') {
        return site
            ? { status: 'ok', message: 'Primär-/Intranet-Site gesetzt' }
            : {
                  status: 'warn',
                  message: 'Primär-Site fehlt – zuerst Blaupause / Planer & Apps'
              };
    }
    return null;
}

/**
 * Generisch für App-Playbooks (Schulaktivitäten, PW, Einsammler, …).
 * Bekannte Step-IDs bekommen Site-/Stammdaten-Hinweise; Rest unknown/optional ok.
 * @param {Record<string, unknown>|null|undefined} settings
 * @param {string[]} [stepIds]
 */
export function evaluateAppPlaybookReadyChecks(settings, stepIds) {
    const c = tenantCountsFromSettings(settings);
    const classes = Number(c.classes) || 0;
    const students = Number(c.students) || 0;
    const site = planerSiteFromSettings(settings);
    const ids = Array.isArray(stepIds) && stepIds.length
        ? stepIds
        : [
              'site',
              'listen',
              'rechte',
              'bookings',
              'test',
              'mail',
              'schulung',
              'konzept',
              'verteiler',
              'hub',
              'lehrer',
              'termine',
              'sa',
              'pw',
              'vertretung',
              'nutzen'
          ];

    /** @type {Record<string, { status: string, message: string }>} */
    const map = {};
    ids.forEach(function (id) {
        const fromSite = siteReady(settings, id);
        if (fromSite) {
            map[id] = fromSite;
            return;
        }
        if (id === 'verteiler') {
            map[id] =
                students >= 1
                    ? { status: 'ok', message: 'Schüler in Stammdaten – Verteiler möglich' }
                    : {
                          status: 'warn',
                          message: 'Schüler/Guardians für Verteiler prüfen'
                      };
            return;
        }
        if (id === 'rechte' || id === 'bookings' || id === 'konzept') {
            map[id] = site
                ? { status: 'ok', message: 'Site vorhanden – im Tool konfigurieren' }
                : {
                      status: 'warn',
                      message: 'Primär-Site zuerst setzen'
                  };
            return;
        }
        if (id === 'lehrer' || id === 'termine' || id === 'nutzen') {
            map[id] =
                classes >= 1 || site
                    ? { status: 'ok', message: 'Grundlage vorhanden' }
                    : { status: 'warn', message: 'Stammdaten / Site prüfen' };
            return;
        }
        if (
            id === 'schulung' ||
            id === 'mail' ||
            id === 'sa' ||
            id === 'pw' ||
            id === 'vertretung' ||
            id === 'test'
        ) {
            map[id] = {
                status: 'ok',
                message: id === 'test' ? 'Manuell im Tool prüfen' : 'Optional / manuell'
            };
            return;
        }
        map[id] = { status: 'unknown', message: '' };
    });
    return map;
}

/**
 * @param {Record<string, unknown>|null|undefined} settings
 */
export function evaluateSchularbeitenPlaybookReadyChecks(settings) {
    const c = tenantCountsFromSettings(settings);
    const classes = Number(c.classes) || 0;
    const subjects = Number(c.subjects) || 0;
    const site = planerSiteFromSettings(settings);

    return {
        stammdaten:
            classes >= 1 && subjects >= 1
                ? { status: 'ok', message: 'Klassen und Fächer vorhanden' }
                : {
                      status: 'missing',
                      message: 'Fächer und Klassen in den Stammdaten pflegen'
                  },
        listen: site
            ? { status: 'ok', message: 'Primär-Site gesetzt – Listen dort anlegen' }
            : {
                  status: 'warn',
                  message: 'Primär-Site fehlt – zuerst Blaupause / Planer & Apps'
              },
        regelwerk: { status: 'unknown', message: '' },
        berechtigungen: { status: 'unknown', message: '' },
        test: { status: 'unknown', message: '' },
        intranet: {
            status: 'ok',
            message: 'Optional – nach Go-Live'
        },
        schulung: {
            status: 'ok',
            message: 'Optional – Einweisung / Mail'
        }
    };
}

/**
 * @param {Record<string, unknown>|null|undefined} settings
 */
export function evaluateFreistellungenPlaybookReadyChecks(settings) {
    const c = tenantCountsFromSettings(settings);
    const classes = Number(c.classes) || 0;
    const site = planerSiteFromSettings(settings);

    return {
        konzept: { status: 'ok', message: 'Konzeptseite lesen' },
        'pa-basis': { status: 'unknown', message: '' },
        vorbereitung: { status: 'unknown', message: '' },
        liste: site
            ? { status: 'ok', message: 'Primär-Site gesetzt' }
            : {
                  status: 'warn',
                  message: 'Primär-Site fehlt – Listen brauchen eine Tool-Site'
              },
        flow: { status: 'unknown', message: '' },
        'planer-config':
            classes >= 1
                ? { status: 'ok', message: 'Klassen für Kategorien vorhanden' }
                : {
                      status: 'warn',
                      message: 'Klassenliste für Genehmiger/Kategorien sinnvoll'
                  },
        live: { status: 'unknown', message: '' }
    };
}

/**
 * @param {Record<string, unknown>|null|undefined} settings
 */
export function evaluateKursteamsPlaybookReadyChecks(settings) {
    const c = tenantCountsFromSettings(settings);
    const classes = Number(c.classes) || 0;
    const subjects = Number(c.subjects) || 0;
    const teachers = Number(c.teachers) || 0;
    const students = Number(c.students) || 0;

    return {
        import:
            teachers >= 1 || subjects >= 1
                ? { status: 'ok', message: 'Lehrkräfte/Fächer in Stammdaten' }
                : {
                      status: 'missing',
                      message: 'Import oder Stammdaten zuerst (Blaupause / Schuljahresstart)'
                  },
        klassen:
            classes >= 1 && students >= 1
                ? { status: 'ok', message: 'Klassen und Schüler vorhanden' }
                : {
                      status: 'warn',
                      message: 'Klassen/Schüler für SLG prüfen'
                  },
        kursteams:
            classes >= 1 && subjects >= 1
                ? { status: 'ok', message: 'Klassen und Fächer – Kursteams sinnvoll' }
                : {
                      status: 'missing',
                      message: 'Klassen und Fächer nötig'
                  },
        templates: { status: 'ok', message: 'Optional' },
        monitor:
            classes >= 1
                ? { status: 'ok', message: 'Klassenliste steht' }
                : { status: 'warn', message: 'Klassenliste zuerst' },
        katalog: { status: 'ok', message: 'Optional' }
    };
}

/**
 * Eltern + Elternsprechtag (Guardians/Verteiler).
 * @param {Record<string, unknown>|null|undefined} settings
 */
export function evaluateElternPlaybookReadyChecks(settings) {
    const c = tenantCountsFromSettings(settings);
    const students = Number(c.students) || 0;
    const classes = Number(c.classes) || 0;

    return {
        guardians:
            students >= 1
                ? { status: 'ok', message: 'Schüler in Stammdaten – Guardians prüfen' }
                : {
                      status: 'missing',
                      message: 'Schüler/Guardians importieren'
                  },
        verteiler:
            students >= 1 && classes >= 1
                ? { status: 'ok', message: 'Klassen und Schüler vorhanden' }
                : {
                      status: 'warn',
                      message: 'Klassen und Schüler für Verteiler'
                  },
        bookings: { status: 'ok', message: 'Optional / Sprechtag-Playbook' },
        intranet: { status: 'ok', message: 'Optional' },
        pa: { status: 'ok', message: 'Optional' },
        termin: { status: 'unknown', message: '' },
        lehrer: { status: 'unknown', message: '' },
        kommunikation: { status: 'unknown', message: '' },
        nachbereitung: { status: 'ok', message: 'Optional' }
    };
}

if (typeof window !== 'undefined') {
    window.ms365ToolPlaybookReady = {
        evaluateAppPlaybookReadyChecks,
        evaluateSchularbeitenPlaybookReadyChecks,
        evaluateFreistellungenPlaybookReadyChecks,
        evaluateKursteamsPlaybookReadyChecks,
        evaluateElternPlaybookReadyChecks
    };
}
