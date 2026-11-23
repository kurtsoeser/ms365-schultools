/**
 * Ready-Checks für Erste Einrichtung · Blaupause (Wave 2).
 */
import { tenantCountsFromSettings } from './playbook-schuljahresstart-gates.js';

/**
 * @param {Record<string, unknown>|null|undefined} settings
 */
export function blaupauseCountsFromSettings(settings) {
    const base = tenantCountsFromSettings(settings);
    const s = settings && typeof settings === 'object' ? settings : {};
    const planerDefaultSiteUrl = String(
        s.planerDefaultSiteUrl || s.tenantPlanerDefaultSiteUrl || ''
    ).trim();
    const intranetSiteUrl = String(s.intranetSiteUrl || s.tenantIntranetSiteUrl || '').trim();
    const itLib = s.stammdatenItLibrary && typeof s.stammdatenItLibrary === 'object' ? s.stammdatenItLibrary : null;
    const itLibrarySiteUrl = String(
        s.itLibrarySiteUrl ||
            s.tenantItLibrarySiteUrl ||
            (itLib && (itLib.siteUrl || itLib.webUrl || itLib.url)) ||
            ''
    ).trim();
    return {
        ...base,
        planerDefaultSiteUrl,
        intranetSiteUrl,
        itLibrarySiteUrl
    };
}

function hasHttpUrl(url) {
    return /^https?:\/\//i.test(String(url || '').trim());
}

/**
 * @param {ReturnType<typeof blaupauseCountsFromSettings>} counts
 * @returns {Record<string, { status: string, message: string }>}
 */
export function evaluateBlaupauseReadyChecks(counts) {
    const c = counts || {};
    const hasSchool = !!(String(c.schoolName || '').trim() || String(c.domain || '').trim());
    const hasDomain = !!String(c.domain || '').trim();
    const hasYear = !!String(c.year || '').trim();
    const toolSite =
        hasHttpUrl(c.planerDefaultSiteUrl) || hasHttpUrl(c.intranetSiteUrl);
    const classes = Number(c.classes) || 0;
    const students = Number(c.students) || 0;
    const teachers = Number(c.teachers) || 0;
    const subjects = Number(c.subjects) || 0;
    const hasPeople = students >= 1 || teachers >= 1;

    /** @type {Record<string, { status: string, message: string }>} */
    const map = {
        register: hasSchool && hasDomain && hasYear
            ? { status: 'ok', message: 'Profil, Domain und Schuljahr gesetzt' }
            : {
                  status: 'missing',
                  message: !hasDomain
                      ? 'Schuldomain und Schuljahr im Profil setzen'
                      : !hasYear
                        ? 'Aktives Schuljahr setzen'
                        : 'Schulname oder Domain im Profil setzen'
              },
        'tool-site': toolSite
            ? {
                  status: 'ok',
                  message: hasHttpUrl(c.planerDefaultSiteUrl)
                      ? 'Primär-Site (Planer) gesetzt'
                      : 'Intranet-Site gesetzt (Planer-URL prüfen)'
              }
            : {
                  status: 'missing',
                  message: 'Primär-Site unter Stammdaten → Planer & Apps eintragen'
              },
        quelle: { status: 'unknown', message: '' },
        'import-datei': hasPeople
            ? { status: 'ok', message: 'Personen in den Stammdaten vorhanden' }
            : {
                  status: 'warn',
                  message: 'Noch keine Schüler/Lehrkräfte – Import ausstehend?'
              },
        'bildungsportal-info': {
            status: 'ok',
            message: 'Optional – für den Start überspringbar'
        },
        review:
            classes >= 1 && hasPeople
                ? { status: 'ok', message: 'Klassen und Personen vorhanden' }
                : {
                      status: 'missing',
                      message: 'Klassen und Import prüfen, bevor Gruppen verknüpft werden'
                  },
        slg: hasPeople
            ? { status: 'ok', message: 'Personen für SLG vorhanden' }
            : { status: 'missing', message: 'Personenliste für Sammelgruppen nötig' },
        klassen:
            classes >= 1
                ? { status: 'ok', message: 'Klassenliste vorhanden' }
                : { status: 'missing', message: 'Mindestens eine Klasse in den Stammdaten' },
        faecher:
            subjects >= 1
                ? { status: 'ok', message: 'Fächerliste vorhanden' }
                : {
                      status: 'warn',
                      message: 'Fächer optional, aber für Fachteams sinnvoll'
                  },
        hygiene: hasPeople
            ? { status: 'ok', message: 'Daten für Abgleich vorhanden' }
            : { status: 'warn', message: 'Erst nach Import sinnvoll' },
        backup: hasHttpUrl(c.itLibrarySiteUrl)
            ? { status: 'ok', message: 'IT-Sicherungsbibliothek gesetzt' }
            : {
                  status: 'warn',
                  message: 'IT-Sicherung oder Browser-Backup noch offen'
              }
    };

    if (toolSite && !hasHttpUrl(c.planerDefaultSiteUrl) && hasHttpUrl(c.intranetSiteUrl)) {
        map['tool-site'] = {
            status: 'warn',
            message: 'Nur Intranet-URL – Planer-Primär-Site explizit setzen'
        };
    }

    return map;
}

if (typeof window !== 'undefined') {
    window.ms365BlaupausePlaybook = {
        blaupauseCountsFromSettings,
        evaluateBlaupauseReadyChecks
    };
}
