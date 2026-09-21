
const KT = window.ms365KursteamTeamNames;
window.ms365AssertModules({ KT }, 'kursteam-team-build.js');

/**
 * Erzeugt die Team-Liste aus gefilterten Unterrichtszeilen (ohne Duplikat-Auflösung).
 * @param {Array} rows ns.filteredData
 * @param {object} options
 */
function resolveFachAndGruppeForTeam(row, options) {
    let fach = row.fach;
    let gruppe = row.gruppe || '';
    const strip = !!options.stripSubjectTrailingDigits;
    const normalizeFn =
        typeof options.normalizeNumberedSubjectFields === 'function'
            ? options.normalizeNumberedSubjectFields
            : null;
    if (strip && normalizeFn) {
        const n = normalizeFn(fach, gruppe);
        if (n && n.changed) {
            fach = n.fach;
            gruppe = n.gruppe;
        }
    }
    return { fach, gruppe };
}

function buildTeamEntriesFromRows(rows, options) {
    const yearPrefix = options.yearPrefix;
    const emailDomain = options.emailDomain;
    const separator = options.separator != null ? options.separator : ' | ';
    const pattern = options.pattern;
    const combineClassNames = options.combineClassNames;
    const buildGruppenmailBase = options.buildGruppenmailBase;
    const formatKlasseSegmentForGruppenmail = options.formatKlasseSegmentForGruppenmail;
    const sanitizeGruppeForMail = options.sanitizeGruppeForMail;
    const INVALID_CHARS_REPLACE = options.INVALID_CHARS_REPLACE;
    const INVALID_CHARS_TEST = options.INVALID_CHARS_TEST;
    const teacherEmailMapping = options.teacherEmailMapping || {};

    const isCombinedFn =
        typeof options.isCombinedClassCell === 'function'
            ? options.isCombinedClassCell
            : function (raw) {
                  return /[,;~]/.test(String(raw || '')) || /(?:\d+[A-Za-z]+){2,}/.test(String(raw || ''));
              };

    return (rows || []).map((row) => {
        const combined = isCombinedFn(row.klasse);
        let klasseForName = row.klasse;
        if (combined) klasseForName = combineClassNames(row.klasse);

        const resolved = resolveFachAndGruppeForTeam(row, options);
        const fachForName = resolved.fach;
        const gruppeForName = resolved.gruppe;

        const teamName = pattern
            ? KT.buildTeamNameFromPattern(pattern, {
                  yearPrefix,
                  klasse: klasseForName,
                  fach: fachForName,
                  gruppe: gruppeForName,
                  lehrer: row.lehrer
              })
            : `${yearPrefix}${separator}${klasseForName}${separator}${fachForName}`;

        // Bei Mehrklassen keinen Klassen-Nick einer Einzelklasse übernehmen (verhindert -hakb- u. ä.)
        let klasseForGruppenmail = klasseForName;
        if (!combined) {
            try {
                const adv = window.ms365AppDataV2;
                if (adv && typeof adv.getClassTeamGruppenmailForKlasse === 'function') {
                    const stable = adv.getClassTeamGruppenmailForKlasse(row.klasse);
                    if (stable) klasseForGruppenmail = stable;
                }
            } catch {
                // ignore
            }
        }
        const mailCtx = {
            yearPrefix,
            klasse: klasseForGruppenmail,
            fach: fachForName,
            gruppe: gruppeForName,
            lehrer: row.lehrer
        };
        const mailHelpers = {
            formatKlasse: formatKlasseSegmentForGruppenmail,
            sanitizeGruppe: sanitizeGruppeForMail
        };
        const gruppenmailRaw = pattern && typeof KT.buildGruppenmailFromPattern === 'function'
            ? KT.buildGruppenmailFromPattern(pattern, mailCtx, mailHelpers)
            : buildGruppenmailBase(yearPrefix, klasseForGruppenmail, fachForName, gruppeForName);

        const originalGruppenmail = gruppenmailRaw;
        let gruppenmail = gruppenmailRaw.replace(INVALID_CHARS_REPLACE, '');

        let besitzer = '';
        const lehrerCode = row.lehrer.toUpperCase().trim();
        if (teacherEmailMapping[lehrerCode]) {
            besitzer = teacherEmailMapping[lehrerCode];
        } else {
            besitzer = row.lehrer.toLowerCase().trim().replace(/\s+/g, '.');
            besitzer = besitzer.replace(INVALID_CHARS_REPLACE, '');
            if (!besitzer.includes('@')) besitzer += emailDomain;
        }

        const hasInvalidChars = INVALID_CHARS_TEST.test(originalGruppenmail);
        const isValid = !hasInvalidChars && teamName && gruppenmail && besitzer && gruppenmail.length > 0;
        const mappingUsed = !!teacherEmailMapping[lehrerCode];

        return {
            teamName,
            gruppenmail,
            besitzer,
            isValid,
            error: hasInvalidChars ? 'Ungültige Zeichen in Gruppenmail' : !isValid ? 'Unvollständige Daten' : null,
            originalClass: row.klasse,
            fach: fachForName,
            fachOriginal: row.fach,
            gruppe: gruppeForName,
            mappingUsed,
            lehrerCode,
            mailNicknameAdjusted: false
        };
    });
}

window.ms365KursteamTeamBuild = {
    buildTeamEntriesFromRows,
    resolveFachAndGruppeForTeam
};
