/**
 * Soft-Error-Klassifikation für election-api (DOM-frei, testbar).
 * Nie als Wipe-Signal missbrauchen – UI behält Meta/Demo-Cache.
 */

/**
 * @typedef {{
 *   code: 'auth_required'|'forbidden'|'not_found'|'not_open'|'network'|'config'|'unknown',
 *   message: string,
 *   status: number,
 *   keepLocal: true,
 *   raw?: unknown
 * }} ElectionSoftError
 */

/**
 * @param {unknown} err
 * @returns {ElectionSoftError}
 */
export function classifyElectionError(err) {
    const status = Number(err && /** @type {{ status?: number }} */ (err).status) || 0;
    const codeRaw = String((err && /** @type {{ code?: string }} */ (err).code) || '')
        .trim()
        .toLowerCase();
    const msg = String((err && /** @type {{ message?: string }} */ (err).message) || err || '');

    if (
        status === 401 ||
        codeRaw === 'auth_required' ||
        /nicht angemeldet|anmelden|unauthorized|login_required|interaction_required|consent_required/i.test(
            msg
        )
    ) {
        return {
            code: 'auth_required',
            message: 'Bitte mit dem Schul-Konto anmelden – die Wahl-API braucht Ihre Anmeldung.',
            status: status || 401,
            keepLocal: true,
            raw: err
        };
    }

    if (status === 403 || codeRaw === 'forbidden' || /keine berechtigung|forbidden|accessdenied/i.test(msg)) {
        return {
            code: 'forbidden',
            message: 'Kein Zugriff auf diese Aktion. Lokale bzw. bereits geladene Daten bleiben sichtbar.',
            status: status || 403,
            keepLocal: true,
            raw: err
        };
    }

    if (status === 404 || codeRaw === 'not_found' || /nicht gefunden/i.test(msg)) {
        return {
            code: 'not_found',
            message: 'Wahl nicht gefunden. Bitte den aktuellen Link bzw. QR-Code der Schule verwenden.',
            status: status || 404,
            keepLocal: true,
            raw: err
        };
    }

    if (
        codeRaw === 'not_open' ||
        /noch nicht eröffnet|fenster ist geschlossen|keine stimmabgabe möglich|derzeit keine stimmabgabe/i.test(
            msg
        )
    ) {
        return {
            code: 'not_open',
            message: msg.slice(0, 200) || 'Derzeit keine Stimmabgabe möglich.',
            status: status || 409,
            keepLocal: true,
            raw: err
        };
    }

    if (
        /failed to fetch|networkerror|load failed|offline|timeout|econnrefused|enotfound/i.test(msg)
    ) {
        return {
            code: 'network',
            message: 'Wahl-API gerade nicht erreichbar. Bereits geladene Daten bleiben erhalten.',
            status: status || 0,
            keepLocal: true,
            raw: err
        };
    }

    if (/baseurl|ms365_election_api|nicht gesetzt/i.test(msg)) {
        return {
            code: 'config',
            message: 'Wahl-API ist nicht konfiguriert. Demo lokal mit ?demo=1 möglich.',
            status: status || 0,
            keepLocal: true,
            raw: err
        };
    }

    return {
        code: 'unknown',
        message: msg.slice(0, 180) || 'Unbekannter Fehler – Daten wurden nicht verworfen.',
        status,
        keepLocal: true,
        raw: err
    };
}

/**
 * Kurztext für Badges / Toasts (ohne Roh-HTTP).
 * @param {unknown} err
 */
export function softElectionMessage(err) {
    return classifyElectionError(err).message;
}
