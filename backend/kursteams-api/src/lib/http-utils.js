'use strict';

const CORS_HEADERS = {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization'
};

function jsonResponse(status, body) {
    return {
        status,
        headers: CORS_HEADERS,
        jsonBody: body
    };
}

function corsPreflightResponse() {
    return {
        status: 204,
        headers: {
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization',
            'Access-Control-Max-Age': '86400'
        }
    };
}

function bearerTokenFromRequest(request) {
    const h =
        (request.headers &&
            (request.headers.get('authorization') || request.headers.get('Authorization'))) ||
        '';
    const m = String(h).match(/^Bearer\s+(.+)$/i);
    return m ? m[1].trim() : '';
}

function validateTeamsPayload(body) {
    if (!body || typeof body !== 'object') {
        return { error: 'JSON-Body erforderlich.' };
    }
    if (!Array.isArray(body.teams) || !body.teams.length) {
        return { error: 'teams (Array, mindestens 1 Eintrag) ist erforderlich.' };
    }
    const teams = [];
    for (let i = 0; i < body.teams.length; i++) {
        const t = body.teams[i] || {};
        const teamName = String(t.teamName || '').trim();
        const gruppenmail = String(t.gruppenmail || '').trim();
        const besitzer = String(t.besitzer || '').trim();
        if (!teamName || !gruppenmail || !besitzer) {
            return {
                error:
                    'Team #' +
                    (i + 1) +
                    ': teamName, gruppenmail und besitzer sind erforderlich.'
            };
        }
        teams.push({ teamName, gruppenmail, besitzer });
    }
    const mailDomain = String(body.mailDomain || '')
        .trim()
        .replace(/^@+/, '');
    return { teams, mailDomain };
}

/**
 * @param {import('@azure/functions').InvocationContext} context
 * @param {string} label
 * @param {Error & { status?: number, cause?: unknown }} e
 */
function errorResponse(context, label, e) {
    const status = e && e.status && Number.isFinite(e.status) ? e.status : 500;
    if (status >= 500) context.error(label, e);
    else context.error(label, status, e && e.cause ? e.cause : '');
    const safe =
        status >= 500
            ? 'Kursteams-Backend ist fehlgeschlagen.'
            : (e && e.message) || 'Anfrage abgelehnt.';
    return jsonResponse(status >= 400 && status < 600 ? status : 500, { error: safe });
}

module.exports = {
    jsonResponse,
    corsPreflightResponse,
    validateTeamsPayload,
    bearerTokenFromRequest,
    errorResponse,
    CORS_HEADERS
};
