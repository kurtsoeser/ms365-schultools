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

function validateTeamsPayload(body) {
    if (!body || typeof body !== 'object') {
        return { error: 'JSON-Body erforderlich.' };
    }
    const tenantId = String(body.tenantId || '').trim();
    if (!tenantId) {
        return { error: 'tenantId ist erforderlich.' };
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
    return { tenantId, teams };
}

module.exports = { jsonResponse, corsPreflightResponse, validateTeamsPayload, CORS_HEADERS };
