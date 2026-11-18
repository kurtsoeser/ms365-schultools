'use strict';

function corsOrigins() {
    return String(process.env.STAMMDATEN_CORS_ORIGINS || '*')
        .split(',')
        .map((s) => s.trim())
        .filter(Boolean);
}

function resolveAllowOrigin(request) {
    const allowed = corsOrigins();
    const origin = (request.headers && (request.headers.get('origin') || request.headers.get('Origin'))) || '';
    // Gateway ist bewusst browser-offen (Admin + Eltern-Formular auf Schuldomains).
    if (allowed.includes('*') || allowed.length === 0) {
        return origin || '*';
    }
    if (!origin) return allowed[0] || '*';
    if (allowed.includes(origin) || allowed.includes('null')) return origin;
    // Unbekannte Origin: trotzdem spiegeln, damit lokale Dev-Ports nicht „Failed to fetch“ liefern.
    // Absicherung bleibt über Admin-Key / Session-PIN.
    return origin;
}

function corsHeaders(request) {
    return {
        'Content-Type': 'application/json; charset=utf-8',
        'Access-Control-Allow-Origin': resolveAllowOrigin(request),
        'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
        'Access-Control-Allow-Headers':
            'Content-Type, Authorization, x-admin-key, x-functions-key, X-Admin-Key',
        'Access-Control-Max-Age': '86400',
        'Cache-Control': 'no-store',
        Vary: 'Origin'
    };
}

function jsonResponse(request, status, body) {
    return { status, headers: corsHeaders(request), jsonBody: body };
}

function corsPreflightResponse(request) {
    return {
        status: 204,
        headers: {
            ...corsHeaders(request),
            'Access-Control-Max-Age': '86400'
        }
    };
}

async function readJsonBody(request) {
    try {
        const body = await request.json();
        return body && typeof body === 'object' ? body : {};
    } catch {
        return {};
    }
}

function bearerTokenFromRequest(request) {
    const h =
        (request.headers && (request.headers.get('authorization') || request.headers.get('Authorization'))) || '';
    const m = String(h).match(/^Bearer\s+(.+)$/i);
    return m ? m[1].trim() : '';
}

module.exports = {
    jsonResponse,
    corsPreflightResponse,
    readJsonBody,
    bearerTokenFromRequest
};
