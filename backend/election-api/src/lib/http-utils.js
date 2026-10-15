'use strict';

const CORS_HEADERS = {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, PATCH, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, x-functions-key',
    'Cache-Control': 'no-store'
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
            'Access-Control-Allow-Methods': 'GET, POST, PUT, PATCH, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, x-functions-key',
            'Access-Control-Max-Age': '86400'
        }
    };
}

/**
 * @param {import('@azure/functions').HttpRequest} request
 */
function bearerTokenFromRequest(request) {
    const h =
        (request.headers && (request.headers.get('authorization') || request.headers.get('Authorization'))) ||
        '';
    const m = String(h).match(/^Bearer\s+(.+)$/i);
    return m ? m[1].trim() : '';
}

/**
 * @param {import('@azure/functions').HttpRequest} request
 */
async function readJsonBody(request) {
    try {
        const body = await request.json();
        return body && typeof body === 'object' ? body : {};
    } catch {
        return {};
    }
}

/**
 * @param {Error & { status?: number, code?: string, meta?: unknown }} e
 */
function publicErrorMessage(e) {
    const status = e && e.status && Number.isFinite(e.status) ? e.status : 500;
    const raw = String((e && e.message) || '').trim();
    if (status >= 500 && status !== 503) return 'Interner Fehler.';
    if (!raw) return 'Anfrage abgelehnt.';
    if (/AADSTS|client.?secret|Bearer\s|stack|at\s+\S+\s+\(/i.test(raw)) {
        return 'Anfrage abgelehnt.';
    }
    return raw;
}

/**
 * @param {import('@azure/functions').InvocationContext} context
 * @param {string} label
 * @param {Error & { status?: number, code?: string, meta?: unknown }} e
 */
function errorJsonResponse(context, label, e) {
    const status = e && e.status && Number.isFinite(e.status) ? e.status : 500;
    if (status >= 500) context.error(label, e);
    else context.error(label, status, (e && e.code) || '', (e && e.message) || '');
    return jsonResponse(status >= 400 && status < 600 ? status : 500, {
        error: publicErrorMessage(e),
        ...(e && e.code ? { code: e.code } : {}),
        ...(e && e.meta ? { meta: e.meta } : {})
    });
}

module.exports = {
    jsonResponse,
    corsPreflightResponse,
    bearerTokenFromRequest,
    readJsonBody,
    publicErrorMessage,
    errorJsonResponse,
    CORS_HEADERS
};
