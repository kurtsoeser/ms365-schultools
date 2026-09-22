'use strict';

const CORS_HEADERS = {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PATCH, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization',
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
            'Access-Control-Allow-Methods': 'GET, POST, PATCH, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization',
            'Access-Control-Max-Age': '86400'
        }
    };
}

/**
 * @param {import('@azure/functions').HttpRequest} request
 * @returns {string}
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

module.exports = {
    jsonResponse,
    corsPreflightResponse,
    bearerTokenFromRequest,
    readJsonBody,
    CORS_HEADERS
};
