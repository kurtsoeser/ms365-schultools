'use strict';

function corsOrigins() {
    return String(process.env.EINSAMMLER_CORS_ORIGINS || '*')
        .split(',')
        .map((s) => s.trim())
        .filter(Boolean);
}

function corsHeaders(request) {
    const origins = corsOrigins();
    const reqOrigin = (request.headers && request.headers.get('origin')) || '';
    const allow =
        origins.includes('*') || !reqOrigin
            ? origins.includes('*')
                ? '*'
                : origins[0] || '*'
            : origins.includes(reqOrigin)
              ? reqOrigin
              : origins[0] || '*';
    return {
        'Access-Control-Allow-Origin': allow,
        'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type,x-admin-key,x-functions-key',
        'Access-Control-Max-Age': '86400'
    };
}

function jsonResponse(request, status, body) {
    return {
        status,
        headers: {
            'Content-Type': 'application/json; charset=utf-8',
            ...corsHeaders(request)
        },
        body: JSON.stringify(body)
    };
}

function corsPreflightResponse(request) {
    return {
        status: 204,
        headers: corsHeaders(request)
    };
}

async function readJsonBody(request) {
    try {
        const data = await request.json();
        return data && typeof data === 'object' ? data : {};
    } catch {
        return {};
    }
}

module.exports = {
    jsonResponse,
    corsPreflightResponse,
    readJsonBody
};
