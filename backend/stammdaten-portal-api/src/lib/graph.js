'use strict';

const TOKEN_URL = (tenant) => `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`;
const GRAPH = 'https://graph.microsoft.com/v1.0';

let cached = { token: '', exp: 0 };

function env(name, fallback = '') {
    const v = process.env[name];
    if (v === undefined || v === null || String(v).trim() === '') return fallback;
    return String(v).trim();
}

function graphConfigured() {
    return !!(
        env('STAMMDATEN_GRAPH_TENANT_ID') &&
        env('STAMMDATEN_GRAPH_CLIENT_ID') &&
        env('STAMMDATEN_GRAPH_CLIENT_SECRET') &&
        (env('STAMMDATEN_SP_SITE_ID') || env('STAMMDATEN_SP_SITE_URL'))
    );
}

async function getAppToken() {
    const now = Date.now();
    if (cached.token && cached.exp > now + 60_000) return cached.token;
    const tenant = env('STAMMDATEN_GRAPH_TENANT_ID');
    const body = new URLSearchParams({
        client_id: env('STAMMDATEN_GRAPH_CLIENT_ID'),
        client_secret: env('STAMMDATEN_GRAPH_CLIENT_SECRET'),
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials'
    });
    const res = await fetch(TOKEN_URL(tenant), {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok || !data.access_token) {
        throw Object.assign(new Error(data.error_description || 'Graph-Token fehlgeschlagen.'), {
            status: 502,
            code: 'graph_token'
        });
    }
    cached = {
        token: data.access_token,
        exp: now + Number(data.expires_in || 3600) * 1000
    };
    return cached.token;
}

async function graphJson(method, path, body) {
    const token = await getAppToken();
    const url = path.startsWith('http') ? path : `${GRAPH}${path.startsWith('/') ? '' : '/'}${path}`;
    const res = await fetch(url, {
        method,
        headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json',
            Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly'
        },
        body: body == null ? undefined : JSON.stringify(body)
    });
    if (res.status === 204) return null;
    const text = await res.text();
    let data = null;
    try {
        data = text ? JSON.parse(text) : null;
    } catch {
        data = { raw: text };
    }
    if (!res.ok) {
        const msg = (data && (data.error?.message || data.error_description)) || text || res.statusText;
        throw Object.assign(new Error(msg), { status: res.status, code: 'graph', data });
    }
    return data;
}

/**
 * @param {{ siteId?: string, siteUrl?: string }} [opts]
 */
async function resolveSiteId(opts = {}) {
    const direct = String(opts.siteId || env('STAMMDATEN_SP_SITE_ID') || '').trim();
    if (direct) return direct;
    const siteUrl = String(opts.siteUrl || env('STAMMDATEN_SP_SITE_URL') || '').trim();
    if (!siteUrl) {
        throw Object.assign(
            new Error('SharePoint-Site fehlt: bitte Site-URL in der Einrichtung oder STAMMDATEN_SP_SITE_URL setzen.'),
            { status: 400, code: 'site_missing' }
        );
    }
    const u = new URL(siteUrl);
    const host = u.hostname;
    const pathPart = u.pathname.replace(/\/$/, '') || '/';
    const data = await graphJson('GET', `/sites/${host}:${pathPart}`);
    return data.id;
}

module.exports = {
    env,
    graphConfigured,
    getAppToken,
    graphJson,
    resolveSiteId
};
