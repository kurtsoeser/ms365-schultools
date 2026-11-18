#!/usr/bin/env node
/**
 * Lokaler Nextcloud-WebDAV-Helfer (nur 127.0.0.1).
 * Liest .env.local, hängt Auth an, setzt CORS – damit die Browser-App browsen kann.
 *
 * Start: Start-Helfer.cmd  oder  node proxy-server.mjs
 */
import http from 'node:http';
import { getConfig, assertNextcloudAuth, loadDefaultEnv } from './lib/config.mjs';
import { davFilesRoot, davUrl, basicAuthHeader, propfind } from './lib/webdav.mjs';

const PORT = Number(process.env.NC_HELPER_PORT || 3791) || 3791;
const HOST = '127.0.0.1';

loadDefaultEnv();
const cfg = getConfig();

function sendJson(res, status, obj) {
    const body = JSON.stringify(obj, null, 2);
    res.writeHead(status, {
        'Content-Type': 'application/json; charset=utf-8',
        'Content-Length': Buffer.byteLength(body),
        ...corsHeaders()
    });
    res.end(body);
}

function corsHeaders(req) {
    const origin = (req && req.headers && req.headers.origin) || '*';
    return {
        'Access-Control-Allow-Origin': origin === 'null' ? '*' : origin,
        'Access-Control-Allow-Methods': 'GET, HEAD, PROPFIND, OPTIONS',
        'Access-Control-Allow-Headers': 'Authorization, Content-Type, Depth',
        'Access-Control-Max-Age': '86400',
        Vary: 'Origin'
    };
}

/**
 * /dav/... → Relativpfad unter dem Nextcloud-Benutzerroot
 * @param {string} urlPath
 */
function relFromUrlPath(urlPath) {
    const raw = String(urlPath || '');
    const without = raw.replace(/^\/dav\/?/i, '');
    try {
        return decodeURIComponent(without)
            .replace(/^\/+|\/+$/g, '')
            .split('/')
            .filter(Boolean)
            .map((p) => decodeURIComponent(p))
            .join('/');
    } catch {
        return without.replace(/^\/+|\/+$/g, '');
    }
}

async function handleHealth(_req, res) {
    try {
        assertNextcloudAuth(cfg);
        await propfind({
            baseUrl: cfg.nextcloudBaseUrl,
            user: cfg.nextcloudUser,
            appPassword: cfg.nextcloudAppPassword,
            relPath: '',
            depth: '0'
        });
        sendJson(res, 200, {
            ok: true,
            helper: 'nextcloud-teams-migrate',
            nextcloudUrl: cfg.nextcloudBaseUrl,
            user: cfg.nextcloudUser,
            webdavRoot: davFilesRoot(cfg.nextcloudBaseUrl, cfg.nextcloudUser),
            message: 'Helfer bereit – Nextcloud-Login OK.'
        });
    } catch (e) {
        sendJson(res, 503, {
            ok: false,
            helper: 'nextcloud-teams-migrate',
            nextcloudUrl: cfg.nextcloudBaseUrl,
            user: cfg.nextcloudUser,
            error: e && e.message ? e.message : String(e),
            hint:
                'Prüfe scripts/nextcloud-teams-migrate/.env.local (NEXTCLOUD_USER + App-Passwort). Bei 401: App-Passwort neu erzeugen.'
        });
    }
}

/**
 * @param {import('node:http').IncomingMessage} req
 * @param {import('node:http').ServerResponse} res
 * @param {string} method
 * @param {string} relPath
 */
async function proxyDav(req, res, method, relPath) {
    assertNextcloudAuth(cfg);
    const root = davFilesRoot(cfg.nextcloudBaseUrl, cfg.nextcloudUser);
    const target = davUrl(root, relPath);
    /** @type {Record<string, string>} */
    const headers = {
        Authorization: basicAuthHeader(cfg.nextcloudUser, cfg.nextcloudAppPassword)
    };
    if (method === 'PROPFIND') {
        headers.Depth = String(req.headers.depth || '1');
        headers['Content-Type'] = 'application/xml; charset=utf-8';
    }

    const chunks = [];
    for await (const c of req) chunks.push(c);
    const bodyBuf = Buffer.concat(chunks);

    const upstream = await fetch(target, {
        method,
        headers,
        body: method === 'PROPFIND' ? (bodyBuf.length ? bodyBuf : undefined) : undefined
    });

    const outHeaders = {
        ...corsHeaders(req),
        'Content-Type': upstream.headers.get('content-type') || 'application/octet-stream'
    };
    const len = upstream.headers.get('content-length');
    if (len) outHeaders['Content-Length'] = len;

    res.writeHead(upstream.status, outHeaders);
    if (!upstream.body) {
        res.end();
        return;
    }
    const reader = upstream.body.getReader();
    while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        res.write(Buffer.from(value));
    }
    res.end();
}

function readBodyXmlDefault(req) {
    // PROPFIND ohne Body: Standard-Propfind
    return (
        '<?xml version="1.0"?>' +
        '<d:propfind xmlns:d="DAV:">' +
        '<d:prop><d:resourcetype/><d:getcontentlength/><d:getcontenttype/><d:displayname/></d:prop>' +
        '</d:propfind>'
    );
}

const server = http.createServer(async (req, res) => {
    try {
        const host = req.headers.host || `${HOST}:${PORT}`;
        const u = new URL(req.url || '/', `http://${host}`);
        const method = (req.method || 'GET').toUpperCase();

        if (method === 'OPTIONS') {
            res.writeHead(204, corsHeaders(req));
            res.end();
            return;
        }

        if (u.pathname === '/' || u.pathname === '/health') {
            if (method !== 'GET' && method !== 'HEAD') {
                sendJson(res, 405, { ok: false, error: 'Nur GET' });
                return;
            }
            await handleHealth(req, res);
            return;
        }

        if (u.pathname === '/config') {
            sendJson(res, 200, {
                ok: true,
                nextcloudUrl: cfg.nextcloudBaseUrl,
                user: cfg.nextcloudUser,
                hasPassword: !!cfg.nextcloudAppPassword,
                destRootFolder: cfg.destRootFolder
            });
            return;
        }

        if (u.pathname.startsWith('/dav')) {
            const rel = relFromUrlPath(u.pathname);
            if (method === 'PROPFIND') {
                // Wenn kein Body: Default setzen über Tee – einfach erneut lesen
                const chunks = [];
                for await (const c of req) chunks.push(c);
                let body = Buffer.concat(chunks);
                if (!body.length) body = Buffer.from(readBodyXmlDefault(req), 'utf8');

                assertNextcloudAuth(cfg);
                const root = davFilesRoot(cfg.nextcloudBaseUrl, cfg.nextcloudUser);
                const target = davUrl(root, rel);
                const upstream = await fetch(target, {
                    method: 'PROPFIND',
                    headers: {
                        Authorization: basicAuthHeader(cfg.nextcloudUser, cfg.nextcloudAppPassword),
                        Depth: String(req.headers.depth || '1'),
                        'Content-Type': 'application/xml; charset=utf-8'
                    },
                    body
                });
                const text = await upstream.text();
                res.writeHead(upstream.status, {
                    ...corsHeaders(req),
                    'Content-Type': upstream.headers.get('content-type') || 'application/xml'
                });
                res.end(text);
                return;
            }
            if (method === 'GET' || method === 'HEAD') {
                await proxyDav(req, res, method, rel);
                return;
            }
            sendJson(res, 405, { ok: false, error: 'Methode nicht erlaubt: ' + method });
            return;
        }

        sendJson(res, 404, {
            ok: false,
            error: 'Unbekannt',
            routes: ['GET /health', 'PROPFIND|GET /dav/<pfad>']
        });
    } catch (e) {
        sendJson(res, 500, {
            ok: false,
            error: e && e.message ? e.message : String(e)
        });
    }
});

server.on('error', (err) => {
    if (err && err.code === 'EADDRINUSE') {
        console.error('');
        console.error(`Port ${PORT} ist schon belegt – der Helfer läuft vermutlich bereits.`);
        console.error('Entweder das andere Helfer-Fenster nutzen, oder dort schließen und erneut starten.');
        console.error(`Schnelltest im Browser: http://${HOST}:${PORT}/health`);
        console.error('');
        process.exit(1);
    }
    console.error('Helfer-Fehler:', err && err.message ? err.message : err);
    process.exit(1);
});

server.listen(PORT, HOST, () => {
    console.log('');
    console.log('=== Nextcloud → Teams Helfer ===');
    console.log(`Läuft nur lokal:  http://${HOST}:${PORT}/health`);
    console.log(`Nextcloud:        ${cfg.nextcloudBaseUrl || '(fehlt)'}`);
    console.log(`Benutzer:         ${cfg.nextcloudUser || '(fehlt)'}`);
    console.log(`Passwort gesetzt: ${cfg.nextcloudAppPassword ? 'ja' : 'NEIN – .env.local prüfen'}`);
    console.log('');
    console.log('Fenster offen lassen. In der App: Reiter „Lokaler Helfer“ → Verbinden.');
    console.log('Beenden: Strg+C oder Fenster schließen.');
    console.log('');
    // Sofort-Check (nicht blockierend für den Server)
    propfind({
        baseUrl: cfg.nextcloudBaseUrl,
        user: cfg.nextcloudUser,
        appPassword: cfg.nextcloudAppPassword,
        relPath: '',
        depth: '0'
    })
        .then(() => console.log('Nextcloud-Login: OK'))
        .catch((e) => {
            console.log('Nextcloud-Login: FEHLER');
            console.log(' ', e && e.message ? e.message : e);
            console.log(' → App-Passwort in .env.local neu setzen, dann Helfer neu starten.');
            console.log('');
        });
});
