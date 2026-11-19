'use strict';

const { app } = require('@azure/functions');
const { jsonResponse, corsPreflightResponse, readJsonBody } = require('../lib/http-utils');
const store = require('../lib/store');
const { decodePublicToken, publicApiEnabled } = require('../lib/token');
const { createRateLimiter, clientKey } = require('../lib/rate-limit');

const krankmeldungLimiter = createRateLimiter({ windowMs: 60_000, max: 10 });

function adminKeyOk(request) {
    const expected = String(process.env.EINSAMMLER_ADMIN_KEY || '').trim();
    if (!expected) return false;
    const h =
        (request.headers &&
            (request.headers.get('x-admin-key') ||
                request.headers.get('x-functions-key') ||
                request.headers.get('X-Admin-Key'))) ||
        '';
    return String(h).trim() === expected;
}

function remoteDisabledResponse(request) {
    return jsonResponse(request, 503, {
        error: 'Remote-API abgeschaltet. EINSAMMLER_LINK_SECRET setzen und signierte Elternlinks verwenden.',
        code: 'parent_api_disabled'
    });
}

app.http('einsammlerHealth', {
    methods: ['GET', 'OPTIONS'],
    authLevel: 'anonymous',
    route: 'einsammler/health',
    handler: async (request) => {
        if (request.method === 'OPTIONS') return corsPreflightResponse(request);
        return jsonResponse(request, 200, {
            ok: true,
            service: 'einsammler-api',
            store: 'memory',
            publicApi: publicApiEnabled()
        });
    }
});

app.http('einsammlerPublish', {
    methods: ['POST', 'OPTIONS'],
    authLevel: 'anonymous',
    route: 'einsammler/publish',
    handler: async (request) => {
        if (request.method === 'OPTIONS') return corsPreflightResponse(request);
        if (!publicApiEnabled()) return remoteDisabledResponse(request);
        if (!adminKeyOk(request)) {
            return jsonResponse(request, 401, { error: 'Admin-Key fehlt oder ungültig.' });
        }
        const body = await readJsonBody(request);
        const packs = Array.isArray(body.packs) ? body.packs : [];
        let n = 0;
        for (const p of packs) {
            if (!p || !p.aid) continue;
            store.upsertPack(p);
            n++;
        }
        return jsonResponse(request, 200, { ok: true, published: n });
    }
});

app.http('einsammlerView', {
    methods: ['GET', 'OPTIONS'],
    authLevel: 'anonymous',
    route: 'einsammler/view',
    handler: async (request) => {
        if (request.method === 'OPTIONS') return corsPreflightResponse(request);
        if (!publicApiEnabled()) return remoteDisabledResponse(request);
        const url = new URL(request.url);
        const token = String(url.searchParams.get('p') || url.searchParams.get('token') || '').trim();
        const decoded = decodePublicToken(token);
        if (!decoded.ok) return jsonResponse(request, 400, { error: decoded.error });
        const pack = store.getPackByTokenPayload(decoded.payload);
        if (pack && pack.answered) {
            return jsonResponse(request, 409, {
                error: 'Bereits beantwortet.',
                payload: pack,
                answered: true
            });
        }
        return jsonResponse(request, 200, { payload: pack || decoded.payload });
    }
});

app.http('einsammlerRespond', {
    methods: ['POST', 'OPTIONS'],
    authLevel: 'anonymous',
    route: 'einsammler/respond',
    handler: async (request) => {
        if (request.method === 'OPTIONS') return corsPreflightResponse(request);
        if (!publicApiEnabled()) return remoteDisabledResponse(request);
        const body = await readJsonBody(request);
        const token = String(body.token || '').trim();
        const status = String(body.status || '')
            .trim()
            .toLowerCase();
        const kommentar = String(body.kommentar || '').trim();
        const decoded = decodePublicToken(token);
        if (!decoded.ok) return jsonResponse(request, 400, { error: decoded.error });
        const allowed = ['zugestimmt', 'abgelehnt', 'gelesen', 'abgegeben'];
        if (!allowed.includes(status)) {
            return jsonResponse(request, 400, { error: 'Status ungültig.' });
        }
        const existing = store.getResponse(decoded.payload.aid);
        if (existing) {
            return jsonResponse(request, 409, { error: 'Bereits beantwortet.', response: existing });
        }
        const anhangName = String(body.anhangName || '').trim();
        const anhangMime = String(body.anhangMime || '').trim();
        const anhangSize = Number(body.anhangSize) || 0;
        const anhangDataUrl = String(body.anhangDataUrl || '');
        // ~5.5MB JSON-Grenze grob einhalten
        if (anhangDataUrl && anhangDataUrl.length > 5.5e6) {
            return jsonResponse(request, 413, { error: 'Anhang zu groß für API.' });
        }
        const response = store.saveResponse({
            eid: decoded.payload.eid,
            aid: decoded.payload.aid,
            status,
            kommentar,
            schuelerName: decoded.payload.schuelerName || '',
            antwortAm: new Date().toISOString(),
            anhangName,
            anhangMime,
            anhangSize,
            anhangDataUrl: anhangDataUrl || undefined
        });
        // ensure pack exists for later admin poll
        store.upsertPack({ ...decoded.payload, answered: true, answerStatus: status });
        return jsonResponse(request, 200, {
            ok: true,
            response: {
                ...response,
                anhangDataUrl: anhangDataUrl ? '[stored]' : ''
            }
        });
    }
});

app.http('krankmeldungHealth', {
    methods: ['GET', 'OPTIONS'],
    authLevel: 'anonymous',
    route: 'krankmeldung/health',
    handler: async (request) => {
        if (request.method === 'OPTIONS') return corsPreflightResponse(request);
        return jsonResponse(request, 200, {
            ok: true,
            service: 'krankmeldung',
            store: 'memory',
            publicApi: publicApiEnabled()
        });
    }
});

app.http('krankmeldungSubmit', {
    methods: ['POST', 'OPTIONS'],
    authLevel: 'anonymous',
    route: 'krankmeldung',
    handler: async (request) => {
        if (request.method === 'OPTIONS') return corsPreflightResponse(request);
        // Anonyme Krankmeldung nur mit freigeschalteter API + Rate-Limit (R2/a5)
        if (!publicApiEnabled()) return remoteDisabledResponse(request);
        const limited = krankmeldungLimiter.check(clientKey(request));
        if (!limited.ok) {
            return jsonResponse(request, 429, {
                error: 'Zu viele Meldungen. Bitte später erneut versuchen.',
                retryAfterSec: limited.retryAfterSec
            });
        }
        const body = await readJsonBody(request);
        const m = body.meldung || body;
        const schuelerName = String(m.schuelerName || '').trim();
        const klasseCode = String(m.klasseCode || '').trim();
        const von = String(m.von || '').slice(0, 10);
        const bis = String(m.bis || von).slice(0, 10);
        if (!schuelerName || !klasseCode || !von) {
            return jsonResponse(request, 400, { error: 'Schülername, Klasse und Von-Datum sind Pflicht.' });
        }
        const id =
            String(m.id || '').trim() ||
            'km-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
        const saved = store.upsertKrankmeldung({
            id,
            schuelerName,
            klasseCode,
            von,
            bis,
            kategorie: String(m.kategorie || 'krankheit').toLowerCase(),
            gemeldetAm: new Date().toISOString(),
            gemeldetVonEmail: String(m.gemeldetVonEmail || '').trim().toLowerCase(),
            gemeldetVonName: String(m.gemeldetVonName || '').trim(),
            status: 'neu',
            kenntnisLehrerEmail: '',
            notiz: String(m.notiz || '').trim(),
            quelle: 'public'
        });
        return jsonResponse(request, 200, { ok: true, meldung: saved });
    }
});

app.http('krankmeldungList', {
    methods: ['GET', 'OPTIONS'],
    authLevel: 'anonymous',
    route: 'krankmeldung/list',
    handler: async (request) => {
        if (request.method === 'OPTIONS') return corsPreflightResponse(request);
        if (!publicApiEnabled()) return remoteDisabledResponse(request);
        if (!adminKeyOk(request)) {
            return jsonResponse(request, 401, { error: 'Admin-Key fehlt oder ungültig.' });
        }
        return jsonResponse(request, 200, { meldungen: store.listKrankmeldungen() });
    }
});
