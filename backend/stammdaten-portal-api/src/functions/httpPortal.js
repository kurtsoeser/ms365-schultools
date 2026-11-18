'use strict';

const crypto = require('node:crypto');
const { app } = require('@azure/functions');
const {
    jsonResponse,
    corsPreflightResponse,
    readJsonBody,
    bearerTokenFromRequest
} = require('../lib/http-utils');
const { verifyPin, createSessionToken, readSessionToken, hashPin } = require('../lib/security');
const store = require('../lib/store');
const {
    STATUS,
    isParentEditLocked,
    canAdminTransition,
    applyGeprueftStamp,
    applyImportedStamp,
    normalizeStatus
} = require('../lib/status-transitions');

const MAX_ATTEMPTS = 5;
const LOCK_MINUTES = 15;

function adminKeyOk(request) {
    const expected = String(process.env.STAMMDATEN_ADMIN_KEY || '').trim();
    if (!expected) return false;
    const h =
        (request.headers &&
            (request.headers.get('x-admin-key') ||
                request.headers.get('x-functions-key') ||
                request.headers.get('X-Admin-Key'))) ||
        '';
    return String(h).trim() === expected;
}

function requireAdmin(request) {
    if (adminKeyOk(request)) return null;
    return jsonResponse(request, 401, { error: 'Admin-Key fehlt oder ungültig.' });
}

app.http('health', {
    methods: ['GET', 'OPTIONS'],
    authLevel: 'anonymous',
    route: 'health',
    handler: async (request) => {
        if (request.method === 'OPTIONS') return corsPreflightResponse(request);
        return jsonResponse(request, 200, {
            ok: true,
            service: 'stammdaten-portal-api',
            store: store.backendMode()
        });
    }
});

app.http('unlock', {
    methods: ['POST', 'OPTIONS'],
    authLevel: 'anonymous',
    route: 'unlock',
    handler: async (request) => {
        if (request.method === 'OPTIONS') return corsPreflightResponse(request);
        const body = await readJsonBody(request);
        const id = String(body.id || '').trim();
        const pin = String(body.pin || '').trim();
        if (!id || !pin) return jsonResponse(request, 401, { error: 'Ungültiger Zugang.' });

        const record = await store.getByGuid(id);
        if (!record) return jsonResponse(request, 401, { error: 'Ungültiger Zugang.' });

        const now = Date.now();
        if (record.lockedUntil && new Date(record.lockedUntil).getTime() > now) {
            return jsonResponse(request, 429, {
                error: 'Zugang vorübergehend gesperrt. Bitte später erneut versuchen.'
            });
        }

        const ok = await verifyPin(pin, record.accessPinHash);
        if (!ok) {
            const failed = Number(record.failedAttempts || 0) + 1;
            const next = { ...record, failedAttempts: failed, updatedAt: new Date().toISOString() };
            if (failed >= MAX_ATTEMPTS) {
                next.lockedUntil = new Date(now + LOCK_MINUTES * 60_000).toISOString();
                next.failedAttempts = 0;
            }
            await store.upsert(next);
            return jsonResponse(request, 401, { error: 'Ungültiger Zugang.' });
        }

        const unlocked = {
            ...record,
            failedAttempts: 0,
            lockedUntil: null,
            updatedAt: new Date().toISOString()
        };
        await store.upsert(unlocked);

        return jsonResponse(request, 200, {
            token: createSessionToken(id),
            record: publicView(unlocked)
        });
    }
});

app.http('load', {
    methods: ['GET', 'POST', 'OPTIONS'],
    authLevel: 'anonymous',
    route: 'load',
    handler: async (request) => {
        if (request.method === 'OPTIONS') return corsPreflightResponse(request);
        const session = readSessionToken(bearerTokenFromRequest(request));
        if (!session) return jsonResponse(request, 401, { error: 'Sitzung abgelaufen. Bitte erneut anmelden.' });
        const record = await store.getByGuid(session.guid);
        if (!record) return jsonResponse(request, 401, { error: 'Ungültiger Zugang.' });
        return jsonResponse(request, 200, { record: publicView(record) });
    }
});

app.http('save', {
    methods: ['POST', 'OPTIONS'],
    authLevel: 'anonymous',
    route: 'save',
    handler: async (request) => {
        if (request.method === 'OPTIONS') return corsPreflightResponse(request);
        const session = readSessionToken(bearerTokenFromRequest(request));
        if (!session) return jsonResponse(request, 401, { error: 'Sitzung abgelaufen. Bitte erneut anmelden.' });
        const record = await store.getByGuid(session.guid);
        if (!record) return jsonResponse(request, 401, { error: 'Ungültiger Zugang.' });

        if (isParentEditLocked(record.status)) {
            return jsonResponse(request, 409, {
                error:
                    'Dieser Datensatz wurde vom Sekretariat geprüft und kann nicht mehr geändert werden. Bei Korrekturen bitte die Schule kontaktieren.',
                code: 'locked_after_review',
                status: normalizeStatus(record.status)
            });
        }

        const body = await readJsonBody(request);
        const payload = body.payload && typeof body.payload === 'object' ? body.payload : {};
        const submit = !!body.submit;
        const next = {
            ...record,
            payload,
            status: submit ? STATUS.VOLLSTAENDIG : STATUS.IN_BEARBEITUNG,
            updatedAt: new Date().toISOString()
        };
        const s = payload.schueler || {};
        next.title = [s.familienname, s.vorname].filter(Boolean).join(', ') || record.title;
        await store.upsert(next);
        return jsonResponse(request, 200, { record: publicView(next) });
    }
});

/** Admin: Liste anlegen / Spalten ergänzen */
app.http('sspEnsureList', {
    methods: ['POST', 'OPTIONS'],
    authLevel: 'anonymous',
    route: 'ssp-ensure-list',
    handler: async (request) => {
        if (request.method === 'OPTIONS') return corsPreflightResponse(request);
        const denied = requireAdmin(request);
        if (denied) return denied;
        try {
            const body = await readJsonBody(request);
            const result = await store.ensureList({
                siteUrl: body.siteUrl || body.spSiteUrl || '',
                siteId: body.siteId || '',
                listName: body.listName || body.spListName || ''
            });
            return jsonResponse(request, 200, { ok: true, ...result });
        } catch (e) {
            return jsonResponse(request, e.status || 500, { error: e.message || 'ensure-list fehlgeschlagen.' });
        }
    }
});

/** Admin: Datensatz anlegen (PIN einmal zurück) */
app.http('sspEnroll', {
    methods: ['POST', 'OPTIONS'],
    authLevel: 'anonymous',
    route: 'ssp-enroll',
    handler: async (request) => {
        if (request.method === 'OPTIONS') return corsPreflightResponse(request);
        const denied = requireAdmin(request);
        if (denied) return denied;
        const body = await readJsonBody(request);
        const uniqueGuid = body.id || crypto.randomUUID();
        const pin = String(body.pin || '').trim() || String(100000 + crypto.randomInt(900000));
        const now = new Date().toISOString();
        const record = {
            uniqueGuid,
            accessPinHash: await hashPin(pin),
            status: 'Vorerfasst',
            title: body.title || 'Neue Anmeldung',
            payload: body.payload || {},
            failedAttempts: 0,
            lockedUntil: null,
            createdAt: now,
            updatedAt: now
        };
        await store.upsert(record);
        return jsonResponse(request, 200, { id: uniqueGuid, pin, record: publicView(record) });
    }
});

/** Admin: PIN neu */
app.http('sspResetPin', {
    methods: ['POST', 'OPTIONS'],
    authLevel: 'anonymous',
    route: 'ssp-reset-pin',
    handler: async (request) => {
        if (request.method === 'OPTIONS') return corsPreflightResponse(request);
        const denied = requireAdmin(request);
        if (denied) return denied;
        const body = await readJsonBody(request);
        const id = String(body.id || '').trim();
        if (!id) return jsonResponse(request, 400, { error: 'id fehlt.' });
        const record = await store.getByGuid(id);
        if (!record) return jsonResponse(request, 404, { error: 'Datensatz nicht gefunden.' });
        const pin = String(body.pin || '').trim() || String(100000 + crypto.randomInt(900000));
        const next = {
            ...record,
            accessPinHash: await hashPin(pin),
            failedAttempts: 0,
            lockedUntil: null,
            updatedAt: new Date().toISOString()
        };
        await store.upsert(next);
        return jsonResponse(request, 200, { id, pin, record: publicView(next) });
    }
});

/** Admin: Liste */
app.http('sspList', {
    methods: ['GET', 'OPTIONS'],
    authLevel: 'anonymous',
    route: 'ssp-list',
    handler: async (request) => {
        if (request.method === 'OPTIONS') return corsPreflightResponse(request);
        const denied = requireAdmin(request);
        if (denied) return denied;
        const rows = await store.listAll();
        return jsonResponse(request, 200, {
            records: rows.map(publicView)
        });
    }
});

/** Admin: Status setzen (z. B. Sekretariats-Stempel „Geprüft“ oder nach Export „Importiert“). */
app.http('sspSetStatus', {
    methods: ['POST', 'OPTIONS'],
    authLevel: 'anonymous',
    route: 'ssp-set-status',
    handler: async (request) => {
        if (request.method === 'OPTIONS') return corsPreflightResponse(request);
        const denied = requireAdmin(request);
        if (denied) return denied;
        const body = await readJsonBody(request);
        const ids = Array.isArray(body.ids) ? body.ids.map(String) : body.id ? [String(body.id)] : [];
        const status = String(body.status || '').trim();
        const actorEmail = String(body.actorEmail || body.reviewedBy || '').trim();
        const actorName = String(body.actorName || body.reviewedByName || '').trim();
        if (!ids.length) return jsonResponse(request, 400, { error: 'ids fehlen.' });
        if (!STATUS_LIST_HAS(status)) {
            return jsonResponse(request, 400, {
                error: 'Ungültiger Status.',
                allowed: Object.values(STATUS)
            });
        }
        const updated = [];
        const rejected = [];
        for (const id of ids) {
            const record = await store.getByGuid(id);
            if (!record) continue;
            const gate = canAdminTransition(record.status, status);
            if (!gate.ok) {
                rejected.push({ id, error: gate.message, from: gate.from, to: gate.to });
                continue;
            }
            if (gate.from === gate.to) {
                updated.push(publicView(record));
                continue;
            }
            let next = record;
            if (normalizeStatus(status) === STATUS.GEPRUEFT) {
                const stamped = applyGeprueftStamp(record, { actorEmail, actorName });
                if (!stamped.ok) {
                    rejected.push({ id, error: stamped.message });
                    continue;
                }
                next = stamped.record;
            } else if (normalizeStatus(status) === STATUS.IMPORTIERT) {
                const stamped = applyImportedStamp(record);
                if (!stamped.ok) {
                    rejected.push({ id, error: stamped.message });
                    continue;
                }
                next = stamped.record;
            } else {
                next = { ...record, status: normalizeStatus(status), updatedAt: new Date().toISOString() };
            }
            await store.upsert(next);
            updated.push(publicView(next));
        }
        if (!updated.length && rejected.length) {
            return jsonResponse(request, 409, {
                error: rejected[0].error || 'Statuswechsel nicht erlaubt.',
                rejected,
                updated
            });
        }
        return jsonResponse(request, 200, { updated, rejected });
    }
});

function STATUS_LIST_HAS(status) {
    return Object.values(STATUS).includes(String(status || '').trim());
}

function publicView(record) {
    const status = normalizeStatus(record.status);
    return {
        id: record.uniqueGuid,
        status,
        title: record.title || '',
        payload: record.payload || {},
        updatedAt: record.updatedAt || null,
        reviewedAt: record.reviewedAt || null,
        reviewedBy: record.reviewedBy || '',
        reviewedByName: record.reviewedByName || '',
        parentEditLocked: isParentEditLocked(status)
    };
}
