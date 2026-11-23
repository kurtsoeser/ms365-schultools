'use strict';

const { app } = require('@azure/functions');
const {
    jsonResponse,
    corsPreflightResponse,
    bearerTokenFromRequest,
    readJsonBody,
    errorJsonResponse
} = require('../lib/http-utils');
const { requireOperatorCaller, requireCatalogReader } = require('../lib/require-operator');
const { readReleaseNotes, writeReleaseNotes } = require('../lib/sharepoint-catalog');
const { normalizeReleaseNotes, notesForView } = require('../lib/release-notes-payload');

app.http('httpReleaseNotesOptions', {
    methods: ['OPTIONS'],
    authLevel: 'anonymous',
    route: 'license/release-notes',
    handler: async () => corsPreflightResponse()
});

app.http('httpReleaseNotesGet', {
    methods: ['GET'],
    authLevel: 'anonymous',
    route: 'license/release-notes',
    handler: async (request, context) => {
        try {
            const wantPublic = String(request.query.get('view') || '').toLowerCase() === 'public';
            const token = bearerTokenFromRequest(request);
            if (!wantPublic && token) {
                await requireCatalogReader(token);
            }
            const stored = await readReleaseNotes();
            const view = !wantPublic && token ? 'full' : 'public';
            return jsonResponse(200, {
                ok: true,
                missing: !!stored.missing,
                updatedAt: stored.updatedAt || '',
                view: view,
                notes: notesForView(stored.notes, view)
            });
        } catch (e) {
            return errorJsonResponse(context, 'license/release-notes GET:', e, { notes: [], missing: true });
        }
    }
});

app.http('httpReleaseNotesPut', {
    methods: ['PUT'],
    authLevel: 'anonymous',
    route: 'license/release-notes',
    handler: async (request, context) => {
        try {
            const caller = await requireOperatorCaller(bearerTokenFromRequest(request));
            const body = await readJsonBody(request);
            const notes = normalizeReleaseNotes(body.notes || body);
            if (!notes.length) {
                const err = new Error('Mindestens ein Neuigkeiten-Eintrag mit Titel ist nötig.');
                err.status = 400;
                throw err;
            }
            const stored = await writeReleaseNotes(notes, caller.upn || '');
            return jsonResponse(200, {
                ok: true,
                count: stored.notes.length,
                updatedAt: stored.updatedAt,
                publicCount: notesForView(stored.notes, 'public').length
            });
        } catch (e) {
            return errorJsonResponse(context, 'license/release-notes PUT:', e);
        }
    }
});
