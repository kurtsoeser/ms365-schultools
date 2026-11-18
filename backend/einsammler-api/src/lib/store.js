'use strict';

/** @type {Map<string, object>} */
const g = globalThis;
if (!g.__einsammlerPacks) g.__einsammlerPacks = new Map();
if (!g.__einsammlerResponses) g.__einsammlerResponses = new Map();
if (!g.__krankmeldungen) g.__krankmeldungen = new Map();

function packs() {
    return g.__einsammlerPacks;
}

function responses() {
    return g.__einsammlerResponses;
}

/**
 * @param {object} pack
 */
function upsertPack(pack) {
    const aid = String(pack && pack.aid ? pack.aid : '').trim();
    if (!aid) throw new Error('aid fehlt');
    packs().set(aid, { ...pack });
    return packs().get(aid);
}

/**
 * @param {string} aid
 */
function getPack(aid) {
    return packs().get(String(aid || '').trim()) || null;
}

/**
 * @param {string} tokenOrAid
 */
function getPackByTokenPayload(payload) {
    if (!payload || !payload.aid) return null;
    const stored = getPack(payload.aid);
    return stored || payload;
}

/**
 * @param {object} response
 */
function saveResponse(response) {
    const aid = String(response && response.aid ? response.aid : '').trim();
    if (!aid) throw new Error('aid fehlt');
    responses().set(aid, { ...response, savedAt: new Date().toISOString() });
    const pack = getPack(aid);
    if (pack) {
        packs().set(aid, {
            ...pack,
            answered: true,
            answerStatus: response.status,
            answerKommentar: response.kommentar || '',
            antwortAm: response.antwortAm || new Date().toISOString()
        });
    }
    return responses().get(aid);
}

/**
 * @param {string} aid
 */
function getResponse(aid) {
    return responses().get(String(aid || '').trim()) || null;
}

function krankmeldungen() {
    return g.__krankmeldungen;
}

/**
 * @param {object} meldung
 */
function upsertKrankmeldung(meldung) {
    const id = String(meldung && meldung.id ? meldung.id : '').trim();
    if (!id) throw new Error('id fehlt');
    krankmeldungen().set(id, { ...meldung });
    return krankmeldungen().get(id);
}

function listKrankmeldungen() {
    return [...krankmeldungen().values()];
}

module.exports = {
    upsertPack,
    getPack,
    getPackByTokenPayload,
    saveResponse,
    getResponse,
    upsertKrankmeldung,
    listKrankmeldungen
};
