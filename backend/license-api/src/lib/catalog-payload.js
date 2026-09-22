'use strict';

const KIND = 'ms365-kursteam-templates';
const VERSION = 3;
const MAX_TEMPLATES = 400;
const MAX_CHANNELS = 100;
const MAX_SCHOOL_FORMS = 40;

/**
 * @param {string} message
 * @param {number} [status]
 */
function httpError(message, status) {
    const err = new Error(message);
    err.status = status || 400;
    return err;
}

/**
 * @param {unknown} value
 * @param {number} max
 */
function clip(value, max) {
    return String(value == null ? '' : value)
        .trim()
        .slice(0, max);
}

/**
 * @param {unknown} raw
 * @param {number} index
 * @param {boolean} strict
 * @returns {object|null}
 */
function sanitizeTemplate(raw, index, strict) {
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) {
        if (strict) throw httpError('Vorlage ' + (index + 1) + ' ist kein Objekt.');
        return null;
    }
    const name = clip(raw.name, 200);
    if (!name) {
        if (strict) throw httpError('Vorlage ' + (index + 1) + ' hat keinen Namen.');
        return null;
    }
    const channelsIn = Array.isArray(raw.channels) ? raw.channels : [];
    if (channelsIn.length > MAX_CHANNELS) {
        throw httpError('Vorlage „' + name + '“ hat zu viele Kanäle (max. ' + MAX_CHANNELS + ').');
    }
    const channels = [];
    for (let i = 0; i < channelsIn.length; i++) {
        const ch = channelsIn[i];
        const displayName = clip(ch && (ch.displayName || ch.name), 250);
        if (!displayName) continue;
        const id = clip(ch && ch.id, 80) || 'ch-' + (i + 1);
        channels.push({ id, displayName });
    }
    return {
        id: clip(raw.id, 80) || 'tpl-' + (index + 1),
        name,
        schoolForm: clip(raw.schoolForm, 80),
        subjectCode: clip(raw.subjectCode, 40),
        schulstufe: clip(raw.schulstufe, 20),
        semester: clip(raw.semester, 8),
        description: clip(raw.description, 2000),
        channels,
        materialsPath: clip(raw.materialsPath, 260),
        updatedAt: clip(raw.updatedAt, 40)
    };
}

/**
 * @param {unknown} list
 * @param {boolean} strict
 */
function sanitizeTemplates(list, strict) {
    if (!Array.isArray(list)) {
        if (strict) throw httpError('templates[] fehlt.');
        return [];
    }
    if (list.length > MAX_TEMPLATES) {
        throw httpError('Zu viele Vorlagen (max. ' + MAX_TEMPLATES + ').');
    }
    const out = [];
    const seen = new Set();
    for (let i = 0; i < list.length; i++) {
        const tpl = sanitizeTemplate(list[i], i, strict);
        if (!tpl) continue;
        if (seen.has(tpl.id)) {
            if (strict) throw httpError('Doppelte Vorlagen-ID: ' + tpl.id);
            continue;
        }
        seen.add(tpl.id);
        out.push(tpl);
    }
    return out;
}

/**
 * @param {unknown} list
 */
function sanitizeSchoolForms(list) {
    if (!Array.isArray(list)) return [];
    const out = [];
    const seen = new Set();
    for (let i = 0; i < list.length && out.length < MAX_SCHOOL_FORMS; i++) {
        const name = clip(list[i], 80);
        if (!name) continue;
        const key = name.toLowerCase();
        if (seen.has(key)) continue;
        seen.add(key);
        out.push(name);
    }
    return out;
}

/**
 * @param {unknown} body
 * @param {{ updatedBy?: string, strict?: boolean, updatedAt?: string }} [meta]
 */
function buildStoredCatalog(body, meta) {
    const strict = !meta || meta.strict !== false;
    const src = body && typeof body === 'object' && !Array.isArray(body) ? body : {};
    if (strict && (!body || typeof body !== 'object' || Array.isArray(body))) {
        throw httpError('Katalog-JSON fehlt.');
    }
    const templates = sanitizeTemplates(src.templates, strict);
    return {
        kind: KIND,
        version: VERSION,
        updatedAt: (meta && meta.updatedAt) || new Date().toISOString(),
        updatedBy: clip(meta && meta.updatedBy, 200),
        schoolForms: sanitizeSchoolForms(src.schoolForms),
        templates
    };
}

/**
 * @param {{ catalogLibraryName?: string, catalogKursteamPath?: string, siteWebUrl?: string }} cfg
 * @param {{ missing?: boolean, message?: string }} [extra]
 */
function emptyCatalog(cfg, extra) {
    return {
        kind: KIND,
        version: VERSION,
        updatedAt: null,
        updatedBy: '',
        schoolForms: [],
        templates: [],
        missing: !extra || extra.missing !== false,
        message: (extra && extra.message) || '',
        library: (cfg && cfg.catalogLibraryName) || '',
        path: (cfg && cfg.catalogKursteamPath) || '',
        siteWebUrl: (cfg && cfg.siteWebUrl) || '',
        webUrl: ''
    };
}

module.exports = {
    KIND,
    VERSION,
    buildStoredCatalog,
    emptyCatalog,
    sanitizeTemplates
};
