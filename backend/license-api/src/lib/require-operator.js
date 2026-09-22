'use strict';

const { getConfig } = require('./config');
const { validateCallerToken } = require('./validate-token');

/**
 * @param {string} token
 * @returns {Promise<{ tid: string, oid: string, upn: string, name: string }>}
 */
async function requireOperatorCaller(token) {
    const cfg = getConfig();
    const caller = await validateCallerToken(token, cfg.tokenAudiences);
    const upn = String(caller.upn || '')
        .trim()
        .toLowerCase();
    if (!upn || !cfg.operatorUpns.includes(upn)) {
        const err = new Error('Nur Betreiber-Konten dürfen Lizenzen verwalten.');
        err.status = 403;
        throw err;
    }
    return caller;
}

module.exports = { requireOperatorCaller };
