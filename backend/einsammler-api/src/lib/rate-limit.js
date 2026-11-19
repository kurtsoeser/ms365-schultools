'use strict';

/**
 * Einfaches In-Memory Sliding-Window Rate-Limit (pro Function-Instanz).
 * @param {{ windowMs?: number, max?: number }} [opts]
 */
function createRateLimiter(opts = {}) {
    const windowMs = Math.max(1000, Number(opts.windowMs) || 60_000);
    const max = Math.max(1, Number(opts.max) || 20);
    /** @type {Map<string, number[]>} */
    const hits = new Map();

    /**
     * @param {string} key
     * @returns {{ ok: boolean, retryAfterSec?: number }}
     */
    function check(key) {
        const id = String(key || 'unknown');
        const now = Date.now();
        const cut = now - windowMs;
        const prev = (hits.get(id) || []).filter((t) => t > cut);
        if (prev.length >= max) {
            hits.set(id, prev);
            const retryAfterSec = Math.max(1, Math.ceil((prev[0] + windowMs - now) / 1000));
            return { ok: false, retryAfterSec };
        }
        prev.push(now);
        hits.set(id, prev);
        return { ok: true };
    }

    return { check, windowMs, max };
}

/**
 * @param {import('@azure/functions').HttpRequest} request
 */
function clientKey(request) {
    const h = request.headers;
    const xf =
        (h && (h.get('x-forwarded-for') || h.get('X-Forwarded-For'))) ||
        '';
    if (xf) return String(xf).split(',')[0].trim();
    return String((h && (h.get('x-real-ip') || h.get('X-Real-IP'))) || 'local');
}

module.exports = { createRateLimiter, clientKey };
