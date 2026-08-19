/**
 * Lokale Override-Konfiguration (ohne Server).
 * - Admin kann damit User-PINs und aktiv/deaktiviert für die PIN-Sperre setzen.
 * - Gilt nur für diesen Browser/Profil (localStorage), weil es keinen Server gibt.
 */

export const ACCESS_OVERRIDE_KEY = 'ms365-schooltool-access-override-v1';

function safeJsonParse(raw) {
    try {
        return JSON.parse(String(raw));
    } catch {
        return null;
    }
}

function normalizePins(pins) {
    const arr = Array.isArray(pins) ? pins : [];
    const out = [];
    const seen = new Set();
    arr.forEach((p) => {
        const v = String(p == null ? '' : p).trim();
        if (!v) return;
        const key = v.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        out.push(v);
    });
    return out;
}

export function loadAccessOverride(storage = localStorage) {
    try {
        const raw = storage.getItem(ACCESS_OVERRIDE_KEY);
        if (!raw) return null;
        const parsed = safeJsonParse(raw);
        if (!parsed || typeof parsed !== 'object') return null;

        const enabled = typeof parsed.enabled === 'boolean' ? parsed.enabled : null;
        const pins = Array.isArray(parsed.pins) ? normalizePins(parsed.pins) : [];

        return { enabled: enabled, pins: pins };
    } catch {
        return null;
    }
}

export function saveAccessOverride({ enabled, pins } = {}, storage = localStorage) {
    const outPins = normalizePins(pins);
    const outEnabled = typeof enabled === 'boolean' ? enabled : true;
    const payload = { enabled: outEnabled, pins: outPins };
    try {
        storage.setItem(ACCESS_OVERRIDE_KEY, JSON.stringify(payload));
    } catch {
        // ignore
    }
    return payload;
}

export function getActiveUserAccessConfig(staticConfig, storage = localStorage) {
    const cfg = staticConfig && typeof staticConfig === 'object' ? staticConfig : {};
    const override = loadAccessOverride(storage);

    const enabled =
        override && typeof override.enabled === 'boolean' ? override.enabled : cfg.enabled !== false;

    const staticPins = Array.isArray(cfg.pins) ? normalizePins(cfg.pins) : [];
    const effectivePins = override && override.pins && override.pins.length ? override.pins : staticPins;

    return { enabled: enabled, pins: effectivePins };
}

