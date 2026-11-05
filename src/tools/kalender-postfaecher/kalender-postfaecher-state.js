/**
 * Local-first State für geplante Kalender-Postfächer.
 */

import { safeJsonParse } from '../../shared/utils/json.js';

export const STORAGE_KEY = 'ms365-kalender-postfaecher-v1';

/**
 * @returns {{ items: object[], domain: string, suggestedAdminUpn: string, inventory: object[] }}
 */
export function loadState() {
    const raw = typeof localStorage !== 'undefined' ? localStorage.getItem(STORAGE_KEY) : null;
    const parsed = safeJsonParse(raw, null);
    if (!parsed || typeof parsed !== 'object') {
        return { items: [], domain: '', suggestedAdminUpn: '', inventory: [] };
    }
    return {
        items: Array.isArray(parsed.items) ? parsed.items : [],
        domain: String(parsed.domain || '').trim(),
        suggestedAdminUpn: String(parsed.suggestedAdminUpn || '').trim(),
        inventory: Array.isArray(parsed.inventory) ? parsed.inventory : []
    };
}

/**
 * @param {{ items?: object[], domain?: string, suggestedAdminUpn?: string, inventory?: object[] }} state
 */
export function saveState(state) {
    const s = state && typeof state === 'object' ? state : {};
    const payload = {
        items: Array.isArray(s.items) ? s.items : [],
        domain: String(s.domain || '').trim(),
        suggestedAdminUpn: String(s.suggestedAdminUpn || '').trim(),
        inventory: Array.isArray(s.inventory) ? s.inventory : []
    };
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    } catch {
        /* ignore quota */
    }
    return payload;
}
