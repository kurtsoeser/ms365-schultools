/**
 * Dialog / norm / field helpers for tenant-settings-ui.
 */
function normStr(v) {
    try {
        if (typeof window !== 'undefined' && window.ms365Strings && typeof window.ms365Strings.normStr === 'function') {
            return window.ms365Strings.normStr(v);
        }
    } catch {
        /* fallback */
    }
    return String(v ?? '').trim();
}

function dlgAlert(msg, opts) {
    if (typeof window.ms365AppDialogAlert === 'function') {
        return window.ms365AppDialogAlert(msg, opts);
    }
    window.alert(msg);
    return Promise.resolve();
}

function dlgConfirm(msg, opts) {
    if (typeof window.ms365AppDialogConfirm === 'function') {
        return window.ms365AppDialogConfirm(msg, opts);
    }
    return Promise.resolve(window.confirm(msg));
}

function dlgPrompt(msg, def, opts) {
    if (typeof window.ms365AppDialogPrompt === 'function') {
        return window.ms365AppDialogPrompt(msg, def, opts);
    }
    return Promise.resolve(window.prompt(msg, def));
}

function normCode(v) {
    try {
        if (typeof window !== 'undefined' && window.ms365Strings && typeof window.ms365Strings.normCode === 'function') {
            return window.ms365Strings.normCode(v);
        }
    } catch {
        /* fallback */
    }
    return normStr(v).toUpperCase();
}

function escapeHtml(s) {
    try {
        if (typeof window !== 'undefined' && window.ms365Strings && typeof window.ms365Strings.escapeHtml === 'function') {
            return window.ms365Strings.escapeHtml(s);
        }
    } catch {
        /* fallback */
    }
    return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function safeJsonParse(s) {
    try {
        return JSON.parse(String(s));
    } catch {
        return null;
    }
}

function normHeaderKey(k) {
    return String(k ?? '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, '')
        .replace(/ä/g, 'ae')
        .replace(/ö/g, 'oe')
        .replace(/ü/g, 'ue')
        .replace(/ß/g, 'ss')
        .replace(/[^a-z0-9]/g, '');
}

function getField(row, candidates) {
    if (!row || typeof row !== 'object') return '';
    const map = new Map();
    Object.keys(row).forEach((k) => map.set(normHeaderKey(k), row[k]));
    for (const c of candidates) {
        const v = map.get(normHeaderKey(c));
        if (v != null && String(v).trim() !== '') return String(v).trim();
    }
    return '';
}

export {
    normStr,
    dlgAlert,
    dlgConfirm,
    dlgPrompt,
    normCode,
    escapeHtml,
    safeJsonParse,
    normHeaderKey,
    getField
};
